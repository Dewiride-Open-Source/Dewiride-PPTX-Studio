/**
 * Experiment T13, step 2 - score every candidate reading against the browser.
 *
 * ```
 * node tools/ground-truth/fonts/metrics/analyse.ts <work-dir> [--fixture <path>]
 * ```
 *
 * Nothing here is allowed to conclude by inspection. Each question below is a
 * set of rival readings, every one is scored against every measurement, and the
 * script throws rather than emit a fixture where two of them tie. ADR 0042.
 */

import { readFileSync, writeFileSync } from 'node:fs';
import { join, resolve } from 'node:path';

// The built package, not its source: `tools/` is run by Node directly and the
// sources carry `.js` specifiers. Run `pnpm build` first.
import { facesIn } from '../../../../packages/cli/dist/index.js';
import { BOX_ROUNDINGS, boxRoundingNamed } from '../../lib/box-rounding.ts';
import { buildFont, type Baselines, type FontSpec, type KernPair } from '../../lib/truetype.ts';
import { PROBES, STRINGS, SIZES, BOX_PX, BASELINE_SIZES, probeFonts } from './probes.ts';

interface BrowserBaselines {
  readonly alphabetic: number;
  readonly hanging: number;
  readonly ideographic: number;
}

interface Measured {
  readonly chromium: string;
  readonly boxPx: number;
  readonly sizes: readonly number[];
  readonly baselineSizes: readonly number[];
  readonly strings: readonly string[];
  readonly measured: readonly {
    readonly id: string;
    readonly widths: Record<string, Record<string, number>>;
    readonly box: { readonly ascent: number; readonly descent: number };
    readonly baselines: Record<string, BrowserBaselines>;
    readonly hanBaselines: Record<string, BrowserBaselines & { width: number }> | null;
  }[];
}

interface Scored {
  readonly model: string;
  readonly fits: number;
  readonly of: number;
  /** The largest absolute disagreement, in the unit of the question. */
  readonly worst: number;
}

/** Refuse a question whose winner is not alone at the top. */
function decide<T extends Scored>(question: string, scored: readonly T[]): T {
  const ranked = [...scored].sort((a, b) => b.fits - a.fits || a.worst - b.worst);
  const best = ranked[0];
  if (best === undefined) throw new Error(`${question}: nothing was scored`);
  if (best.fits !== best.of) {
    throw new Error(
      `${question}: the best reading "${best.model}" fits only ${String(best.fits)}/${String(best.of)}`,
    );
  }
  const rival = ranked[1];
  if (rival !== undefined && rival.fits === best.fits) {
    throw new Error(
      `${question}: "${best.model}" and "${rival.model}" both fit ${String(best.fits)}/${String(best.of)}; ` +
        'no probe separates them, so neither may be committed as the answer',
    );
  }
  return best;
}

const work = resolve(process.argv[2] ?? '');
if (process.argv[2] === undefined) throw new Error('usage: analyse.ts <work-dir> [--fixture <p>]');
const fixtureAt = process.argv.indexOf('--fixture');
const fixture = fixtureAt < 0 ? null : process.argv[fixtureAt + 1];

const browser = JSON.parse(readFileSync(join(work, 'browser-metrics.json'), 'utf8')) as Measured;

const rowOf = (id: string): Measured['measured'][number] => {
  const row = browser.measured.find((r) => r.id === id);
  if (row === undefined) throw new Error(`no browser row for ${id}`);
  return row;
};

/* -------------------------------------------------------------------------- */
/* question 1 - which table the font bounding box comes from, rounded how      */
/* -------------------------------------------------------------------------- */

const BOX_READINGS = {
  'hhea.ascender/descender': (s: ReturnType<typeof buildFont>['spec']) => ({
    ascent: s.hheaAscender,
    descent: -s.hheaDescender,
  }),
  'OS/2.usWinAscent/usWinDescent': (s: ReturnType<typeof buildFont>['spec']) => ({
    ascent: s.winAscent,
    descent: s.winDescent,
  }),
  'OS/2.sTypoAscender/sTypoDescender': (s: ReturnType<typeof buildFont>['spec']) => ({
    ascent: s.typoAscender,
    descent: -s.typoDescender,
  }),
  'usWin, or sTypo when fsSelection bit 7 is set': (s: ReturnType<typeof buildFont>['spec']) =>
    s.useTypoMetrics === true
      ? { ascent: s.typoAscender, descent: -s.typoDescender }
      : { ascent: s.winAscent, descent: s.winDescent },
};

interface BoxScored extends Scored {
  readonly table: string;
  readonly rounding: string;
  /** The pair the table names, so a later question reads the box the way this one settled. */
  readonly read: (s: FontSpec) => { ascent: number; descent: number };
  /** Every `id:ascent|descent` the model missed, so a table can say what separated it. */
  readonly missed: readonly string[];
}

// Table and rounding are scored as one model: the 1000 em probes tie every
// rounding, and only the three fractional probes separate them.
const boxScores: BoxScored[] = Object.entries(BOX_READINGS).flatMap(([table, read]) =>
  Object.entries(BOX_ROUNDINGS).map(([rounding, round]): BoxScored => {
    let fits = 0;
    let of = 0;
    let worst = 0;
    const missed: string[] = [];
    for (const probe of PROBES) {
      const { spec } = buildFont(probe.spec);
      const want = round(read(spec), spec.unitsPerEm, BOX_PX);
      const got = rowOf(probe.id).box;
      for (const [side, predicted, measured] of [
        ['ascent', want.ascent, got.ascent],
        ['descent', want.descent, got.descent],
      ] as const) {
        of += 1;
        const off = Math.abs(predicted - measured);
        if (off < 1e-9) fits += 1;
        else {
          worst = Math.max(worst, off);
          missed.push(`${probe.id}:${side}`);
        }
      }
    }
    return { model: `${table}; ${rounding}`, table, rounding, read, fits, of, worst, missed };
  }),
);
const boxAnswer = decide('the font bounding box', boxScores);
const boxRounding = boxRoundingNamed(boxAnswer.rounding);

/* -------------------------------------------------------------------------- */
/* question 2 - which kerning table the browser honours                        */
/* -------------------------------------------------------------------------- */

const specOf = (id: string): FontSpec =>
  buildFont(PROBES.find((p) => p.id === id)?.spec ?? {}).spec;

/** The `A`/`B` adjustment a table declares, in font units. */
const declared = (pairs: readonly KernPair[] | undefined): number =>
  pairs?.find((p) => p.left === 'A' && p.right === 'B')?.adjust ?? 0;

/** What GPOS yields a reader that does or does not follow a type 9 Extension. */
const fromGpos = (s: FontSpec, extensions: boolean): number | undefined =>
  s.gpos === undefined || (s.gposExtension === true && !extensions) ? undefined : declared(s.gpos);

/** What each probe's tables say, so a reading can be scored without the reader. */
const KERN_READINGS: Record<string, (s: FontSpec) => number> = {
  'the legacy kern table only': (s) => declared(s.kern),
  'GPOS PairPos only, type 2 lookups only': (s) => fromGpos(s, false) ?? 0,
  'GPOS PairPos only, following one type 9 Extension lookup': (s) => fromGpos(s, true) ?? 0,
  'GPOS when it yields a pair, else the kern table, type 2 lookups only': (s) =>
    fromGpos(s, false) ?? declared(s.kern),
  'GPOS when it yields a pair, else the kern table, following one type 9 Extension lookup': (s) =>
    fromGpos(s, true) ?? declared(s.kern),
  'the kern table when present, else GPOS': (s) =>
    s.kern === undefined ? (fromGpos(s, true) ?? 0) : declared(s.kern),
  neither: () => 0,
};

const kernScores: Scored[] = Object.entries(KERN_READINGS).map(([model, adjustOf]) => {
  let fits = 0;
  let of = 0;
  let worst = 0;
  for (const probe of PROBES) {
    const row = rowOf(probe.id);
    const spec = specOf(probe.id);
    const adjust = adjustOf(spec);
    for (const px of SIZES) {
      // `AB` is the kerned pair; `AC` is the control that never moves.
      const kerned = row.widths[String(px)]?.['AB'] ?? 0;
      const control = row.widths[String(px)]?.['AC'] ?? 0;
      const predicted = control + (adjust * px) / spec.unitsPerEm;
      of += 1;
      // The rival readings are 100 font units apart, which is 0.8px at the
      // smallest size probed; the browser quantisation this tolerance absorbs is
      // three orders of magnitude smaller, and is question 3.
      const off = Math.abs(predicted - kerned);
      if (off < 1e-3) fits += 1;
      else worst = Math.max(worst, off);
    }
  }
  return { model, fits, of, worst };
});
const kernAnswer = decide('kerning', kernScores);

/* -------------------------------------------------------------------------- */
/* question 3 - the arithmetic that turns font units into a width              */
/* -------------------------------------------------------------------------- */

const FIXED = 65536;
const QUANTISERS = {
  'exact float, no quantisation': { advance: (x: number) => x, kern: (x: number) => x },
  'advance and kern truncated at 1/65536 px': {
    advance: (x: number) => Math.trunc(x * FIXED) / FIXED,
    kern: (x: number) => Math.trunc(x * FIXED) / FIXED,
  },
  'advance and kern rounded at 1/65536 px': {
    advance: (x: number) => Math.round(x * FIXED) / FIXED,
    kern: (x: number) => Math.round(x * FIXED) / FIXED,
  },
  'advance truncated, kern rounded, both at 1/65536 px': {
    advance: (x: number) => Math.trunc(x * FIXED) / FIXED,
    kern: (x: number) => Math.round(x * FIXED) / FIXED,
  },
  'the whole string quantised once at 1/65536 px': {
    advance: (x: number) => x,
    kern: (x: number) => x,
  },
};

/** The reader's own answer for one string, under a candidate quantisation. */
function widthUnder(
  faceFor: string,
  text: string,
  px: number,
  q: { advance: (x: number) => number; kern: (x: number) => number },
): number {
  const built = buildFont(PROBES.find((p) => p.id === faceFor)?.spec ?? {});
  // T13 was measured in a Chromium reading these tables through DirectWrite.
  const face = facesIn(built.bytes, faceFor, 'directwrite')[0]!;
  const em = face.metrics.unitsPerEm;
  const points = [...text];
  let width = 0;
  for (let i = 0; i < points.length; i++) {
    const code = points[i]!.codePointAt(0)!;
    width += q.advance(((face.advanceOf(code) ?? 0) * px) / em);
    if (i + 1 < points.length) {
      const adjust = face.kernBetween(code, points[i + 1]!.codePointAt(0)!);
      if (adjust !== 0) width += q.kern((adjust * px) / em);
    }
  }
  return width;
}

const widthScores: Scored[] = Object.entries(QUANTISERS).map(([model, q]) => {
  let fits = 0;
  let of = 0;
  let worst = 0;
  for (const probe of PROBES) {
    const row = rowOf(probe.id);
    for (const px of SIZES) {
      for (const text of STRINGS) {
        const raw = widthUnder(probe.id, text, px, q);
        const predicted =
          model === 'the whole string quantised once at 1/65536 px'
            ? Math.trunc(raw * FIXED) / FIXED
            : raw;
        const measured = row.widths[String(px)]?.[text] ?? 0;
        of += 1;
        const off = Math.abs(predicted - measured);
        if (off < 1e-12) fits += 1;
        else worst = Math.max(worst, off);
      }
    }
  }
  return { model, fits, of, worst };
});
const widthAnswer = decide('the width arithmetic', widthScores);

/* -------------------------------------------------------------------------- */
/* question 4 - where the ideographic baseline comes from                      */
/* -------------------------------------------------------------------------- */

/** The script a run is in, which a reading may or may not turn out to consult. */
type Script = 'latn' | 'hani';

/** The face box as question 1 settled it, table and all, in font units. */
const boxOf = boxAnswer.read;

const toPx = (units: number, s: FontSpec, px: number): number => (units * px) / s.unitsPerEm;

/** Pixels below the alphabetic baseline at `px`, positive downwards. */
type Fallback = (s: FontSpec, px: number) => number;

/** A reading in font units, which every table reading is. */
const inUnits =
  (units: (s: FontSpec) => number): Fallback =>
  (s, px) =>
    toPx(units(s), s, px);

const boxDescent = inUnits((s) => boxOf(s).descent);

/** The descent of the box as the browser reported it, whole pixels included. */
const roundedBoxDescent: Fallback = (s, px) => boxRounding(boxOf(s), s.unitsPerEm, px).descent;

const FALLBACKS: Record<string, Fallback> = {
  'the face box descent': boxDescent,
  'the face box descent, rounded as the box is': roundedBoxDescent,
  'OS/2.usWinDescent': inUnits((s) => s.winDescent),
  'OS/2.sTypoDescender': inUnits((s) => -s.typoDescender),
  'hhea.descender': inUnits((s) => -s.hheaDescender),
  'sTypoDescender where fsSelection bit 7 is set and hhea.descender otherwise': inUnits((s) =>
    s.useTypoMetrics === true ? -s.typoDescender : -s.hheaDescender,
  ),
};

/** Which `BaseScript` a reading looks the coordinate up in. */
const LOOKUPS: Record<string, (s: FontSpec, script: Script) => Baselines | undefined> = {
  'the run script': (s, script) => s.base?.[script],
  'the DFLT script': (s) => s.base?.DFLT,
  'the latn script': (s) => s.base?.latn,
  'the hani script': (s) => s.base?.hani,
  'the DFLT script or the latn script': (s) => s.base?.DFLT ?? s.base?.latn,
  'the DFLT script or the first script named': (s) =>
    s.base?.DFLT ?? s.base?.hani ?? s.base?.kana ?? s.base?.latn,
};

/** Pixels below the alphabetic baseline at `px`, positive downwards. */
type Reading = (s: FontSpec, script: Script, px: number) => number;

const IDEOGRAPHIC_READINGS: Record<string, Reading> = {
  'the face box descent': (s, _script, px) => boxDescent(s, px),
  'the face box descent, rounded as the box is': (s, _script, px) => roundedBoxDescent(s, px),
  'OS/2.usWinDescent': (s, _script, px) => toPx(s.winDescent, s, px),
  'OS/2.sTypoDescender': (s, _script, px) => toPx(-s.typoDescender, s, px),
  'hhea.descender': (s, _script, px) => toPx(-s.hheaDescender, s, px),
  'zero, the alphabetic baseline': () => 0,
};

for (const [where, lookup] of Object.entries(LOOKUPS)) {
  for (const [tag, fallbacks] of [
    ['ideo', where === 'the run script' || where === 'the DFLT script' ? FALLBACKS : {}],
    ['icfb', {}],
  ] as const) {
    const only = { 'the face box descent': boxDescent };
    for (const [otherwise, fallback] of Object.entries({ ...only, ...fallbacks })) {
      IDEOGRAPHIC_READINGS[`the BASE ${tag} coordinate of ${where}, else ${otherwise}`] = (
        s,
        script,
        px,
      ) => {
        const coordinate = lookup(s, script)?.[tag];
        return coordinate === undefined ? fallback(s, px) : toPx(-coordinate, s, px);
      };
    }
  }
}

// A coordinate outside the em is where `packages/text`'s own probe gives up and
// answers zero, so whether the browser hands it over intact decides whether the
// two engines can agree on such a face at all.
IDEOGRAPHIC_READINGS[
  'the BASE ideo coordinate of the DFLT script where it falls inside the em, else the face box descent'
] = (s, _script, px) => {
  const coordinate = s.base?.DFLT?.ideo;
  const units = coordinate === undefined ? undefined : -coordinate;
  return units === undefined || units <= 0 || units >= s.unitsPerEm
    ? boxDescent(s, px)
    : toPx(units, s, px);
};

interface Observation {
  readonly id: string;
  readonly px: number;
  readonly script: Script;
  readonly measured: number;
}

const observations: Observation[] = [];
for (const row of browser.measured) {
  for (const px of BASELINE_SIZES) {
    const latin = row.baselines[String(px)];
    if (latin === undefined) throw new Error(`${row.id}: no baselines at ${String(px)}px`);
    observations.push({ id: row.id, px, script: 'latn', measured: -latin.ideographic });
    const han = row.hanBaselines?.[String(px)];
    if (han !== undefined) {
      observations.push({ id: row.id, px, script: 'hani', measured: -han.ideographic });
    }
  }
}

/* Guard 1: a non-zero alphabetic baseline means the browser answered about some
   other face, or about a convention we have misread, and nothing here is
   evidence. */
for (const row of browser.measured) {
  for (const [px, read] of [
    ...Object.entries(row.baselines),
    ...Object.entries(row.hanBaselines ?? {}),
  ]) {
    if (Math.abs(read.alphabetic) > 1e-9) {
      throw new Error(
        `${row.id} at ${px}px reported an alphabetic baseline of ${String(read.alphabetic)}, ` +
          'which is not the origin every other baseline here is measured from',
      );
    }
  }
}

/* Guard 2: the ideograph must have been drawn in the probe, not in a face off
   this machine, and its own advance is what says so. */
for (const row of browser.measured) {
  if (row.hanBaselines === null) continue;
  const spec = specOf(row.id);
  for (const [px, read] of Object.entries(row.hanBaselines)) {
    const want = (spec.advance * Number(px)) / spec.unitsPerEm;
    if (Math.abs(read.width - want) > 1e-9) {
      throw new Error(
        `${row.id} drew U+${(spec.ideograph ?? 0).toString(16)} ${String(read.width)}px wide at ` +
          `${px}px where its own advance is ${String(want)}px, so some other face answered`,
      );
    }
  }
}

const ideographicScores: Scored[] = Object.entries(IDEOGRAPHIC_READINGS).map(([model, read]) => {
  let fits = 0;
  let worst = 0;
  for (const at of observations) {
    const spec = specOf(at.id);
    const predicted = read(spec, at.script, at.px);
    const off = Math.abs(predicted - at.measured);
    if (off < 1e-9) fits += 1;
    else worst = Math.max(worst, off);
  }
  return { model, fits, of: observations.length, worst };
});
const ideographicAnswer = decide('the ideographic baseline', ideographicScores);

/* Guard 3: a null result and a malformed BASE table look identical, because
   HarfBuzz sanitises a bad one and reads it as absent. The `hang` coordinate of
   the same table is what tells them apart. */
const hangingHonoured = browser.measured
  .filter((row) => specOf(row.id).base !== undefined)
  .map((row) => {
    const spec = specOf(row.id);
    const hang = spec.base?.DFLT?.hang;
    const read = row.baselines[String(BOX_PX)];
    const want = ((hang ?? 0) * BOX_PX) / spec.unitsPerEm;
    return {
      id: row.id,
      honoured: hang !== undefined && Math.abs((read?.hanging ?? 0) - want) < 1e-9,
    };
  });
if (!ideographicAnswer.model.includes('BASE') && !hangingHonoured.some((row) => row.honoured)) {
  throw new Error(
    `the ideographic baseline scored as "${ideographicAnswer.model}", but no BASE font's hang ` +
      'coordinate came back through hangingBaseline either, so "the browser ignores BASE" and ' +
      '"the probe emitted a BASE table HarfBuzz sanitised away" are indistinguishable; the ' +
      'experiment is inconclusive and must not be committed',
  );
}

/* -------------------------------------------------------------------------- */
/* the fixture                                                                */
/* -------------------------------------------------------------------------- */

const linear = widthScores.find((s) => s.model === 'exact float, no quantisation');

const answers = {
  $comment:
    'Experiment T13: what a font-table reader must do to agree with the browser it is standing ' +
    'in for. Regenerate with tools/ground-truth/fonts/metrics/analyse.ts. The fonts are built by ' +
    'tools/ground-truth/lib/truetype.ts and exist nowhere else, so no installed face can answer ' +
    'for a probe.',
  experiment: 'T13',
  subPhase: '3.10',
  adr: 'docs/adr/phase-3-text/0042-rendering-without-a-browser.md',
  chromium: browser.chromium,
  probes: PROBES.map((p) => ({ id: p.id, asks: p.asks })),
  // The bytes measured, and what the browser said about them. Committed so the
  // suite re-derives every rule from the measurement rather than from the code
  // that implements it - a test written off the implementation passes forever.
  fonts: probeFonts().map((font) => ({
    id: font.id,
    family: font.family,
    spec: font.spec,
    base64: font.base64,
  })),
  browser: {
    sizes: SIZES,
    strings: STRINGS,
    boxPx: BOX_PX,
    baselineSizes: BASELINE_SIZES,
    rows: browser.measured.map((row) => ({
      id: row.id,
      box: row.box,
      widths: row.widths,
      baselines: row.baselines,
      hanBaselines: row.hanBaselines,
    })),
  },
  measurements: {
    widths: PROBES.length * SIZES.length * STRINGS.length,
    boxes: PROBES.length * 2,
    baselines: observations.length,
  },
  faceBox: {
    $comment:
      'The ascent and descent a browser reports for a face. Every real font sets hhea and usWin ' +
      'to the same numbers, so only a font built to disagree can say which was read; and every ' +
      'probe on a 1000 em lands on a whole pixel, so only one that does not can say how the ' +
      'browser rounds it.',
    answer: boxAnswer.table,
    rounding: boxAnswer.rounding,
    scores: boxScores,
  },
  kerning: {
    $comment:
      'A font may carry pair adjustments in a legacy kern table, in GPOS, or in both. The third ' +
      'probe sets them to different values, which is the only way to find out which one loses, ' +
      'and the fourth reaches its PairPos through a type 9 Extension lookup, which is how a font ' +
      'compiler emits one and where a reader that skips the type silently loses all kerning.',
    answer: kernAnswer.model,
    scores: kernScores,
  },
  width: {
    $comment:
      'Whether an advance is linear in size - that is, unhinted - and how the browser quantises ' +
      'it. The linear reading is what matters on a slide; the quantisation is what lets the ' +
      "reader's test assert equality instead of a tolerance.",
    answer: widthAnswer.model,
    linearWithinPx: linear?.worst ?? null,
    scores: widthScores,
  },
  ideographic: {
    $comment:
      'Where an upright East Asian glyph puts its baseline. A BASE table names a coordinate per ' +
      'script per baseline tag, so the reading has to say which tag, which script, and what a ' +
      'font without one falls back to; three fonts disagree with themselves on all three.',
    answer: ideographicAnswer.model,
    hangingHonoured,
    scores: ideographicScores,
  },
} as const;

const text = `${JSON.stringify(answers, null, 2)}\n`;
writeFileSync(join(work, 't13-answers.json'), text);
if (fixture !== undefined && fixture !== null) writeFileSync(fixture, text);

console.log(`chromium ${browser.chromium}`);
console.log(`face box   ${boxAnswer.model}  (${String(boxAnswer.fits)}/${String(boxAnswer.of)})`);
console.log(
  `kerning    ${kernAnswer.model}  (${String(kernAnswer.fits)}/${String(kernAnswer.of)})`,
);
console.log(
  `width      ${widthAnswer.model}  (${String(widthAnswer.fits)}/${String(widthAnswer.of)})`,
);
console.log(`  linear reading is within ${String(linear?.worst ?? 0)} px`);
console.log(
  `ideographic ${ideographicAnswer.model}  ` +
    `(${String(ideographicAnswer.fits)}/${String(ideographicAnswer.of)})`,
);
for (const row of hangingHonoured) {
  console.log(`  ${row.id} hang coordinate honoured: ${String(row.honoured)}`);
}
for (const [name, scores] of [
  ['face box', boxScores],
  ['kerning', kernScores],
  ['width', widthScores],
  ['ideographic', ideographicScores],
] as const) {
  console.log(`\n${name}`);
  for (const s of [...scores].sort((a, b) => b.fits - a.fits)) {
    console.log(`  ${String(s.fits).padStart(4)}/${String(s.of).padEnd(5)} ${s.model}`);
  }
}
