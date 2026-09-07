/**
 * Experiment T7, step 4 - score every candidate reading, then write the fixture.
 *
 * ```
 * node tools/ground-truth/fonts/substitution/analyse.ts <work-dir>
 * ```
 *
 * Throws rather than emitting a fixture it cannot fit perfectly. ADR 0033.
 */

import { readFileSync, writeFileSync } from 'node:fs';
import { format, resolveConfig } from 'prettier';
import { join, resolve } from 'node:path';

import { readEmf } from '../../lib/emf.ts';
import { repoPath } from '../../../repo/root.ts';
import {
  ABSENT_FACES,
  ALIAS_FACES,
  INSTALLED_FACES,
  WIDTH_FACES,
  WIDTH_SIZES,
  WIDTH_TEXT,
  BASE_ABSENT,
  type Face,
  type Slot,
} from './probes.ts';

const arg = process.argv[2];
if (arg === undefined) throw new Error('usage: analyse.ts <work-dir>');
const work: string = resolve(arg);

/* -------------------------------------------------------------------------- */
/* the readings                                                               */
/* -------------------------------------------------------------------------- */

interface ProbeSpec {
  readonly id: string;
  readonly kind: string;
  readonly slide: number;
  readonly sz: number;
  readonly text: string;
  readonly faces: Readonly<Partial<Record<Slot, Face>>>;
  readonly control?: string;
}
interface DeckSpec {
  readonly key: string;
  readonly file: string;
  readonly themeFonts: Record<string, string>;
  readonly probes: readonly ProbeSpec[];
}
interface ShapeReading {
  readonly fontName: string | null;
  readonly fontNameFE: string | null;
  readonly fontNameC: string | null;
  readonly boundWidth: number | null;
  readonly boundHeight: number | null;
}
interface SlideReading {
  readonly slide: number;
  readonly shape: ShapeReading | null;
  readonly emf: string | null;
}
interface DeckReading {
  readonly key: string;
  readonly opened: boolean;
  readonly repaired: boolean | null;
  readonly slides: readonly SlideReading[];
}

const inputs = JSON.parse(readFileSync(join(work, 'substitution-inputs.json'), 'utf8')) as {
  readonly decks: readonly DeckSpec[];
};
const readings = JSON.parse(readFileSync(join(work, 'substitution-readings.json'), 'utf8')) as {
  readonly installedFamilies: readonly string[];
  readonly decks: readonly DeckReading[];
};

/** One probe, its specification and every reading taken of it. */
interface Row {
  readonly deck: string;
  readonly probe: ProbeSpec;
  readonly themeFonts: Record<string, string>;
  /** The `LOGFONTW` faces PowerPoint selected, in draw order, without repeats. */
  readonly drawn: readonly string[];
  readonly comName: string | null;
  readonly width: number;
}

const rows: Row[] = [];
for (const deck of readings.decks) {
  if (!deck.opened || deck.repaired === true) {
    throw new Error(`package ${deck.key} did not open cleanly; nothing here is a measurement`);
  }
  const spec = inputs.decks.find((d) => d.key === deck.key);
  if (spec === undefined) throw new Error(`no specification for package ${deck.key}`);
  for (const slide of deck.slides) {
    const probe = spec.probes.find((p) => p.slide === slide.slide);
    if (probe === undefined) throw new Error(`${deck.key} slide ${String(slide.slide)}: no probe`);
    if (slide.emf === null) throw new Error(`${probe.id}: no EMF`);
    const emf = readEmf(new Uint8Array(readFileSync(join(work, slide.emf))));
    const drawn: string[] = [];
    for (const text of emf.texts) {
      if (text.text.trim().length === 0) continue;
      const face = text.font?.face;
      if (face !== undefined && face.length > 0 && !drawn.includes(face)) drawn.push(face);
    }
    const width = slide.shape?.boundWidth;
    if (width === null || width === undefined) throw new Error(`${probe.id}: no BoundWidth`);
    rows.push({
      deck: deck.key,
      probe,
      themeFonts: spec.themeFonts,
      drawn,
      comName: slide.shape?.fontName ?? null,
      width,
    });
  }
}

const rowOf = (deck: string, id: string): Row => {
  const found = rows.find((r) => r.deck === deck && r.probe.id === id);
  if (found === undefined) throw new Error(`no row ${deck}/${id}`);
  return found;
};
const only = (row: Row): string => {
  if (row.drawn.length !== 1) {
    throw new Error(
      `${row.probe.id} drew ${String(row.drawn.length)} faces: ${row.drawn.join(', ')}`,
    );
  }
  return row.drawn[0] as string;
};

/* -------------------------------------------------------------------------- */
/* scoring                                                                    */
/* -------------------------------------------------------------------------- */

interface Candidate<T> {
  readonly name: string;
  readonly predict: (row: T) => string;
}
interface Score {
  readonly name: string;
  readonly fits: number;
  readonly total: number;
}
interface Question {
  readonly key: string;
  readonly question: string;
  readonly answer: string;
  readonly scores: readonly Score[];
}

const questions: Question[] = [];

/**
 * Score every candidate and keep the one that fits every row.
 *
 * Throws on none and on more than one: a question two readings both explain is
 * a question whose probes do not separate them.
 */
function choose<T>(
  key: string,
  question: string,
  answer: string,
  items: readonly T[],
  observe: (item: T) => string,
  candidates: readonly Candidate<T>[],
): void {
  if (items.length === 0) throw new Error(`${key}: no rows`);
  const scores = candidates.map((candidate) => ({
    name: candidate.name,
    fits: items.filter((item) => candidate.predict(item) === observe(item)).length,
    total: items.length,
  }));
  const perfect = scores.filter((s) => s.fits === s.total);
  if (perfect.length !== 1) {
    const table = scores.map((s) => `  ${s.name}: ${String(s.fits)}/${String(s.total)}`).join('\n');
    throw new Error(
      `${key}: ${String(perfect.length)} candidate(s) fit all ${String(items.length)} rows\n${table}`,
    );
  }
  questions.push({ key, question, answer, scores: [...scores].sort((a, b) => b.fits - a.fits) });
}

/* -------------------------------------------------------------------------- */
/* what an absent Latin face becomes                                          */
/* -------------------------------------------------------------------------- */

const absentRows = [
  ...ABSENT_FACES.map((_, i) => rowOf('faces', `absent-${String(i).padStart(2, '0')}`)),
  rowOf('theme-absent', 'theme-major-latin'),
  rowOf('theme-absent', 'theme-minor-latin'),
];

/** A reading anyone writes first: the style word in the name picks the family. */
const byName = (name: string): string => {
  const lower = name.toLowerCase();
  if (lower.includes('mono')) return 'Courier New';
  if (lower.includes('serif') || lower.includes('roman')) return 'Times New Roman';
  return 'Calibri';
};

choose(
  'absent-latin',
  'What does PowerPoint draw a Latin face it does not have in?',
  'Calibri, whatever the name says.',
  absentRows,
  only,
  [
    { name: 'always Calibri', predict: () => 'Calibri' },
    {
      name: 'the style word in the name',
      predict: (r) => byName(r.probe.faces.latin?.typeface ?? ''),
    },
    { name: "the theme's minor Latin face", predict: (r) => r.themeFonts['minorLatin'] ?? '' },
    { name: 'always Arial', predict: () => 'Arial' },
  ],
);

/* -------------------------------------------------------------------------- */
/* the three hints on CT_TextFont                                             */
/* -------------------------------------------------------------------------- */

const hintRows = rows.filter((r) => r.deck === 'hints');

/** The charsets that moved the answer, as measured. Data, not a rule. */
const CHARSET_SUBSTITUTE: Readonly<Record<string, string>> = Object.fromEntries(
  hintRows
    .filter((r) => r.probe.id.startsWith('hint-charset-'))
    .map((r) => [String(r.probe.faces.latin?.charset ?? 0), only(r)]),
);

choose(
  'hints',
  'Do @panose, @pitchFamily and @charset change what an absent face becomes?',
  'Only @charset does. @panose and @pitchFamily are ignored.',
  hintRows,
  only,
  [
    { name: 'all three ignored', predict: () => 'Calibri' },
    {
      name: 'only @charset is read',
      predict: (r) => CHARSET_SUBSTITUTE[String(r.probe.faces.latin?.charset ?? 0)] ?? 'Calibri',
    },
    {
      name: '@pitchFamily picks the family',
      predict: (r) => {
        const family = (r.probe.faces.latin?.pitchFamily ?? 0) >> 4;
        if (family === 1) return 'Times New Roman';
        if (family === 2) return 'Arial';
        if (family === 3) return 'Courier New';
        return 'Calibri';
      },
    },
    {
      name: '@panose picks the family',
      predict: (r) => {
        const panose = r.probe.faces.latin?.panose ?? '';
        if (panose.startsWith('0202')) return 'Times New Roman';
        if (panose.startsWith('020B')) return 'Arial';
        if (panose.startsWith('0207')) return 'Courier New';
        return 'Calibri';
      },
    },
  ],
);

/* -------------------------------------------------------------------------- */
/* is the substitution metric-preserving?                                     */
/* -------------------------------------------------------------------------- */

interface WidthRow {
  readonly sz: number;
  readonly textIndex: number;
  readonly absent: number;
  readonly byFace: Readonly<Record<string, number>>;
}

const widthRows: WidthRow[] = [];
for (const sz of WIDTH_SIZES) {
  for (const [textIndex] of WIDTH_TEXT.entries()) {
    const at = (typeface: string): number =>
      rowOf('widths', `width-${typeface.replace(/\s+/g, '')}-${String(sz)}-${String(textIndex)}`)
        .width;
    widthRows.push({
      sz,
      textIndex,
      absent: at(BASE_ABSENT),
      byFace: Object.fromEntries(
        WIDTH_FACES.filter((f) => f !== BASE_ABSENT).map((f) => [f, at(f)]),
      ),
    });
  }
}

choose(
  'substitute-metrics',
  'Is an absent face laid out with the substitute own metrics, or scaled?',
  'The substitute own, exactly: every reading equals Calibri to the last digit.',
  widthRows,
  (r) => r.absent.toFixed(6),
  WIDTH_FACES.filter((f) => f !== BASE_ABSENT).map((face) => ({
    name: `exactly ${face}`,
    predict: (r: WidthRow) => (r.byFace[face] ?? Number.NaN).toFixed(6),
  })),
);

/* -------------------------------------------------------------------------- */
/* the theme reference                                                        */
/* -------------------------------------------------------------------------- */

const themeDecks = ['theme-present', 'theme-empty', 'theme-empty-alt', 'theme-absent'];
const themeLatinRows = themeDecks.flatMap((deck) => [
  rowOf(deck, 'theme-major-latin'),
  rowOf(deck, 'theme-minor-latin'),
]);

/** What the face resolves to once absence has been accounted for. */
const afterSubstitution = (name: string): string => (name === BASE_ABSENT ? 'Calibri' : name);

choose(
  'theme-latin-ref',
  'Which collection do +mj-lt and +mn-lt name?',
  '+mj-lt is a:majorFont/a:latin and +mn-lt is a:minorFont/a:latin, substitution applied after.',
  themeLatinRows,
  only,
  [
    {
      name: 'major to majorFont, minor to minorFont',
      predict: (r) =>
        afterSubstitution(
          (r.probe.id === 'theme-major-latin'
            ? r.themeFonts['majorLatin']
            : r.themeFonts['minorLatin']) ?? '',
        ),
    },
    {
      name: 'the two collections swapped',
      predict: (r) =>
        afterSubstitution(
          (r.probe.id === 'theme-major-latin'
            ? r.themeFonts['minorLatin']
            : r.themeFonts['majorLatin']) ?? '',
        ),
    },
    {
      name: 'always the minor collection',
      predict: (r) => afterSubstitution(r.themeFonts['minorLatin'] ?? ''),
    },
  ],
);

/* -------------------------------------------------------------------------- */
/* what the object model reports                                              */
/* -------------------------------------------------------------------------- */

const comRows = rows.filter((r) => r.comName !== null && r.drawn.length === 1);

choose(
  'com-reports',
  'Does TextRange2.Font.Name report the requested face or the drawn one?',
  'The requested face, after theme resolution and before substitution.',
  comRows,
  (r) => String(r.comName),
  [
    {
      name: 'the requested face, theme resolved',
      predict: (r) => {
        const typeface = r.probe.faces.latin?.typeface ?? '';
        if (typeface === '+mj-lt') return r.themeFonts['majorLatin'] ?? '';
        if (typeface === '+mn-lt') return r.themeFonts['minorLatin'] ?? '';
        return typeface;
      },
    },
    { name: 'the face actually drawn', predict: only },
  ],
);

/* -------------------------------------------------------------------------- */
/* the browser                                                                */
/* -------------------------------------------------------------------------- */

interface BrowserFamily {
  readonly family: string;
  readonly check: boolean;
  readonly bare: Record<string, number[]>;
  readonly stacked: Record<string, Record<string, number[]>>;
}
interface BrowserFile {
  readonly chromium: string;
  readonly sizes: readonly number[];
  readonly glyphs: readonly string[];
  readonly anchors: readonly string[];
  readonly anchorOnly: Record<string, Record<string, number[]>>;
  readonly families: readonly BrowserFamily[];
}
const browser = JSON.parse(readFileSync(join(work, 'browser-fonts.json'), 'utf8')) as BrowserFile;

const PX = String(browser.sizes[0] ?? 100);
const absentSet = new Set<string>(ABSENT_FACES);
const at = (record: Record<string, number[]>): number[] => {
  const found = record[PX];
  if (found === undefined) throw new Error(`no widths at ${PX}px`);
  return found;
};
const same = (a: readonly number[], b: readonly number[]): boolean =>
  a.length === b.length && a.every((v, i) => v === b[i]);

const anchorEqual = (family: BrowserFamily, anchor: string): boolean => {
  const stacked = family.stacked[anchor];
  const alone = browser.anchorOnly[anchor];
  if (stacked === undefined || alone === undefined) throw new Error(`no anchor ${anchor}`);
  return same(at(stacked), at(alone));
};

choose(
  'detector',
  'How is an absent family detected in the browser?',
  'By measuring it stacked on all three generics: absent means equal to every one of them.',
  browser.families,
  (f) => (absentSet.has(f.family) ? 'absent' : 'present'),
  [
    {
      name: 'equal to all three anchors',
      predict: (f) => (browser.anchors.every((a) => anchorEqual(f, a)) ? 'absent' : 'present'),
    },
    ...browser.anchors.map((anchor) => ({
      name: `equal to ${anchor} alone`,
      predict: (f: BrowserFamily) => (anchorEqual(f, anchor) ? 'absent' : 'present'),
    })),
    {
      name: 'document.fonts.check',
      predict: (f: BrowserFamily) => (f.check ? 'present' : 'absent'),
    },
    {
      name: 'the bare family differs from the browser default',
      predict: (f: BrowserFamily) =>
        same(at(f.bare), at(browser.anchorOnly['serif'] ?? {})) ? 'absent' : 'present',
    },
  ],
);

/* the fingerprint: which glyphs tell the present families apart */

const presentFamilies = browser.families.filter((f) => !absentSet.has(f.family));
/** How many of the present families a glyph subset gives distinct widths to. */
function distinctUnder(indices: readonly number[]): number {
  const seen = new Set<string>();
  for (const family of presentFamilies) {
    const widths = at(family.bare);
    seen.add(indices.map((i) => String(widths[i])).join(','));
  }
  return seen.size;
}

const allIndices = browser.glyphs.map((_, i) => i);
const ceiling = distinctUnder(allIndices);

/** The smallest, then lowest-indexed, glyph set that reaches the ceiling. */
function bestSubset(size: number): number[] | null {
  const pick = (start: number, chosen: number[]): number[] | null => {
    if (chosen.length === size) return distinctUnder(chosen) === ceiling ? [...chosen] : null;
    for (let i = start; i < allIndices.length; i++) {
      chosen.push(i);
      const found = pick(i + 1, chosen);
      chosen.pop();
      if (found !== null) return found;
    }
    return null;
  };
  return pick(0, []);
}

let fingerprint: number[] | null = null;
let fingerprintSize = 0;
for (let size = 1; size <= 4 && fingerprint === null; size++) {
  fingerprint = bestSubset(size);
  fingerprintSize = size;
}
if (fingerprint === null) {
  throw new Error(`no glyph set of four or fewer separates ${String(ceiling)} families`);
}

/** Present families that share every advance: what metric-compatible means, measured. */
const classes = new Map<string, string[]>();
for (const family of presentFamilies) {
  const key = at(family.bare).map(String).join(',');
  const bucket = classes.get(key);
  if (bucket) bucket.push(family.family);
  else classes.set(key, [family.family]);
}
const aliasClasses = [...classes.values()]
  .filter((names) => names.length > 1)
  .map((names) => [...names].sort());

/**
 * The best fingerprint made of single characters, which is what the plan asks
 * for - a word is not a glyph.
 */
const characterIndices = allIndices.filter((i) => (browser.glyphs[i] ?? '').length === 1);
function bestCharacters(size: number): number[] | null {
  const pick = (start: number, chosen: number[]): number[] | null => {
    if (chosen.length === size) return distinctUnder(chosen) === ceiling ? [...chosen] : null;
    for (let k = start; k < characterIndices.length; k++) {
      chosen.push(characterIndices[k] as number);
      const found = pick(k + 1, chosen);
      chosen.pop();
      if (found !== null) return found;
    }
    return null;
  };
  return pick(0, []);
}
let characterFingerprint: number[] | null = null;
let characterSize = 0;
for (let size = 1; size <= 3 && characterFingerprint === null; size++) {
  characterFingerprint = bestCharacters(size);
  characterSize = size;
}

/* the fingerprint at two sizes */

const bigger = browser.sizes[0];
const smaller = browser.sizes[1];
if (bigger === undefined || smaller === undefined) throw new Error('two sizes expected');
const ratio = bigger / smaller;
let worstScaleDeviation = 0;
for (const family of presentFamilies) {
  const big = family.bare[String(bigger)];
  const small = family.bare[String(smaller)];
  if (big === undefined || small === undefined) throw new Error(`${family.family}: one size only`);
  for (const [i, value] of big.entries()) {
    const predicted = (small[i] ?? 0) * ratio;
    if (predicted === 0) continue;
    worstScaleDeviation = Math.max(worstScaleDeviation, Math.abs(value - predicted) / predicted);
  }
}
const scaleFree = worstScaleDeviation < 1e-9;

/* -------------------------------------------------------------------------- */
/* the fixture                                                                */
/* -------------------------------------------------------------------------- */

/** Thousandths of a point or pixel: finer than anything measured here. */
const milli = (value: number): number => Math.round(value * 1000);

const packed = (values: readonly number[]): number[] => values.map(milli);
for (const family of browser.families) {
  const exact = at(family.bare);
  const rounded = packed(exact);
  const other = browser.families.find(
    (f) => f !== family && same(packed(at(f.bare)), rounded) && !same(at(f.bare), exact),
  );
  if (other !== undefined) {
    throw new Error(`packing collapses ${family.family} onto ${other.family}`);
  }
}

const slotRows = rows.filter((r) => r.deck === 'slots');
const themeRows = rows.filter((r) => r.deck.startsWith('theme-'));
const aliasRows = ALIAS_FACES.map((_, i) => rowOf('faces', `alias-${String(i)}`));
const installedRows = INSTALLED_FACES.map((_, i) =>
  rowOf('faces', `installed-${String(i).padStart(2, '0')}`),
);

const fixture = {
  $comment:
    'Experiment T7, sub-phase 3.7. Which typeface a run is drawn in when the named one is absent, and how absence is detected in a browser. Written by tools/ground-truth/fonts/substitution/analyse.ts.',
  experiment: 'T7',
  subPhase: '3.7',
  adr: 'docs/adr/phase-3-text/0033-font-substitution-and-the-guard.md',
  units: {
    width: 'thousandths of a point, from TextRange2.BoundWidth',
    advance: 'thousandths of a CSS pixel, from OffscreenCanvas.measureText',
  },
  measuredOn: {
    application: 'Microsoft PowerPoint',
    version: '16.0',
    platform: 'Windows 11 Enterprise 10.0.26200',
    chromium: browser.chromium,
  },
  questions,
  powerpoint: {
    absentSubstitute: 'Calibri',
    installedFamilyCount: readings.installedFamilies.length,
    absent: absentRows.map((r) => ({
      id: r.probe.id,
      typeface: r.probe.faces.latin?.typeface ?? '',
      drawn: r.drawn,
      width: milli(r.width),
    })),
    installed: installedRows.map((r) => ({
      typeface: r.probe.faces.latin?.typeface ?? '',
      drawn: r.drawn,
      width: milli(r.width),
    })),
    alias: aliasRows.map((r) => ({
      typeface: r.probe.faces.latin?.typeface ?? '',
      drawn: r.drawn,
      width: milli(r.width),
    })),
    charsetSubstitute: CHARSET_SUBSTITUTE,
    hints: rows
      .filter((r) => r.deck === 'hints')
      .map((r) => ({
        id: r.probe.id,
        panose: r.probe.faces.latin?.panose ?? null,
        pitchFamily: r.probe.faces.latin?.pitchFamily ?? null,
        charset: r.probe.faces.latin?.charset ?? null,
        drawn: r.drawn,
        width: milli(r.width),
      })),
    widths: widthRows.map((r) => ({
      sz: r.sz,
      text: WIDTH_TEXT[r.textIndex] ?? '',
      absent: milli(r.absent),
      byFace: Object.fromEntries(Object.entries(r.byFace).map(([k, v]) => [k, milli(v)])),
    })),
    slots: slotRows.map((r) => ({
      id: r.probe.id,
      text: r.probe.text,
      faces: r.probe.faces,
      drawn: r.drawn,
      width: milli(r.width),
    })),
    theme: themeRows.map((r) => ({
      deck: r.deck,
      id: r.probe.id,
      themeFonts: r.themeFonts,
      faces: r.probe.faces,
      comName: r.comName,
      drawn: r.drawn,
      width: milli(r.width),
    })),
  },
  browser: {
    chromium: browser.chromium,
    sizes: browser.sizes,
    glyphs: browser.glyphs,
    anchors: browser.anchors,
    /** The glyphs the exhaustive search chose, as characters. */
    fingerprint: fingerprint.map((i) => browser.glyphs[i] ?? ''),
    fingerprintSize,
    /** The same search restricted to single characters, which is what "3-glyph" means. */
    characterFingerprint:
      characterFingerprint === null
        ? null
        : characterFingerprint.map((i) => browser.glyphs[i] ?? ''),
    characterFingerprintSize: characterFingerprint === null ? null : characterSize,
    distinctFamilies: ceiling,
    presentFamilies: presentFamilies.length,
    scaleFree,
    worstScaleDeviation: Math.round(worstScaleDeviation * 1e6) / 1e6,
    aliasClasses,
    anchorOnly: Object.fromEntries(
      browser.anchors.map((anchor) => [anchor, packed(at(browser.anchorOnly[anchor] ?? {}))]),
    ),
    families: browser.families.map((f) => ({
      family: f.family,
      absent: absentSet.has(f.family),
      check: f.check,
      bare: packed(at(f.bare)),
      anchorEqual: Object.fromEntries(browser.anchors.map((a) => [a, anchorEqual(f, a)])),
    })),
  },
};

const out = repoPath('corpus/ground-truth/font-substitution.json');
// Through Prettier, so the fixture the suite hashes is the one `pnpm format:check` wants.
const config = await resolveConfig(out);
writeFileSync(out, await format(JSON.stringify(fixture), { ...config, filepath: out }));

console.log(`${String(rows.length)} probe(s) across ${String(readings.decks.length)} package(s)`);
for (const q of questions) {
  const table = q.scores.map((s) => `${s.name} ${String(s.fits)}/${String(s.total)}`).join('  |  ');
  console.log(`\n${q.key}\n  ${q.answer}\n  ${table}`);
}
console.log(
  `\nfingerprint: ${fingerprint.map((i) => JSON.stringify(browser.glyphs[i])).join(', ')}` +
    ` separates ${String(ceiling)} of ${String(presentFamilies.length)} present families`,
);
console.log(
  characterFingerprint === null
    ? 'no set of three or fewer single characters reaches that ceiling'
    : `single characters: ${characterFingerprint.map((i) => JSON.stringify(browser.glyphs[i])).join(', ')} (${String(characterSize)})`,
);
console.log(`alias classes: ${aliasClasses.map((c) => c.join(' = ')).join('; ') || '(none)'}`);
console.log(
  `advances linear in size: ${scaleFree ? 'yes' : 'no'}, worst deviation ${(worstScaleDeviation * 100).toFixed(3)}%`,
);
console.log(`\nwrote ${out}`);
