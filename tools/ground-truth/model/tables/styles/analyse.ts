/**
 * Experiment C8, step 3 - score every reading of a table's style against what PowerPoint drew.
 *
 * ```
 * node tools/ground-truth/model/tables/styles/analyse.ts <dir> [--fixture <path>]
 * ```
 *
 * Throws unless exactly one candidate fits every informative row of each question, and unless the
 * roster, the serialisations, the determinism and the contamination guards all hold.
 */

import { createHash } from 'node:crypto';
import { existsSync, readFileSync, writeFileSync } from 'node:fs';
import { join } from 'node:path';

import { readBmp, type Bitmap } from '../../../lib/bmp.ts';
import { fixtureJson } from '../../../lib/json.ts';
import { entry, readZip } from '../../../lib/zip.ts';
import {
  allProbes,
  catalogueFrom,
  definitionsIn,
  stylesIn,
  OFFICE_JS_TABLE_STYLES,
  RECORD_ONLY,
  rosterName,
  TABLE_RECT,
  tableStylesOf,
  type AuthorLog,
  type Observation,
  type Probe,
} from './probes.ts';

const dir = process.argv[2];
if (dir === undefined) {
  throw new Error(
    'usage: tools/ground-truth/model/tables/styles/analyse.ts <dir> [--fixture <path>]',
  );
}
const fixtureAt = process.argv.indexOf('--fixture');
const fixturePath = fixtureAt > 0 ? process.argv[fixtureAt + 1] : undefined;

/** `CAPS.perFile` in tools/corpus/manifest/schema.ts. */
const CORPUS_FILE_CAP = 512 * 1024;
const OFFICE_JS_SOURCE =
  'https://learn.microsoft.com/javascript/api/powerpoint/powerpoint.tablestylesettings';

function fail(message: string): never {
  throw new Error(`C8: ${message}`);
}

const BOM = String.fromCharCode(0xfeff);
const readJson = <T>(name: string): T => {
  const text = readFileSync(join(dir, name), 'utf8');
  return JSON.parse(text.startsWith(BOM) ? text.slice(1) : text) as T;
};

/* -------------------------------------------------------------------------- */
/* inputs                                                                     */
/* -------------------------------------------------------------------------- */

interface Session {
  readonly pid: number;
  readonly started: string;
  readonly version: string;
  readonly build: string;
}

interface FullAuthorLog extends AuthorLog {
  readonly session: Session & {
    readonly uiLanguage: number;
    readonly newTableStyle: { readonly id: string; readonly name: string };
  };
  readonly gallery: readonly {
    readonly seq: number;
    readonly uiaName: string;
    readonly comName: string;
    readonly styleId: string;
    readonly sentinel: string;
    readonly method: string;
  }[];
  readonly categories: readonly { readonly category: string; readonly items: readonly string[] }[];
  readonly decks: Readonly<
    Record<
      string,
      readonly {
        readonly slide: number;
        readonly applied: string;
        readonly id: string;
        readonly name: string;
        readonly cells?: readonly {
          readonly r: number;
          readonly c: number;
          readonly visible: number;
          readonly type: number;
          readonly rgb: number;
          readonly transparency: number;
        }[];
      }[]
    >
  >;
  readonly api: readonly { readonly applied: string; readonly ok: boolean; readonly id: string }[];
}

interface SlideReading {
  readonly styleId: string | null;
  readonly styleName: string | null;
  readonly error: string | null;
  readonly a: string;
  readonly b: string | null;
}

interface ProbeReading {
  readonly id: string;
  readonly opened: boolean;
  readonly repaired: boolean | null;
  readonly slides: readonly SlideReading[];
  readonly resaved: string | null;
  readonly reread: readonly SlideReading[];
  readonly resaved2: string | null;
}

interface Readings {
  readonly session: Session;
  readonly probes: readonly ProbeReading[];
  readonly again: readonly { readonly id: string; readonly slide: SlideReading | null }[];
  readonly reapplied: readonly { readonly id: string; readonly file: string | null }[];
}

const log = readJson<FullAuthorLog>('author-styles-log.json');
const readings = readJson<Readings>('style-readings.json');
const deck = (name: string): Uint8Array => readFileSync(join(dir, name));
const partOf = (name: string): string | null => tableStylesOf(deck(name));

const authoredOff =
  partOf('pp-styles-off.pptx') ?? fail('pp-styles-off.pptx has no table-style part');
const catalogue = catalogueFrom(log, authoredOff);
const probes = allProbes(catalogue);
const readingOf = new Map(readings.probes.map((r) => [r.id, r]));
for (const probe of probes) if (!readingOf.has(probe.id)) fail(`no reading for ${probe.id}`);

/* -------------------------------------------------------------------------- */
/* sessions and the roster                                                    */
/* -------------------------------------------------------------------------- */

if (log.session.build !== readings.session.build) {
  fail(`PowerPoint changed between the runs: ${log.session.build} then ${readings.session.build}`);
}
if (log.session.pid === readings.session.pid && log.session.started === readings.session.started) {
  fail('author.ps1 and read.ps1 ran in one PowerPoint process; the serialisations need two');
}
if (log.session.uiLanguage !== 1033) fail(`UI language ${String(log.session.uiLanguage)}`);

const gallery = log.gallery;
const galleryIds = new Set(gallery.map((g) => g.styleId));
if (gallery.length !== 74 || galleryIds.size !== 74) {
  fail(
    `the gallery gave ${String(gallery.length)} items and ${String(galleryIds.size)} GUIDs, not 74`,
  );
}
for (const g of gallery) {
  if (g.uiaName !== g.comName) fail(`${g.uiaName}: COM names it ${g.comName}`);
  if (g.styleId === g.sentinel) fail(`${g.uiaName}: the invoke left the sentinel in place`);
}
const officeJs = OFFICE_JS_TABLE_STYLES.map((key, index) => ({ key, index, ...rosterName(key) }));
const byUiName = new Map(officeJs.map((o) => [o.ui, o]));
if (byUiName.size !== 74) fail('the Office.js roster maps two keys to one name');
gallery.forEach((g, index) => {
  const o = byUiName.get(g.uiaName) ?? fail(`the gallery's "${g.uiaName}" has no Office.js key`);
  if (o.index !== index)
    fail(`"${g.uiaName}" is item ${String(index)} in the gallery, ${String(o.index)} in Office.js`);
});
const categoryOf = (name: string): string[] =>
  log.categories.filter((c) => c.items.includes(name)).map((c) => c.category);
for (const [id, xml] of catalogue.definitions) {
  const g = gallery.find((x) => x.styleId === id) ?? fail(`no gallery item for ${id}`);
  if (!xml.startsWith(`<a:tblStyle styleId="${id}" styleName="${g.uiaName}">`)) {
    fail(`${g.uiaName}'s saved definition does not name it: ${xml.slice(0, 120)}`);
  }
}

/* -------------------------------------------------------------------------- */
/* the serialisations                                                         */
/* -------------------------------------------------------------------------- */

interface Row {
  readonly source: string;
  readonly session: 1 | 2;
  readonly theme: 'office' | Probe['theme'];
  readonly flags: 'off' | 'on' | 'default' | 'first-band';
  readonly id: string;
  readonly equal: boolean;
}

const serialRows: Row[] = [];
const addSource = (
  source: string,
  part: string | null,
  meta: Pick<Row, 'session' | 'theme' | 'flags'>,
): void => {
  if (part === null) return;
  for (const [id, xml] of stylesIn(part)) {
    const reference = catalogue.definitions.get(id);
    if (reference === undefined) continue;
    serialRows.push({ source, id, ...meta, equal: xml === reference });
  }
};
addSource('pp-styles-on.pptx', partOf('pp-styles-on.pptx'), {
  session: 1,
  theme: 'office',
  flags: 'on',
});
gallery.forEach((_, k) => {
  const name = `pp-one-${String(k + 1).padStart(2, '0')}.pptx`;
  addSource(name, partOf(name), { session: 1, theme: 'office', flags: 'default' });
});
for (const name of ['pp-stock.pptx', 'pp-restyle.pptx', 'pp-deleted.pptx', 'pp-two-same.pptx']) {
  addSource(name, partOf(name), { session: 1, theme: 'office', flags: 'default' });
}
for (const reapplied of readings.reapplied) {
  const probe = probes.find((p) => p.id === reapplied.id) ?? fail(`no probe ${reapplied.id}`);
  if (reapplied.file === null) fail(`${reapplied.id} was not reapplied`);
  const part = partOf(reapplied.file);
  if (part === null || definitionsIn(part).size < 74)
    fail(`${reapplied.file} does not define all 74`);
  addSource(reapplied.file, part, { session: 2, theme: probe.theme, flags: 'first-band' });
}

interface Scored {
  readonly name: string;
  readonly fits: number;
  readonly of: number;
  readonly wrong: readonly string[];
}

function choose(question: string, scores: readonly Scored[]): Scored {
  const perfect = scores.filter((s) => s.fits === s.of && s.of > 0);
  if (perfect.length !== 1) {
    fail(
      `${question}: ${String(perfect.length)} candidate(s) fit every row\n` +
        scores
          .map(
            (s) =>
              `  ${s.name}: ${String(s.fits)}/${String(s.of)} ${s.wrong.slice(0, 6).join(', ')}`,
          )
          .join('\n'),
    );
  }
  return perfect[0] as Scored;
}

const SERIAL_CANDIDATES: Readonly<Record<string, (row: Row) => boolean>> = {
  'one string per GUID': () => true,
  'baked into the theme': (row) => row.theme === 'office' || row.theme === 's1',
  'pruned by the flags': (row) => row.flags === 'off',
  'different per session': (row) => row.session === 1,
};
const serialScores: Scored[] = Object.entries(SERIAL_CANDIDATES).map(([name, predictsEqual]) => {
  const wrong = serialRows.filter((row) => predictsEqual(row) !== row.equal);
  return {
    name,
    fits: serialRows.length - wrong.length,
    of: serialRows.length,
    wrong: wrong.map((r) => `${r.source}:${r.id}`),
  };
});
const serialWinner = choose('serialisation', serialScores);
const serialSources = [...new Set(serialRows.map((r) => r.source))];

/* -------------------------------------------------------------------------- */
/* pixels                                                                     */
/* -------------------------------------------------------------------------- */

const MARGIN = 4;
const bitmaps = new Map<string, Bitmap>();
const bmpBytes = (path: string): Uint8Array => readFileSync(join(dir, path));
function bitmap(path: string): Bitmap {
  const hit = bitmaps.get(path);
  if (hit !== undefined) return hit;
  const b = readBmp(bmpBytes(path));
  bitmaps.set(path, b);
  return b;
}

interface Seen {
  /** SHA-256 of the table's region, grown by `MARGIN` so a centred outer border is inside it. */
  readonly region: string;
  readonly bare: boolean;
  /** The one colour all sixteen cell centres share, or null. */
  readonly flat: string | null;
}

const hexOf = ([r, g, b]: readonly [number, number, number]): string =>
  [r, g, b]
    .map((v) => v.toString(16).padStart(2, '0'))
    .join('')
    .toUpperCase();

function see(path: string): Seen {
  const b = bitmap(path);
  const { x, y, w, h } = TABLE_RECT;
  const hash = createHash('sha256');
  const background = hexOf(b.pixel(20, b.height - 20));
  let bare = true;
  for (let py = y - MARGIN; py < y + h + MARGIN; py++) {
    const row = new Uint8Array((w + 2 * MARGIN) * 3);
    for (let px = x - MARGIN; px < x + w + MARGIN; px++) {
      const p = b.pixel(px, py);
      row.set(p, (px - x + MARGIN) * 3);
      if (bare && hexOf(p) !== background) bare = false;
    }
    hash.update(row);
  }
  const centres = new Set<string>();
  for (let r = 0; r < 4; r++) {
    for (let c = 0; c < 4; c++) centres.add(hexOf(b.pixel(x + 54 + 108 * c, y + 18 + 36 * r)));
  }
  return {
    region: hash.digest('hex'),
    bare,
    flat: centres.size === 1 ? ([...centres][0] ?? null) : null,
  };
}

/** Every slide of every control, by the name predictions use. */
const controls = new Map<string, Seen>();
for (const probe of probes.filter((p) => p.control === true)) {
  const reading = readingOf.get(probe.id);
  if (reading === undefined || !reading.opened || reading.repaired === true)
    fail(`control ${probe.id} did not open clean`);
  reading.slides.forEach((slide, k) => {
    const seen = see(slide.a);
    controls.set(`${probe.id}#${String(k)}`, seen);
    if (reading.slides.length === 1) controls.set(probe.id, seen);
  });
}

// A slide exported twice must be the same picture, or no comparison here means anything.
let determinism = 0;
for (const reading of readings.probes) {
  for (const slide of reading.slides) {
    if (slide.b === null) continue;
    const a = bmpBytes(slide.a);
    const b = bmpBytes(slide.b);
    if (Buffer.compare(Buffer.from(a), Buffer.from(b)) !== 0)
      fail(`${slide.a} and ${slide.b} differ`);
    determinism += 1;
  }
}
// A modified definition must not leak into the next package's drawing of the real one.
for (const again of readings.again) {
  if (again.slide === null) fail(`${again.id} did not reopen`);
  const first = readingOf.get(again.id)?.slides[0]?.a ?? fail(`${again.id} has no first reading`);
  if (see(again.slide.a).region !== see(first).region) {
    fail(`${again.id} drew differently after the probes: a style leaked between packages`);
  }
}

function holds(observation: Observation, seen: Seen): boolean {
  if (observation === 'bare') return seen.bare;
  if (observation.startsWith('signature:'))
    return seen.flat === observation.slice('signature:'.length);
  const control =
    controls.get(observation.slice('control:'.length)) ?? fail(`no control ${observation}`);
  return seen.region === control.region;
}

/** Two predictions no picture could tell apart, judged from the controls alone. */
function indistinguishable(a: Observation, b: Observation): boolean {
  if (a === b) return true;
  const as = (o: Observation): Seen | null =>
    o.startsWith('control:') ? (controls.get(o.slice('control:'.length)) ?? null) : null;
  const ca = as(a);
  const cb = as(b);
  if (ca !== null && cb !== null) return ca.region === cb.region;
  if (ca !== null) return holds(b, ca);
  if (cb !== null) return holds(a, cb);
  return false;
}

/* -------------------------------------------------------------------------- */
/* the questions                                                              */
/* -------------------------------------------------------------------------- */

interface QuestionResult {
  readonly winner: string;
  readonly candidates: readonly Scored[];
  readonly uninformative: readonly string[];
  readonly repaired: readonly string[];
}

const questions = new Map<string, Probe[]>();
for (const probe of probes) {
  if (probe.control === true) continue;
  questions.set(probe.question, [...(questions.get(probe.question) ?? []), probe]);
}

const results: Record<string, QuestionResult> = {};
for (const [question, members] of questions) {
  if (RECORD_ONLY.has(question)) continue;
  const candidates = Object.keys(members[0]?.slides[0]?.predict ?? {});
  const tallies = new Map(candidates.map((c) => [c, { fits: 0, of: 0, wrong: [] as string[] }]));
  const uninformative: string[] = [];
  const repaired: string[] = [];
  for (const probe of members) {
    const reading = readingOf.get(probe.id) ?? fail(`no reading for ${probe.id}`);
    if (!reading.opened || reading.repaired === true) {
      if (probe.hostile !== true)
        fail(`${probe.id} is not hostile, and PowerPoint repaired or refused it`);
      repaired.push(probe.id);
      continue;
    }
    probe.slides.forEach((slide, k) => {
      const row = `${probe.id}#${String(k)}`;
      const predictions = Object.values(slide.predict);
      // A row every candidate predicts alike, judged from the controls, separates nothing.
      const [first] = predictions;
      if (first !== undefined && predictions.every((p) => indistinguishable(first, p))) {
        uninformative.push(row);
        return;
      }
      const seen = see(reading.slides[k]?.a ?? fail(`${row} was not exported`));
      for (const candidate of candidates) {
        const tally = tallies.get(candidate);
        const predicted = slide.predict[candidate];
        if (tally === undefined || predicted === undefined)
          fail(`${row} has no prediction for ${candidate}`);
        tally.of += 1;
        if (holds(predicted, seen)) tally.fits += 1;
        else tally.wrong.push(row);
      }
    });
  }
  const scores = candidates.map((name) => {
    const t = tallies.get(name) ?? fail(name);
    return { name, fits: t.fits, of: t.of, wrong: t.wrong };
  });
  const clean = scores[0]?.of ?? 0;
  const winner =
    clean === 0 && repaired.length > 0 && uninformative.length === 0
      ? 'repaired'
      : choose(question, scores).name;
  results[question] = { winner, candidates: scores, uninformative, repaired };
}

/* -------------------------------------------------------------------------- */
/* what PowerPoint wrote back                                                 */
/* -------------------------------------------------------------------------- */

const utf8 = (bytes: Uint8Array): string => new TextDecoder().decode(bytes);

function slideIds(path: string): (string | null)[] {
  const entries = readZip(deck(path));
  const out: (string | null)[] = [];
  for (let n = 1; ; n++) {
    const slide = entry(entries, `ppt/slides/slide${String(n)}.xml`);
    if (slide === undefined) return out;
    const id = /<a:tableStyleId>([^<]*)<\/a:tableStyleId>/.exec(utf8(slide))?.[1];
    out.push(id ?? null);
  }
}

function markupOf(path: string): {
  tblPr: (string | null)[];
  part: string | null;
  partName: string | null;
  rel: boolean;
} {
  const entries = readZip(deck(path));
  const tblPr: (string | null)[] = [];
  for (let n = 1; ; n++) {
    const slide = entry(entries, `ppt/slides/slide${String(n)}.xml`);
    if (slide === undefined) break;
    const text = utf8(slide);
    tblPr.push(/<a:tblPr[\s\S]*?(?:\/>|<\/a:tblPr>)/.exec(text)?.[0] ?? null);
  }
  const rels = utf8(entry(entries, 'ppt/_rels/presentation.xml.rels') ?? new Uint8Array());
  const target = /Type="[^"]*\/tableStyles"[^>]*Target="([^"]+)"/.exec(rels)?.[1];
  const named = target === undefined ? null : `ppt/${target}`;
  const fixed = entry(entries, 'ppt/tableStyles.xml');
  const at = named ?? (fixed === undefined ? null : 'ppt/tableStyles.xml');
  const bytes = at === null ? undefined : entry(entries, at);
  return {
    tblPr,
    part: bytes === undefined ? null : utf8(bytes),
    partName: at === null ? null : `/${at}`,
    rel: target !== undefined,
  };
}

interface WrittenStyle {
  readonly styleId: string;
  readonly as: 'powerpoint' | 'as-written' | 'other';
}

function writtenBack(probe: Probe, path: string, asWritten: string | null) {
  const part = partOf(path);
  const ids = slideIds(path);
  const own = new Set(asWritten === null ? [] : stylesIn(asWritten).map(([, xml]) => xml));
  const styles: WrittenStyle[] =
    part === null
      ? []
      : stylesIn(part).map(([styleId, xml]) => ({
          styleId,
          as:
            catalogue.definitions.get(styleId) === xml
              ? 'powerpoint'
              : own.has(xml)
                ? 'as-written'
                : 'other',
        }));
  const written = probe.slides.map((slide, k) => {
    const before = /<a:tableStyleId>([^<]*)<\/a:tableStyleId>/.exec(slide.tblPr ?? '')?.[1] ?? null;
    const after = ids[k] ?? null;
    const fate =
      before === after
        ? 'kept'
        : after === null
          ? 'dropped'
          : before !== null && after === before.trim().toUpperCase()
            ? 'normalised'
            : 'replaced';
    return { before, after, fate };
  });
  return {
    part: part !== null,
    def: part === null ? null : (/def="([^"]+)"/.exec(part)?.[1] ?? null),
    styles,
    ids: written,
  };
}

const probeRecords = probes.map((probe) => {
  const reading = readingOf.get(probe.id) ?? fail(probe.id);
  const markup = markupOf(`style-${probe.id}.pptx`);
  const sweep = probe.slides.length > 1;
  const drawn = reading.opened
    ? reading.slides.map((slide) => {
        const seen = see(slide.a);
        // Single-slide controls by their own name, the sweeps' slides by `sweep-ctl#k`.
        const matches = [...controls]
          .filter(([name]) => name.startsWith('sweep') || !name.includes('#'))
          .filter(([name]) => !sweep || name.startsWith('sweep'))
          .filter(([, control]) => control.region === seen.region)
          .map(([name]) => name);
        return { matches, bare: seen.bare, flat: seen.flat };
      })
    : [];
  const back = reading.resaved === null ? null : writtenBack(probe, reading.resaved, markup.part);
  const reread = reading.reread.map((slide, k) => ({
    same:
      reading.slides[k] === undefined
        ? false
        : see(slide.a).region === see(reading.slides[k]?.a ?? '').region,
  }));
  const resaved2Identical =
    reading.resaved === null || reading.resaved2 === null
      ? null
      : Buffer.compare(Buffer.from(deck(reading.resaved)), Buffer.from(deck(reading.resaved2))) ===
        0;
  return {
    id: probe.id,
    question: probe.question,
    theme: probe.theme,
    control: probe.control === true,
    hostile: probe.hostile === true,
    opened: reading.opened,
    repaired: reading.repaired,
    markup: sweep
      ? {
          tblPr: `${String(probe.slides.length)} slides, one per built-in in gallery order: ${markup.tblPr[0] ?? ''}`,
          tableStyles: probe.part === null ? null : 'every definition in gallery order',
          partName: markup.partName,
          rel: markup.rel,
        }
      : {
          tblPr: markup.tblPr[0] ?? null,
          tableStyles: markup.part,
          partName: markup.partName,
          rel: markup.rel,
        },
    predict: sweep ? null : (probe.slides[0]?.predict ?? null),
    drawn: sweep ? summariseSweep(drawn) : (drawn[0] ?? null),
    comStyle: reading.slides
      .slice(0, sweep ? 0 : 1)
      .map((s) => ({ id: s.styleId, name: s.styleName })),
    resaved: back === null ? null : sweep ? summariseWriteBack(back) : back,
    reread: sweep ? null : { same: reread.every((r) => r.same), resaved2Identical },
  };
});

function summariseSweep(drawn: readonly { matches: string[]; bare: boolean }[]) {
  return {
    slides: drawn.length,
    bare: drawn.flatMap((d, k) => (d.bare ? [k] : [])),
    matches: drawn.map((d) => d.matches.filter((m) => m.includes('#'))),
  };
}

function summariseWriteBack(back: ReturnType<typeof writtenBack>) {
  const tally = (fate: string): number => back.ids.filter((i) => i.fate === fate).length;
  return {
    part: back.part,
    def: back.def,
    styles: back.styles.length,
    asPowerPoint: back.styles.filter((s) => s.as === 'powerpoint').length,
    ids: {
      kept: tally('kept'),
      normalised: tally('normalised'),
      dropped: tally('dropped'),
      replaced: tally('replaced'),
    },
  };
}

/* -------------------------------------------------------------------------- */
/* the object model as an instrument                                          */
/* -------------------------------------------------------------------------- */

let comRows = 0;
let comFits = 0;
const comWrong: string[] = [];
for (const deckName of ['pp-styles-off.pptx', 'pp-styles-on.pptx']) {
  const stem = deckName.replace('.pptx', '');
  for (const slide of log.decks[deckName] ?? []) {
    const image = `author-bmp/${stem}-${String(slide.slide).padStart(2, '0')}.bmp`;
    if (!existsSync(join(dir, image))) fail(`${image} is missing`);
    const b = bitmap(image);
    for (const cell of slide.cells ?? []) {
      if (cell.visible !== -1 || cell.transparency !== 0) continue;
      const rgb = cell.rgb;
      const com = hexOf([rgb & 0xff, (rgb >> 8) & 0xff, (rgb >> 16) & 0xff]);
      const pixel = hexOf(b.pixel(72 + 54 + 108 * (cell.c - 1), 72 + 18 + 36 * (cell.r - 1)));
      comRows += 1;
      if (com === pixel) comFits += 1;
      else
        comWrong.push(
          `${stem}#${String(slide.slide)} (${String(cell.r)},${String(cell.c)}) ${com} vs ${pixel}`,
        );
    }
  }
}

/* -------------------------------------------------------------------------- */
/* one template per family                                                    */
/* -------------------------------------------------------------------------- */

const ACCENT = /val="accent([1-6])"/g;
const templates = [...new Set(officeJs.map((o) => o.family))].map((family) => {
  const members = officeJs.filter((o) => o.family === family);
  const xmlOf = (key: string): string => {
    const o = officeJs.find((x) => x.key === key) ?? fail(key);
    const g = gallery.find((x) => x.uiaName === o.ui) ?? fail(o.ui);
    return catalogue.definitions.get(g.styleId) ?? fail(g.styleId);
  };
  // Each accent variant names its own accent and no other: the guard on the GUID-to-name binding.
  for (const o of members) {
    const used = new Set([...xmlOf(o.key).matchAll(ACCENT)].map((m) => Number(m[1])));
    const variant = /^accent(\d)$/.exec(o.variant)?.[1];
    const pair = /^pair(\d)$/.exec(o.variant)?.[1];
    const expected =
      variant !== undefined
        ? [Number(variant)]
        : pair !== undefined
          ? [2 * Number(pair) - 1, 2 * Number(pair)]
          : [];
    const sorted = [...used].sort();
    if (sorted.join() !== expected.join()) {
      fail(`${o.ui} uses accents [${sorted.join()}], its name says [${expected.join()}]`);
    }
  }
  const first = members.find((o) => o.variant === 'accent1' || o.variant === 'pair1');
  if (first === undefined) return { family, accents: null };
  const base = xmlOf(first.key).replace(/^<a:tblStyle [^>]*>/, '');
  const fits = members
    .filter((o) => o.variant !== 'plain' && o !== first)
    .map((o) => {
      const shift =
        first.variant === 'pair1'
          ? 2 * (Number(o.variant.slice(4)) - 1)
          : Number(o.variant.slice(6)) - 1;
      const predicted = base.replace(
        ACCENT,
        (_, n: string) => `val="accent${String(Number(n) + shift)}"`,
      );
      return predicted === xmlOf(o.key).replace(/^<a:tblStyle [^>]*>/, '');
    });
  return { family, accents: `${String(fits.filter(Boolean).length)}/${String(fits.length)}` };
});

/* -------------------------------------------------------------------------- */
/* the fixture                                                                */
/* -------------------------------------------------------------------------- */

const sweepBack = (id: string) => probeRecords.find((p) => p.id === id)?.resaved ?? null;
const answers = {
  stock: definitionsIn(partOf('pp-stock.pptx') ?? '').size,
  restyle: [...definitionsIn(partOf('pp-restyle.pptx') ?? '').keys()],
  deleted: [...definitionsIn(partOf('pp-deleted.pptx') ?? '').keys()],
  twoSame: [...definitionsIn(partOf('pp-two-same.pptx') ?? '').keys()],
};

const fixture = {
  $comment:
    'Experiment C8 - the 74 built-in table styles, enumerated through PowerPoint’s own Table Styles ' +
    'gallery by tools/ground-truth/model/tables/styles/author.ps1 and serialised by PowerPoint; which ' +
    'style an a:tableStyleId draws, measured on packages written by ' +
    'tools/ground-truth/model/tables/styles/build-deck.ts and read by ' +
    'tools/ground-truth/model/tables/styles/read.ps1. Scored by ' +
    'tools/ground-truth/model/tables/styles/analyse.ts. Definitions are PowerPoint’s bytes.',
  experiment: 'C8',
  subPhase: '4.2',
  generatedBy: 'tools/ground-truth/model/tables/styles/analyse.ts',
  measuredOn: {
    application: 'Microsoft PowerPoint',
    version: log.session.version,
    build: log.session.build,
    uiLanguage: log.session.uiLanguage,
    sessions: 2,
  },
  roster: {
    source: OFFICE_JS_SOURCE,
    styles: gallery.map((g, index) => {
      const o = byUiName.get(g.uiaName) ?? fail(g.uiaName);
      return {
        key: o.key,
        name: g.uiaName,
        id: g.styleId,
        family: o.family,
        variant: o.variant,
        gallery: { index, categories: categoryOf(g.uiaName) },
      };
    }),
  },
  definitions: Object.fromEntries(catalogue.definitions),
  defaults: {
    newTable: log.session.newTableStyle.id,
    def: /def="([^"]+)"/.exec(authoredOff)?.[1] ?? null,
  },
  api: log.api,
  findings: {
    serialisation: {
      winner: serialWinner.name,
      sources: serialSources.length,
      rows: serialRows.length,
      candidates: serialScores.map(({ name, fits, of, wrong }) => ({
        name,
        fits,
        of,
        wrong: wrong.slice(0, 5),
      })),
    },
    instrument: {
      name: 'COM Cell.Shape.Fill against the exported pixel',
      fits: comFits,
      of: comRows,
      wrong: comWrong.slice(0, 10),
    },
    determinism: { pairs: determinism },
    ...Object.fromEntries(
      Object.entries(results).map(([q, r]) => [
        q,
        { ...r, candidates: r.candidates.map((c) => ({ ...c, wrong: c.wrong.slice(0, 8) })) },
      ]),
    ),
    writeBack: {
      authored: answers,
      sweeps: {
        'sweep-nopart': sweepBack('sweep-nopart'),
        'sweep-empty': sweepBack('sweep-empty'),
        'sweep2-nopart': sweepBack('sweep2-nopart'),
        'sweep-ctl': sweepBack('sweep-ctl'),
      },
    },
    templates,
  },
  probes: probeRecords,
};

const text = fixtureJson(fixture);
if (Buffer.byteLength(text) > CORPUS_FILE_CAP) {
  fail(
    `the fixture is ${String(Buffer.byteLength(text))} bytes, over the corpus cap of ${String(CORPUS_FILE_CAP)}`,
  );
}
for (const [q, r] of Object.entries(results)) {
  const winner = r.candidates.find((c) => c.name === r.winner);
  console.log(
    `${q.padEnd(18)} ${r.winner.padEnd(16)} ${winner === undefined ? '' : `${String(winner.fits)}/${String(winner.of)}`}  ${r.candidates
      .filter((c) => c.name !== r.winner)
      .map((c) => `${c.name} ${String(c.fits)}/${String(c.of)}`)
      .join(', ')}${r.repaired.length > 0 ? `  repaired: ${r.repaired.join(' ')}` : ''}`,
  );
}
console.log(
  `serialisation      ${serialWinner.name} ${String(serialWinner.fits)}/${String(serialWinner.of)} over ${String(serialSources.length)} sources`,
);
console.log(
  `instrument         COM fill = pixel ${String(comFits)}/${String(comRows)}; determinism ${String(determinism)} pairs`,
);
console.log(
  `templates          ${templates.map((t) => `${t.family} ${t.accents ?? '-'}`).join('; ')}`,
);
if (fixturePath !== undefined) {
  const before = existsSync(fixturePath) ? readFileSync(fixturePath, 'utf8') : null;
  writeFileSync(fixturePath, text);
  console.log(
    before === text
      ? 'fixture unchanged'
      : `fixture rewritten (${String(Buffer.byteLength(text))} bytes)`,
  );
}
