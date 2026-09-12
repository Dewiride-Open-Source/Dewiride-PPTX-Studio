/**
 * Experiment T14 - the gate, as a pure function over one measurement.
 *
 * Split out of `tools/ground-truth/fonts/real-faces/analyse.ts` because this
 * experiment's analysis is a CI gate, and a gate that has never been shown to
 * go red is not a gate. It touches no font, no browser and no filesystem, so a
 * maintainer can re-score the artifact on a machine with no fonts at all.
 */

import {
  ADVANCE_QUANTUM,
  BOX_READINGS,
  BOX_TOLERANCE,
  EXACT_PX,
  READINGS,
  SHIPPED,
  UNKERNED,
  glyphsOf,
  rasteriserOf,
  shippedBoxFor,
  widthTolerance,
  type BoxCandidates,
  type BoxPair,
  type Reading,
} from './probes.ts';
import { BOX_ROUNDINGS, type BoxRounding } from '../../lib/box-rounding.ts';

/* -------------------------------------------------------------------------- */
/* the measurement                                                            */
/* -------------------------------------------------------------------------- */

export interface RunSample {
  readonly id: string;
  readonly script: string;
  readonly gated: boolean;
  readonly asks: string;
  readonly text: string;
}

export interface Row {
  readonly sample: string;
  readonly px: number;
  /** `measureText(...).width`, in CSS pixels. */
  readonly browser: number;
  /** Adjacent pairs the reader kerned, each carrying its own browser rounding. */
  readonly kerns?: number;
  readonly readings: Readonly<Partial<Record<Reading, number>>>;
}

export interface RunFace {
  readonly file: string;
  readonly sha256: string;
  readonly family: string;
  readonly subfamily: string;
  readonly unitsPerEm: number;
  readonly metricsSource: string;
  /**
   * The browser's box in pixels at `boxPx`; the reader's in font units.
   *
   * `candidates` is the three pairs the reader chose between; a run that
   * carries none can be scored for width but cannot be asked which table won.
   */
  readonly box: {
    readonly browser: BoxPair;
    readonly reader: BoxPair;
    readonly candidates?: BoxCandidates;
  };
  /** Samples this face has no glyph for, which neither side was asked about. */
  readonly uncovered: readonly string[];
  readonly rows: readonly Row[];
}

export interface Run {
  readonly experiment: string;
  readonly subPhase: string;
  readonly adr: string;
  readonly chromium: string;
  readonly image: {
    readonly platform: string;
    readonly release: string;
    readonly imageOs: string | null;
    readonly imageVersion: string | null;
  };
  readonly directories: readonly string[];
  readonly sizes: readonly number[];
  readonly boxPx: number;
  readonly samples: readonly RunSample[];
  readonly indexed: { readonly files: number; readonly faces: number; readonly measured: number };
  readonly skipped: readonly { readonly file: string; readonly why: string }[];
  readonly faces: readonly RunFace[];
}

/* -------------------------------------------------------------------------- */
/* the verdict                                                                */
/* -------------------------------------------------------------------------- */

export type FailureKind =
  | 'no-directories'
  | 'no-faces'
  | 'no-gated-comparisons'
  | 'unmeasured'
  | 'advance-quantum'
  | 'width'
  | 'rival-beats-shipped'
  | 'kerning-unseparated'
  | 'face-box'
  | 'face-box-rival'
  | 'shaping-invisible';

export interface Failure {
  readonly kind: FailureKind;
  /** The font file, the reading, or the sample. */
  readonly subject: string;
  /** Both widths, the delta in px and in font units, and what it was allowed. */
  readonly detail: string;
}

export interface ReadingScore {
  readonly reading: string;
  /** Rows the reading lands on within `widthTolerance`. */
  readonly inside: number;
  /** Rows it reproduces to the last bit, which a whole-pixel browser decides. */
  readonly exact: number;
  readonly of: number;
  readonly worstOverTolerance: number;
}

export interface ScriptScore {
  readonly script: string;
  readonly gated: boolean;
  readonly faces: number;
  readonly comparisons: number;
  readonly inside: number;
  readonly medianRelative: number;
  readonly worstOverTolerance: number;
}

export interface BoxScore {
  readonly name: string;
  readonly fits: number;
  readonly of: number;
}

export interface BoxReadingScore {
  readonly reading: string;
  /** Faces whose recorded candidates carry the pair this reading names. */
  readonly of: number;
  /** Faces where it lands within one rounding of the browser's box. */
  readonly fits: number;
  /** Faces where it predicts a different box from the shipped reading. */
  readonly separates: number;
  /** The largest miss on any face, in pixels at `boxPx`. */
  readonly worstPx: number;
}

export interface ShapingScore {
  readonly sample: string;
  readonly faces: number;
  /** The reader's width over the browser's: the price of the missing feature. */
  readonly medianRatio: number;
  readonly worstRatio: number;
  /** Rows where the price is larger than the browser's own rounding. */
  readonly outsideTolerance: number;
  readonly of: number;
}

export interface Verdict {
  /** Empty means green. */
  readonly failures: readonly Failure[];
  readonly scoreboard: readonly ReadingScore[];
  readonly byScript: readonly ScriptScore[];
  readonly faceBox: {
    readonly agreed: number;
    readonly of: number;
    /** Faces that recorded the pairs the reader chose between. */
    readonly scorable: number;
    readonly readings: readonly BoxReadingScore[];
    readonly variants: readonly BoxScore[];
  };
  readonly shaping: readonly ShapingScore[];
}

/* -------------------------------------------------------------------------- */
/* scoring                                                                    */
/* -------------------------------------------------------------------------- */

function median(values: readonly number[]): number {
  if (values.length === 0) return 0;
  const sorted = [...values].sort((a, b) => a - b);
  const middle = sorted.length >> 1;
  const upper = sorted[middle] ?? 0;
  if (sorted.length % 2 === 1) return upper;
  return ((sorted[middle - 1] ?? 0) + upper) / 2;
}

/** The delta as a fraction of what the browser said. */
function relative(reading: number, browser: number): number {
  return Math.abs(reading - browser) / browser;
}

function usable(row: Row): boolean {
  return Number.isFinite(row.browser) && row.browser > 0;
}

function deltaDetail(row: Row, face: RunFace, reading: number, tolerance: number): string {
  const delta = reading - row.browser;
  const units = (delta * face.unitsPerEm) / row.px;
  return (
    `${row.sample} at ${String(row.px)}px: browser ${String(row.browser)}, ` +
    `reader ${String(reading)}, delta ${delta.toExponential(3)} px ` +
    `(${units.toExponential(3)} font units, ` +
    `${relative(reading, row.browser).toExponential(3)} relative), ` +
    `over the ${tolerance.toExponential(3)} px ${Math.abs(delta / tolerance).toFixed(2)} times`
  );
}

interface Comparison {
  readonly face: RunFace;
  readonly row: Row;
  readonly shipped: number;
  readonly tolerance: number;
}

/** Every gated row that carries a shipped reading and a usable browser width. */
function gatedComparisons(run: Run, failures: Failure[]): Comparison[] {
  const byId = new Map(run.samples.map((sample) => [sample.id, sample]));
  const out: Comparison[] = [];
  for (const face of run.faces) {
    for (const row of face.rows) {
      const sample = byId.get(row.sample);
      if (sample === undefined || !sample.gated) continue;
      const shipped = row.readings[SHIPPED];
      if (shipped === undefined) {
        failures.push({
          kind: 'unmeasured',
          subject: face.file,
          detail: `${row.sample} at ${String(row.px)}px carries no ${SHIPPED} reading`,
        });
        continue;
      }
      if (!usable(row)) {
        failures.push({
          kind: 'unmeasured',
          subject: face.file,
          detail:
            `${row.sample} at ${String(row.px)}px: ` +
            `the browser reported ${String(row.browser)}`,
        });
        continue;
      }
      if (row.kerns === undefined) {
        failures.push({
          kind: 'unmeasured',
          subject: face.file,
          detail:
            `${row.sample} at ${String(row.px)}px carries no kern count, and the tolerance ` +
            'is one browser rounding per glyph and one per applied kern',
        });
        continue;
      }
      out.push({
        face,
        row,
        shipped,
        tolerance: widthTolerance(glyphsOf(sample.text), row.kerns),
      });
    }
  }
  return out;
}

/**
 * The regime the tolerance is derived under, asserted rather than assumed.
 *
 * A browser that has stopped rounding each advance to a whole pixel is measuring
 * something this gate has no bound for, and the bound would be four orders
 * looser than that browser needs.
 */
function advanceQuantum(run: Run, failures: Failure[]): void {
  let fractional = 0;
  let first: string | undefined;
  for (const face of run.faces) {
    for (const row of face.rows) {
      if (!usable(row) || Number.isInteger(row.browser / ADVANCE_QUANTUM)) continue;
      fractional += 1;
      first ??= `${face.file} ${row.sample} at ${String(row.px)}px: ${String(row.browser)}`;
    }
  }
  if (fractional === 0) return;
  failures.push({
    kind: 'advance-quantum',
    subject: run.chromium,
    detail:
      `${String(fractional)} browser width(s) are not whole multiples of ` +
      `${String(ADVANCE_QUANTUM)} px, starting at ${String(first)}; the tolerance is ` +
      'derived from per-glyph rounding to that quantum and has to be derived again',
  });
}

function scoreReadings(comparisons: readonly Comparison[]): ReadingScore[] {
  return READINGS.map((reading) => {
    let inside = 0;
    let exact = 0;
    let of = 0;
    let worstOverTolerance = 0;
    for (const { row, tolerance } of comparisons) {
      const value = row.readings[reading];
      if (value === undefined) continue;
      of += 1;
      const delta = Math.abs(value - row.browser);
      if (delta <= tolerance) inside += 1;
      if (delta < EXACT_PX) exact += 1;
      worstOverTolerance = Math.max(worstOverTolerance, delta / tolerance);
    }
    return { reading, inside, exact, of, worstOverTolerance };
  });
}

interface ScriptBucket {
  readonly script: string;
  readonly gated: boolean;
  readonly faces: Set<string>;
  readonly relatives: number[];
  inside: number;
  worstOverTolerance: number;
}

function scoreScripts(run: Run): ScriptScore[] {
  const buckets = new Map<string, ScriptBucket>();
  for (const sample of run.samples) {
    for (const face of run.faces) {
      for (const row of face.rows) {
        const shipped = row.readings[SHIPPED];
        if (row.sample !== sample.id || shipped === undefined || !usable(row)) continue;
        const tolerance = widthTolerance(glyphsOf(sample.text), row.kerns ?? 0);
        const key = `${sample.script}/${String(sample.gated)}`;
        let bucket = buckets.get(key);
        if (bucket === undefined) {
          bucket = {
            script: sample.script,
            gated: sample.gated,
            faces: new Set<string>(),
            relatives: [],
            inside: 0,
            worstOverTolerance: 0,
          };
          buckets.set(key, bucket);
        }
        const delta = Math.abs(shipped - row.browser);
        bucket.faces.add(face.sha256);
        bucket.relatives.push(relative(shipped, row.browser));
        if (delta <= tolerance) bucket.inside += 1;
        bucket.worstOverTolerance = Math.max(bucket.worstOverTolerance, delta / tolerance);
      }
    }
  }
  return [...buckets.values()]
    .map((bucket) => ({
      script: bucket.script,
      gated: bucket.gated,
      faces: bucket.faces.size,
      comparisons: bucket.relatives.length,
      inside: bucket.inside,
      medianRelative: median(bucket.relatives),
      worstOverTolerance: bucket.worstOverTolerance,
    }))
    .sort((a, b) => Number(b.gated) - Number(a.gated) || a.script.localeCompare(b.script));
}

/** The reader's box under one rounding, against the whole pixel the browser reported. */
function fitsBox(round: BoxRounding, face: RunFace, boxPx: number): boolean {
  const scale = boxPx / face.unitsPerEm;
  const { ascent, descent } = round({
    ascent: face.box.reader.ascent * scale,
    descent: face.box.reader.descent * scale,
  });
  return (
    Math.abs(ascent - face.box.browser.ascent) < EXACT_PX &&
    Math.abs(descent - face.box.browser.descent) < EXACT_PX
  );
}

function boxDeltas(face: RunFace, boxPx: number): { ascent: number; descent: number } {
  const scale = boxPx / face.unitsPerEm;
  return {
    ascent: face.box.reader.ascent * scale - face.box.browser.ascent,
    descent: face.box.reader.descent * scale - face.box.browser.descent,
  };
}

/** How far a candidate pair sits from the box the browser reported, in pixels. */
function boxMiss(pair: BoxPair, face: RunFace, boxPx: number): number {
  const scale = boxPx / face.unitsPerEm;
  return Math.max(
    Math.abs(pair.ascent * scale - face.box.browser.ascent),
    Math.abs(pair.descent * scale - face.box.browser.descent),
  );
}

function samePair(a: BoxPair, b: BoxPair): boolean {
  return a.ascent === b.ascent && a.descent === b.descent;
}

/** Every reading of the face box, over the faces whose candidates were recorded. */
function scoreBoxReadings(
  scorable: readonly RunFace[],
  boxPx: number,
  shippedBox: string,
): BoxReadingScore[] {
  return Object.entries(BOX_READINGS).map(([reading, read]) => {
    let of = 0;
    let fits = 0;
    let separates = 0;
    let worstPx = 0;
    for (const face of scorable) {
      const candidates = face.box.candidates;
      if (candidates === undefined) continue;
      const pair = read(candidates);
      if (pair === undefined) continue;
      of += 1;
      const miss = boxMiss(pair, face, boxPx);
      if (miss <= BOX_TOLERANCE) fits += 1;
      worstPx = Math.max(worstPx, miss);
      const shipped = BOX_READINGS[shippedBox]?.(candidates);
      if (shipped !== undefined && !samePair(pair, shipped)) separates += 1;
    }
    return { reading, of, fits, separates, worstPx };
  });
}

function scoreShaping(run: Run): ShapingScore[] {
  const out: ShapingScore[] = [];
  for (const sample of run.samples) {
    if (sample.gated) continue;
    const faces = new Set<string>();
    const ratios: number[] = [];
    let outsideTolerance = 0;
    for (const face of run.faces) {
      for (const row of face.rows) {
        const shipped = row.readings[SHIPPED];
        if (row.sample !== sample.id || shipped === undefined || !usable(row)) continue;
        const tolerance = widthTolerance(glyphsOf(sample.text), row.kerns ?? 0);
        faces.add(face.sha256);
        ratios.push(shipped / row.browser);
        if (Math.abs(shipped - row.browser) > tolerance) outsideTolerance += 1;
      }
    }
    out.push({
      sample: sample.id,
      faces: faces.size,
      medianRatio: median(ratios),
      worstRatio: ratios.reduce((a, b) => Math.max(a, b), 0),
      outsideTolerance,
      of: ratios.length,
    });
  }
  return out;
}

/** A1 to A7, each collecting every offender rather than stopping at the first. */
export function score(run: Run): Verdict {
  const failures: Failure[] = [];

  if (run.directories.length === 0) {
    failures.push({
      kind: 'no-directories',
      subject: '--font-dir',
      detail: 'no font directory was named, so nothing could have been measured',
    });
  }
  if (run.faces.length === 0) {
    failures.push({
      kind: 'no-faces',
      subject: run.directories.join(', '),
      detail:
        `${String(run.indexed.faces)} face(s) were indexed and ` +
        `${String(run.skipped.length)} file(s) skipped; none reached a measurement`,
    });
  }

  advanceQuantum(run, failures);

  const comparisons = gatedComparisons(run, failures);
  if (comparisons.length === 0) {
    failures.push({
      kind: 'no-gated-comparisons',
      subject: run.directories.join(', '),
      detail: 'no gated sample was covered by any face, so a green table would say nothing',
    });
  }

  for (const comparison of comparisons) {
    if (Math.abs(comparison.shipped - comparison.row.browser) <= comparison.tolerance) continue;
    failures.push({
      kind: 'width',
      subject: comparison.face.file,
      detail: deltaDetail(
        comparison.row,
        comparison.face,
        comparison.shipped,
        comparison.tolerance,
      ),
    });
  }

  const scoreboard = scoreReadings(comparisons);
  const shipped = scoreboard.find((entry) => entry.reading === SHIPPED);
  if (shipped !== undefined && comparisons.length > 0) {
    for (const rival of scoreboard) {
      if (rival.reading === SHIPPED || rival.inside <= shipped.inside) continue;
      failures.push({
        kind: 'rival-beats-shipped',
        subject: rival.reading,
        detail:
          `lands inside the tolerance on ${String(rival.inside)}/${String(rival.of)} widths ` +
          `against ${SHIPPED} at ${String(shipped.inside)}/${String(shipped.of)}`,
      });
    }
    const separating = comparisons.filter((comparison) => {
      const unkerned = comparison.row.readings[UNKERNED];
      if (unkerned === undefined) return false;
      const kerned = Math.abs(comparison.shipped - comparison.row.browser) <= comparison.tolerance;
      return kerned && Math.abs(unkerned - comparison.row.browser) > comparison.tolerance;
    });
    if (separating.length === 0) {
      failures.push({
        kind: 'kerning-unseparated',
        subject: UNKERNED,
        detail:
          `no gated row has ${SHIPPED} inside the tolerance and ${UNKERNED} outside it, so ` +
          'either no face here kerns or the reader never does; the kerning path is unmeasured',
      });
    }
  }

  const variants: BoxScore[] = Object.entries(BOX_ROUNDINGS).map(([name, round]) => ({
    name,
    fits: run.faces.filter((face) => fitsBox(round, face, run.boxPx)).length,
    of: run.faces.length,
  }));
  let agreed = 0;
  const missed: string[] = [];
  for (const face of run.faces) {
    const delta = boxDeltas(face, run.boxPx);
    if (Math.abs(delta.ascent) <= BOX_TOLERANCE && Math.abs(delta.descent) <= BOX_TOLERANCE) {
      agreed += 1;
      continue;
    }
    const scale = run.boxPx / face.unitsPerEm;
    // The browser's whole pixels back in font units, which is the pair some
    // table of this face has to carry for any reading to fit it.
    const wants = {
      ascent: face.box.browser.ascent / scale,
      descent: face.box.browser.descent / scale,
    };
    missed.push(
      `${face.file}: browser ${String(face.box.browser.ascent)}/` +
        `${String(face.box.browser.descent)} px is ${wants.ascent.toFixed(1)}/` +
        `${wants.descent.toFixed(1)} units over ${String(face.unitsPerEm)} per em, where ` +
        `${face.metricsSource} says ${String(face.box.reader.ascent)}/` +
        `${String(face.box.reader.descent)}; out by ` +
        `${delta.ascent.toExponential(3)}/${delta.descent.toExponential(3)} px`,
    );
  }

  const shippedBoxName = shippedBoxFor(run.image.platform);
  const readings = scoreBoxReadings(run.faces, run.boxPx, shippedBoxName);
  const scorable = run.faces.filter((face) => face.box.candidates !== undefined);
  const unscorable = run.faces.filter((face) => face.box.candidates === undefined);
  if (unscorable.length > 0) {
    failures.push({
      kind: 'unmeasured',
      subject: 'FaceMetrics.candidates',
      detail:
        `${String(unscorable.length)} of ${String(run.faces.length)} measured face(s) carry no ` +
        `candidate pairs, starting at ${String(unscorable[0]?.file)}; which table the browser ` +
        'read the box out of cannot be scored on a run that did not record them',
    });
  }
  const shippedBox = readings.find((entry) => entry.reading === shippedBoxName);
  if (shippedBox !== undefined && run.faces.length > 0) {
    if (agreed < run.faces.length && !readings.some((e) => e.fits === run.faces.length)) {
      failures.push({
        kind: 'face-box',
        subject: run.image.platform,
        detail:
          `no reading of the face box fits all ${String(run.faces.length)} face(s) - ` +
          `${readings.map((e) => `${e.reading} ${String(e.fits)}/${String(e.of)}`).join('; ')} - ` +
          `and ${String(run.faces.length - agreed)} face(s) sit over the ` +
          `${String(BOX_TOLERANCE)} px one rounding can cost: ${missed.join(' | ')}`,
      });
    }
    for (const rival of readings) {
      if (rival.reading === shippedBoxName || rival.fits <= shippedBox.fits) continue;
      failures.push({
        kind: 'face-box-rival',
        subject: rival.reading,
        detail:
          `fits ${String(rival.fits)}/${String(rival.of)} face box(es) against ` +
          `${shippedBoxName} at ${String(shippedBox.fits)}/${String(shippedBox.of)}, and ` +
          `predicts a different box on ${String(rival.separates)} of them`,
      });
    }
  }

  const shaping = scoreShaping(run);
  for (const tier of shaping) {
    if (tier.faces === 0 || tier.outsideTolerance > 0) continue;
    failures.push({
      kind: 'shaping-invisible',
      subject: tier.sample,
      detail:
        `${String(tier.faces)} face(s) cover it and not one of ${String(tier.of)} row(s) is ` +
        'outside the tolerance a whole-pixel browser already costs; a recorded sample that ' +
        'measures like a gated one is pricing nothing',
    });
  }

  return {
    failures,
    scoreboard,
    byScript: scoreScripts(run),
    faceBox: {
      agreed,
      of: run.faces.length,
      scorable: scorable.length,
      readings,
      variants,
    },
    shaping,
  };
}

/* -------------------------------------------------------------------------- */
/* the report                                                                 */
/* -------------------------------------------------------------------------- */

/** `1/65536 + 1/131072` reads as a step; its decimal expansion does not. */
const READER_STEP_TEXT = '1/65536 + 1/131072';

function table(header: readonly string[], rows: readonly (readonly string[])[]): string {
  const lines = [`| ${header.join(' | ')} |`, `| ${header.map(() => '---').join(' | ')} |`];
  for (const row of rows) lines.push(`| ${row.join(' | ')} |`);
  return lines.join('\n');
}

export function reportMarkdown(run: Run, verdict: Verdict): string {
  const parts: string[] = [];
  const image =
    run.image.imageOs === null
      ? run.image.platform
      : `${run.image.imageOs} ${String(run.image.imageVersion)}`;

  parts.push(`## ${run.experiment} - the font reader against Chromium ${run.chromium}`);
  parts.push(
    `${String(run.indexed.measured)} of ${String(run.indexed.faces)} indexed face(s) from ` +
      `${run.directories.join(', ')} on ${image} (${run.image.platform} ${run.image.release}), ` +
      `${String(run.skipped.length)} file(s) skipped.`,
  );
  parts.push(
    `**The gate.** This browser quantises every glyph advance to ${String(ADVANCE_QUANTUM)} px, ` +
      'so a reader that sums fractional advances can never reproduce it and is not asked to. ' +
      `A width may sit ${String(ADVANCE_QUANTUM / 2)} px away from the browser's per glyph ` +
      'and once more per kern the reader applied, because the browser rounds the advance and ' +
      `the adjustment separately, plus the reader's own ${READER_STEP_TEXT} px per glyph, and ` +
      'the face box one rounding. A single joint rounding is refuted rather than assumed: it ' +
      'could not cost the 0.648 px a glyph T14 measured on Lato Thin, where a kern-free sample ' +
      'on the same faces never passes 0.45. Nothing here is a tolerance chosen to pass. ADR 0045.',
  );

  parts.push('### Every reading, scored against the tolerance and at equality');
  parts.push(
    table(
      ['reading', 'inside', 'exact', 'of', 'worst / tolerance'],
      verdict.scoreboard.map((entry) => [
        entry.reading === SHIPPED ? `**${entry.reading}**` : entry.reading,
        String(entry.inside),
        String(entry.exact),
        String(entry.of),
        entry.worstOverTolerance.toExponential(3),
      ]),
    ),
  );
  parts.push(
    '**What this run cannot see.** The `exact` column is the browser rounding, not the reader: ' +
      'at a whole-pixel quantum the four readings that differ only below 1/65536 px - ' +
      '`shipped`, `exact`, `advanceRounded`, `wholeStringTruncated` - cannot be told apart, ' +
      'and T13 remains the only experiment that separates them. Only `unkerned` is separable ' +
      'here, because a missing kern is a fraction of the width rather than of a pixel.',
  );

  parts.push('### By script');
  parts.push(
    table(
      ['script', 'gated', 'faces', 'comparisons', 'inside', 'median relative', 'worst / tolerance'],
      verdict.byScript.map((entry) => [
        entry.script,
        entry.gated ? 'yes' : 'no',
        String(entry.faces),
        String(entry.comparisons),
        String(entry.inside),
        entry.medianRelative.toExponential(3),
        entry.worstOverTolerance.toExponential(3),
      ]),
    ),
  );

  parts.push('### The face box');
  parts.push(
    `${String(verdict.faceBox.agreed)} of ${String(verdict.faceBox.of)} face(s) inside ` +
      `${String(BOX_TOLERANCE)} px, and ${String(verdict.faceBox.scorable)} carrying the ` +
      'candidate pairs the question below is asked of.',
  );
  parts.push('#### Which table the box came out of');
  parts.push(
    table(
      ['reading', 'fits', 'of', 'differs from shipped', 'worst px'],
      verdict.faceBox.readings.map((entry) => [
        entry.reading === shippedBoxFor(run.image.platform)
          ? `**${entry.reading}**`
          : entry.reading,
        String(entry.fits),
        String(entry.of),
        String(entry.separates),
        entry.worstPx.toExponential(3),
      ]),
    ),
  );
  parts.push(
    `This runner reads the box through ${rasteriserOf(run.image.platform)}, ` +
      `so **${shippedBoxFor(run.image.platform)}** is the shipped reading here; the other ` +
      'rasteriser reads a different table and is scored by T13. ADR 0045.',
  );
  parts.push(
    'A reading is scored over the faces whose tables carry the pair it names, and fits a face ' +
      `whose box it lands on within the ${String(BOX_TOLERANCE)} px one rounding can cost. ` +
      'The `differs from shipped` column is what carries the information: where every table of ' +
      'a face agrees with itself, so does every reading, and the row says nothing.',
  );
  parts.push('#### Which rounding the browser applied');
  parts.push(
    table(
      ['reading', 'fits', 'of'],
      verdict.faceBox.variants.map((entry) => [entry.name, String(entry.fits), String(entry.of)]),
    ),
  );
  parts.push(
    'That table scores the browser, not the reader: a box that lands on a whole pixel is ' +
      'reproduced by every rounding, so only the faces whose tables put it off one - a 2000 em ' +
      'on the half, a 2048 em elsewhere - separate the rows. T13 settled round half up on ' +
      'DirectWrite, 28 of 28. ADR 0052.',
  );

  parts.push('### What the missing features cost');
  parts.push(
    verdict.shaping.length === 0
      ? 'No recorded sample was declared.'
      : table(
          ['sample', 'faces', 'median reader/browser', 'worst', 'outside tolerance', 'of'],
          verdict.shaping.map((tier) => [
            tier.sample,
            String(tier.faces),
            tier.faces === 0 ? 'no face covers it' : tier.medianRatio.toFixed(6),
            tier.faces === 0 ? '-' : tier.worstRatio.toFixed(6),
            String(tier.outsideTolerance),
            String(tier.of),
          ]),
        ),
  );

  if (verdict.failures.length === 0) {
    parts.push('**Green.** Every gate held.');
  } else {
    parts.push(`### ${String(verdict.failures.length)} failure(s)`);
    parts.push(
      table(
        ['gate', 'subject', 'detail'],
        verdict.failures.map((failure) => [failure.kind, failure.subject, failure.detail]),
      ),
    );
  }

  parts.push('The rule is committed in tools/ground-truth/fonts/real-faces/probes.ts.');
  return `${parts.join('\n\n')}\n`;
}
