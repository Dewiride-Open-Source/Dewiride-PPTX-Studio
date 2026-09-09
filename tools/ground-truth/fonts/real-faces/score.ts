/**
 * Experiment T14 - the gate, as a pure function over one measurement.
 *
 * Split out of `tools/ground-truth/fonts/real-faces/analyse.ts` because this
 * experiment's analysis is a CI gate, and a gate that has never been shown to
 * go red is not a gate. It touches no font, no browser and no filesystem, so a
 * maintainer can re-score the artifact on a machine with no fonts at all.
 */

import {
  AGREEMENT,
  BOX_VARIANTS,
  EXACT_PX,
  READINGS,
  SHIPPED,
  applyBoxVariant,
  type BoxVariant,
  type Reading,
} from './probes.ts';

/* -------------------------------------------------------------------------- */
/* the measurement                                                            */
/* -------------------------------------------------------------------------- */

export interface RunSample {
  readonly id: string;
  readonly script: string;
  readonly gated: boolean;
  readonly shaped: boolean;
  readonly asks: string;
  readonly text: string;
}

export interface Row {
  readonly sample: string;
  readonly px: number;
  /** `measureText(...).width`, in CSS pixels. */
  readonly browser: number;
  readonly readings: Readonly<Partial<Record<Reading, number>>>;
}

export interface BoxPair {
  readonly ascent: number;
  readonly descent: number;
}

export interface RunFace {
  readonly file: string;
  readonly sha256: string;
  readonly family: string;
  readonly subfamily: string;
  readonly unitsPerEm: number;
  readonly metricsSource: string;
  /** The browser's box in pixels at `boxPx`; the reader's in font units. */
  readonly box: { readonly browser: BoxPair; readonly reader: BoxPair };
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
  readonly agreement: number;
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
  | 'agreement'
  | 'rival-beats-shipped'
  | 'face-box'
  | 'shaping-invisible';

export interface Failure {
  readonly kind: FailureKind;
  /** The font file, the reading, or the sample. */
  readonly subject: string;
  /** Both widths, the delta in px and in font units. */
  readonly detail: string;
}

export interface ReadingScore {
  readonly reading: string;
  readonly exact: number;
  readonly of: number;
  readonly worstRelative: number;
}

export interface ScriptScore {
  readonly script: string;
  readonly gated: boolean;
  readonly faces: number;
  readonly comparisons: number;
  readonly exact: number;
  readonly medianRelative: number;
  readonly worstRelative: number;
}

export interface BoxScore {
  readonly name: string;
  readonly fits: number;
  readonly of: number;
}

export interface ShapingScore {
  readonly sample: string;
  readonly faces: number;
  readonly medianOver: number;
  readonly worstOver: number;
}

export interface Verdict {
  /** Empty means green. */
  readonly failures: readonly Failure[];
  readonly scoreboard: readonly ReadingScore[];
  readonly byScript: readonly ScriptScore[];
  readonly faceBox: { readonly variants: readonly BoxScore[]; readonly answer: string | null };
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

function deltaDetail(row: Row, face: RunFace, reading: number): string {
  const delta = reading - row.browser;
  const units = (delta * face.unitsPerEm) / row.px;
  return (
    `${row.sample} at ${String(row.px)}px: browser ${String(row.browser)}, ` +
    `reader ${String(reading)}, delta ${delta.toExponential(3)} px ` +
    `(${units.toExponential(3)} font units, ` +
    `${relative(reading, row.browser).toExponential(3)} relative)`
  );
}

interface Comparison {
  readonly face: RunFace;
  readonly row: Row;
  readonly shipped: number;
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
          kind: 'agreement',
          subject: face.file,
          detail: `${row.sample} at ${String(row.px)}px carries no ${SHIPPED} reading`,
        });
        continue;
      }
      if (!Number.isFinite(row.browser) || row.browser <= 0) {
        failures.push({
          kind: 'agreement',
          subject: face.file,
          detail:
            `${row.sample} at ${String(row.px)}px: ` +
            `the browser reported ${String(row.browser)}`,
        });
        continue;
      }
      out.push({ face, row, shipped });
    }
  }
  return out;
}

function scoreReadings(comparisons: readonly Comparison[]): ReadingScore[] {
  return READINGS.map((reading) => {
    let exact = 0;
    let of = 0;
    let worstRelative = 0;
    for (const { row } of comparisons) {
      const value = row.readings[reading];
      if (value === undefined) continue;
      of += 1;
      if (Math.abs(value - row.browser) < EXACT_PX) exact += 1;
      worstRelative = Math.max(worstRelative, relative(value, row.browser));
    }
    return { reading, exact, of, worstRelative };
  });
}

interface ScriptBucket {
  readonly script: string;
  readonly gated: boolean;
  readonly faces: Set<string>;
  readonly relatives: number[];
  exact: number;
}

function scoreScripts(run: Run): ScriptScore[] {
  const byId = new Map(run.samples.map((sample) => [sample.id, sample]));
  const buckets = new Map<string, ScriptBucket>();
  for (const face of run.faces) {
    for (const row of face.rows) {
      const sample = byId.get(row.sample);
      const shipped = row.readings[SHIPPED];
      if (sample === undefined || shipped === undefined || row.browser <= 0) continue;
      const key = `${sample.script}/${String(sample.gated)}`;
      let bucket = buckets.get(key);
      if (bucket === undefined) {
        bucket = {
          script: sample.script,
          gated: sample.gated,
          faces: new Set<string>(),
          relatives: [],
          exact: 0,
        };
        buckets.set(key, bucket);
      }
      bucket.faces.add(face.sha256);
      bucket.relatives.push(relative(shipped, row.browser));
      if (Math.abs(shipped - row.browser) < EXACT_PX) bucket.exact += 1;
    }
  }
  return [...buckets.values()]
    .map((bucket) => ({
      script: bucket.script,
      gated: bucket.gated,
      faces: bucket.faces.size,
      comparisons: bucket.relatives.length,
      exact: bucket.exact,
      medianRelative: median(bucket.relatives),
      worstRelative: bucket.relatives.reduce((a, b) => Math.max(a, b), 0),
    }))
    .sort((a, b) => Number(b.gated) - Number(a.gated) || a.script.localeCompare(b.script));
}

/** The reader's box under one rounding, against what the browser reported. */
function fitsBox(variant: BoxVariant, face: RunFace, boxPx: number): boolean {
  const scale = boxPx / face.unitsPerEm;
  const ascent = applyBoxVariant(variant, face.box.reader.ascent * scale);
  const descent = applyBoxVariant(variant, face.box.reader.descent * scale);
  return (
    Math.abs(ascent - face.box.browser.ascent) < EXACT_PX &&
    Math.abs(descent - face.box.browser.descent) < EXACT_PX
  );
}

function scoreShaping(run: Run): ShapingScore[] {
  const out: ShapingScore[] = [];
  for (const sample of run.samples) {
    if (!sample.shaped) continue;
    const faces = new Set<string>();
    const overs: number[] = [];
    for (const face of run.faces) {
      for (const row of face.rows) {
        const shipped = row.readings[SHIPPED];
        if (row.sample !== sample.id || shipped === undefined || row.browser <= 0) continue;
        faces.add(face.sha256);
        overs.push(shipped / row.browser - 1);
      }
    }
    out.push({
      sample: sample.id,
      faces: faces.size,
      medianOver: median(overs),
      worstOver: overs.reduce((a, b) => Math.max(a, b), 0),
    });
  }
  return out;
}

/** A1 to A6, each collecting every offender rather than stopping at the first. */
export function score(run: Run): Verdict {
  const failures: Failure[] = [];
  const limit = run.agreement;

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

  const comparisons = gatedComparisons(run, failures);
  if (comparisons.length === 0) {
    failures.push({
      kind: 'no-gated-comparisons',
      subject: run.directories.join(', '),
      detail: 'no gated sample was covered by any face, so a green table would say nothing',
    });
  }

  for (const comparison of comparisons) {
    if (relative(comparison.shipped, comparison.row.browser) <= limit) continue;
    failures.push({
      kind: 'agreement',
      subject: comparison.face.file,
      detail:
        `${deltaDetail(comparison.row, comparison.face, comparison.shipped)}, ` +
        `over ${limit.toExponential(0)}`,
    });
  }

  const scoreboard = scoreReadings(comparisons);
  const shipped = scoreboard.find((entry) => entry.reading === SHIPPED);
  if (shipped !== undefined && comparisons.length > 0) {
    const rivals = scoreboard.filter((entry) => entry.reading !== SHIPPED && entry.of > 0);
    const best = rivals.reduce((top, entry) => Math.max(top, entry.exact), -1);
    for (const rival of rivals) {
      if (best < shipped.exact || rival.exact !== best) continue;
      failures.push({
        kind: 'rival-beats-shipped',
        subject: rival.reading,
        detail:
          `reproduces ${String(rival.exact)}/${String(rival.of)} widths exactly against ` +
          `${SHIPPED} at ${String(shipped.exact)}/${String(shipped.of)}; the shipped reading ` +
          'must be alone at the top or these faces do not separate the two',
      });
    }
  }

  const variants: BoxScore[] = BOX_VARIANTS.map((variant) => ({
    name: variant,
    fits: run.faces.filter((face) => fitsBox(variant, face, run.boxPx)).length,
    of: run.faces.length,
  }));
  const perfect = variants.filter((variant) => variant.of > 0 && variant.fits === variant.of);
  if (run.faces.length > 0 && perfect.length === 0) {
    for (const face of run.faces) {
      if (BOX_VARIANTS.some((variant) => fitsBox(variant, face, run.boxPx))) continue;
      const scale = run.boxPx / face.unitsPerEm;
      failures.push({
        kind: 'face-box',
        subject: face.file,
        detail:
          `browser ${String(face.box.browser.ascent)}/${String(face.box.browser.descent)}, ` +
          `reader ${String(face.box.reader.ascent * scale)}/` +
          `${String(face.box.reader.descent * scale)} from ${face.metricsSource} over ` +
          `${String(face.unitsPerEm)} units per em; no exact, round, floor or ceil ` +
          'reading reproduces both',
      });
    }
  }

  const shaping = scoreShaping(run);
  for (const tier of shaping) {
    if (tier.faces === 0 || tier.worstOver > limit) continue;
    failures.push({
      kind: 'shaping-invisible',
      subject: tier.sample,
      detail:
        `${String(tier.faces)} face(s) cover it and the worst over-measurement is ` +
        `${tier.worstOver.toExponential(3)}, inside the ${limit.toExponential(0)} gate; ` +
        'a shaped script that measures like Latin is not exercising a shaper',
    });
  }

  return {
    failures,
    scoreboard,
    byScript: scoreScripts(run),
    faceBox: { variants, answer: perfect.length === 1 ? (perfect[0]?.name ?? null) : null },
    shaping,
  };
}

/* -------------------------------------------------------------------------- */
/* the report                                                                 */
/* -------------------------------------------------------------------------- */

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
      `${String(run.skipped.length)} file(s) skipped. ` +
      `Gate: ${run.agreement.toExponential(0)} relative.`,
  );

  parts.push('### Every reading, scored at exact equality');
  parts.push(
    table(
      ['reading', 'exact', 'of', 'worst relative'],
      verdict.scoreboard.map((entry) => [
        entry.reading === SHIPPED ? `**${entry.reading}**` : entry.reading,
        String(entry.exact),
        String(entry.of),
        entry.worstRelative.toExponential(3),
      ]),
    ),
  );

  parts.push('### By script');
  parts.push(
    table(
      ['script', 'gated', 'faces', 'comparisons', 'exact', 'median relative', 'worst relative'],
      verdict.byScript.map((entry) => [
        entry.script,
        entry.gated ? 'yes' : 'no',
        String(entry.faces),
        String(entry.comparisons),
        String(entry.exact),
        entry.medianRelative.toExponential(3),
        entry.worstRelative.toExponential(3),
      ]),
    ),
  );

  parts.push('### The face box');
  parts.push(
    table(
      ['reading', 'fits', 'of'],
      verdict.faceBox.variants.map((entry) => [entry.name, String(entry.fits), String(entry.of)]),
    ),
  );
  parts.push(
    verdict.faceBox.answer === null
      ? 'Not separated by these faces: either more than one reading fits every box, or none does.'
      : `The face box is the **${verdict.faceBox.answer}** reading of the measured source.`,
  );

  parts.push('### What the missing shaper costs');
  parts.push(
    verdict.shaping.length === 0
      ? 'No shaped sample was declared.'
      : table(
          ['sample', 'faces', 'median over', 'worst over'],
          verdict.shaping.map((tier) => [
            tier.sample,
            String(tier.faces),
            tier.faces === 0 ? 'no face covers it' : tier.medianOver.toExponential(3),
            tier.faces === 0 ? '-' : tier.worstOver.toExponential(3),
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

  parts.push(
    `The ${String(AGREEMENT)} threshold is committed in ` +
      'tools/ground-truth/fonts/real-faces/probes.ts.',
  );
  return `${parts.join('\n\n')}\n`;
}
