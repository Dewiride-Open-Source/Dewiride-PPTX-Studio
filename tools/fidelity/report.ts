/**
 * What a run writes down: a machine-readable score and a table for a reviewer.
 *
 * The table leads with the worst slide and the count of slides that changed,
 * not with the corpus mean. A mean over 149 slides and 1.2 million cells cannot
 * move for a defect on one of them - all the text vanishing from a slide is
 * worth about three basis points - so putting the mean first would train
 * everyone to ignore the report. ADR 0035.
 */

import type { Blame } from './blame.ts';
import { agreementBp, type SlideScore } from './metric/score.ts';

export interface SlideResult {
  readonly key: string;
  readonly deck: string;
  readonly slide: number;
  readonly score: SlideScore;
  readonly rasterSha256: string;
  readonly svgSha256: string;
  /** Set when the recorded digest for this slide disagrees with this run. */
  readonly changed: boolean;
  /** The difference split over the shapes that painted it, worst first. */
  readonly blame: readonly Blame[];
}

export interface RunReport {
  readonly envId: string;
  readonly slides: readonly SlideResult[];
  /** The worst PowerPoint disagreed with itself, from the oracle fixture. */
  readonly noiseFloorBp: number;
  readonly noiseFloorMaxD: number;
  /** Faces the corpus draws with that this machine does not have. */
  readonly substituted: readonly string[];
}

/** Cells that agree, over all slides, as basis points. */
export function corpusMeanBp(slides: readonly SlideResult[]): number {
  if (slides.length === 0) return 10000;
  let sumD = 0;
  let cells = 0;
  for (const slide of slides) {
    sumD += slide.score.sumD;
    cells += slide.score.cells;
  }
  return agreementBp(sumD, cells);
}

/**
 * The scores file.
 *
 * Key order is fixed and nothing dated or host-specific goes in it, so two runs
 * of the same tree produce the same bytes and a diff of this file is a diff of
 * the renderer.
 */
export function scoresJson(report: RunReport): string {
  const slides = [...report.slides].sort((a, b) => (a.key < b.key ? -1 : a.key > b.key ? 1 : 0));
  return `${JSON.stringify(
    {
      envId: report.envId,
      slides: slides.length,
      corpusMeanBp: corpusMeanBp(slides),
      noiseFloorBp: report.noiseFloorBp,
      noiseFloorMaxD: report.noiseFloorMaxD,
      substituted: report.substituted,
      perSlide: slides.map((slide) => ({
        key: slide.key,
        meanBp: slide.score.meanBp,
        maxD: slide.score.maxD,
        hist: slide.score.hist,
        blame: slide.blame.slice(0, BLAMED_PER_SLIDE),
        rasterSha256: slide.rasterSha256,
        svgSha256: slide.svgSha256,
      })),
    },
    null,
    2,
  )}\n`;
}

/** How many shapes a slide is broken down into, in the report and the file. */
const BLAMED_PER_SLIDE = 3;

/** How far down the worst-first list the shape breakdown is printed. */
const LOCALISED_SLIDES = 10;

/** How many shapes the corpus-wide ranking names. */
const WORST_SHAPES = 12;

interface BlamedShape {
  readonly key: string;
  readonly blamed: Blame;
  /** Agreement over the shape's own area, so size does not decide the rank. */
  readonly meanBp: number;
}

/**
 * The shapes drawn furthest from PowerPoint, over their own area.
 *
 * Ranked per cell the shape owns rather than by total difference, because a
 * slide-sized diagram nobody has built yet outweighs every real defect in the
 * corpus put together and would be the whole table otherwise.
 */
function worstShapes(slides: readonly SlideResult[], limit: number): readonly BlamedShape[] {
  const rows: BlamedShape[] = [];
  for (const slide of slides) {
    for (const blamed of slide.blame) {
      rows.push({ key: slide.key, blamed, meanBp: agreementBp(blamed.sumD, blamed.covered) });
    }
  }
  rows.sort(
    (a, b) =>
      a.meanBp - b.meanBp ||
      b.blamed.maxD - a.blamed.maxD ||
      (a.key < b.key ? -1 : a.key > b.key ? 1 : 0),
  );
  return rows.slice(0, limit);
}

function worstFirst(slides: readonly SlideResult[]): readonly SlideResult[] {
  return [...slides].sort((a, b) => a.score.meanBp - b.score.meanBp || (a.key < b.key ? -1 : 1));
}

export function reportMarkdown(report: RunReport): string {
  const changed = report.slides.filter((slide) => slide.changed);
  const worst = worstFirst(report.slides).slice(0, 10);
  const lines: string[] = [];

  lines.push(
    `## Fidelity — ${report.envId} — ${String(report.slides.length)} slides — ` +
      `GATE: ${changed.length === 0 ? 'PASS' : 'FAIL'}`,
    '',
    '| | |',
    '| --- | --- |',
    `| slides whose raster changed | ${String(changed.length)} |`,
    `| worst slide | ${worst[0]?.key ?? '-'} at ${String(worst[0]?.score.meanBp ?? 10000)} bp |`,
    `| corpus mean (provenance, not signal) | ${String(corpusMeanBp(report.slides))} bp |`,
    `| oracle noise floor | ${String(report.noiseFloorBp)} bp, maxD ${String(report.noiseFloorMaxD)} |`,
    '',
  );

  if (changed.length > 0) {
    lines.push(
      '### Slides whose raster changed',
      '',
      '| slide | meanBp | maxD |',
      '| --- | ---: | ---: |',
    );
    for (const slide of changed) {
      lines.push(`| ${slide.key} | ${String(slide.score.meanBp)} | ${String(slide.score.maxD)} |`);
    }
    lines.push('');
  }

  lines.push(
    '### Furthest from PowerPoint',
    '',
    '| slide | meanBp | maxD | cells differing |',
    '| --- | ---: | ---: | ---: |',
  );
  for (const slide of worst) {
    const differing = slide.score.cells - (slide.score.hist[0] ?? 0);
    lines.push(
      `| ${slide.key} | ${String(slide.score.meanBp)} | ${String(slide.score.maxD)} | ${String(differing)} |`,
    );
  }

  lines.push(
    '',
    '### Where the difference is',
    '',
    'Every differing cell, attributed to the shape that painted it last.',
    '',
    '| slide | shape | cells | maxD | share |',
    '| --- | --- | ---: | ---: | ---: |',
  );
  for (const slide of worst.slice(0, LOCALISED_SLIDES)) {
    for (const blamed of slide.blame.slice(0, BLAMED_PER_SLIDE)) {
      const share = Math.round((1000 * blamed.sumD) / slide.score.sumD) / 10;
      lines.push(
        `| ${slide.key} | ${blamed.name} | ${String(blamed.cells)} | ` +
          `${String(blamed.maxD)} | ${share.toFixed(1)}% |`,
      );
    }
  }

  lines.push(
    '',
    '### Worst-drawn shapes in the corpus',
    '',
    "Agreement over each shape's own area, so a small shape drawn badly outranks a",
    'large one drawn nearly right.',
    '',
    '| shape | slide | meanBp | maxD | cells |',
    '| --- | --- | ---: | ---: | ---: |',
  );
  for (const row of worstShapes(report.slides, WORST_SHAPES)) {
    lines.push(
      `| ${row.blamed.name} | ${row.key} | ${String(row.meanBp)} | ` +
        `${String(row.blamed.maxD)} | ${String(row.blamed.cells)}/${String(row.blamed.covered)} |`,
    );
  }

  if (report.substituted.length > 0) {
    lines.push(
      '',
      '### Faces this machine does not have',
      '',
      'Text in these was drawn in a substitute, so its score measures two different substitutions.',
      '',
      report.substituted.map((family) => `- ${family}`).join('\n'),
    );
  }

  return `${lines.join('\n')}\n`;
}
