/**
 * What a run writes down: a machine-readable score and a table for a reviewer.
 *
 * The table leads with the worst slide and the count of slides that changed,
 * not with the corpus mean. A mean over 149 slides and 1.2 million cells cannot
 * move for a defect on one of them - all the text vanishing from a slide is
 * worth about three basis points - so putting the mean first would train
 * everyone to ignore the report. ADR 0035.
 */

import type { SlideScore } from './metric/score.ts';

export interface SlideResult {
  readonly key: string;
  readonly deck: string;
  readonly slide: number;
  readonly score: SlideScore;
  readonly rasterSha256: string;
  readonly svgSha256: string;
  /** Set when the recorded digest for this slide disagrees with this run. */
  readonly changed: boolean;
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
  return 10000 - Math.floor((20000 * sumD + 255 * cells) / (510 * cells));
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
        rasterSha256: slide.rasterSha256,
        svgSha256: slide.svgSha256,
      })),
    },
    null,
    2,
  )}\n`;
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
