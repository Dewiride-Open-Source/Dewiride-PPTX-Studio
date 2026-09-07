/**
 * Gate 2's verification: the nine features, drawn beside PowerPoint's own.
 *
 * The gate asks for a side-by-side with a per-slide diff heatmap. The plan
 * names LibreOffice as the reference; this asks PowerPoint instead, through the
 * oracle sub-phase 3.9 already commits, and says so on the record in ADR 0038.
 * The substitution is not a softening: LibreOffice is a second implementation
 * with opinions of its own, and PowerPoint wrote the format.
 *
 * What is gated has no tolerance in it. Every feature Gate 2 names must be
 * carried by at least one committed slide, every one of those slides must draw
 * without the renderer refusing, and every one must have an oracle to be
 * compared against. The scores are reported beside them and gate nothing, for
 * the reason 3.9 gives: the two sides differ by a few levels everywhere, and a
 * pass/fail on that needs a tolerance, and a tolerance is what gets loosened.
 */

import { readFileSync } from 'node:fs';
import { join } from 'node:path';

import type { Page } from 'playwright';

import { FidelityError } from '../fidelity/errors.ts';
import { gridSvg, heatmapSvg } from '../fidelity/heatmap.ts';
import { decodeGridSet } from '../fidelity/metric/grid.ts';
import type { Grid } from '../fidelity/metric/reduce.ts';
import { differenceOf, regionsOf, scoreOf, type Region } from '../fidelity/metric/score.ts';
import { openHarness } from '../fidelity/raster/browser.ts';
import { injectHarness, renderSlide, slideMarkup } from '../fidelity/raster/render.ts';
import { fidelityProbes, slideKey } from '../ground-truth/render/fidelity/probes.ts';
import { repoPath } from '../repo/root.ts';

import { GATE2_FEATURES, MARKUP_PATTERNS, type SlideFacts } from './features.ts';

const ORACLE = repoPath('corpus/ground-truth/render/fidelity');

/** The floor a region has to reach to be worth ranking, in levels of 255. */
const REGION_FLOOR = 16;

export interface GateOptions {
  readonly out: string;
  /** A directory of PowerPoint's own PNGs, when one has just been captured. */
  readonly capture: string | null;
}

export interface SlideReport {
  readonly key: string;
  readonly deck: string;
  readonly slide: number;
  /** Gate feature keys this slide demonstrates. */
  readonly features: readonly string[];
  readonly meanBp: number;
  readonly maxD: number;
  readonly hist: readonly number[];
  readonly regions: readonly Region[];
  /** Our renderer's SVG, at full resolution. */
  readonly oursSvg: string;
  /** PowerPoint's, reconstructed from the committed grid it is scored against. */
  readonly theirsSvg: string;
  /** PowerPoint's PNG, when `--capture` pointed at one. */
  readonly theirsPng: string | null;
  readonly heatSvg: string;
  readonly cell: number;
}

export interface FeatureCoverage {
  readonly key: string;
  readonly label: string;
  readonly slides: readonly string[];
}

export interface GateRun {
  readonly slides: readonly SlideReport[];
  readonly coverage: readonly FeatureCoverage[];
  readonly notDrawn: readonly { readonly key: string; readonly reason: string }[];
  readonly scanned: number;
  readonly ok: boolean;
}

/* ------------------------------------------------------------ the detection */

interface RawFacts {
  readonly slide: number;
  readonly matched: readonly string[];
  readonly groupDepth: number;
  readonly transformedGroups: number;
}

/**
 * One deck's slides, measured against the markup patterns.
 *
 * Read from the part bytes rather than from our model, so a feature the parser
 * silently dropped still counts as present and still has to be drawn.
 */
async function scanDeck(
  page: Page,
  deckUrl: string,
  patterns: readonly { key: string; pattern: string }[],
): Promise<readonly RawFacts[]> {
  return await page.evaluate(
    async ({ url, rules }) => {
      const api = (
        globalThis as unknown as {
          pptx: {
            opc: {
              PartStore: { open: (bytes: Uint8Array) => { read: (p: string) => Uint8Array } };
            };
            model: {
              loadDocument: (store: unknown) => { slides: readonly { partName: string }[] };
            };
          };
        }
      ).pptx;
      const bytes = new Uint8Array(await (await fetch(url)).arrayBuffer());
      const store = api.opc.PartStore.open(bytes);
      const document_ = api.model.loadDocument(store);
      const decoder = new TextDecoder();
      const out: RawFacts[] = [];

      for (let index = 0; index < document_.slides.length; index++) {
        const sheet = document_.slides[index];
        if (sheet === undefined) continue;
        const xml = decoder.decode(store.read(sheet.partName));

        const matched: string[] = [];
        for (const rule of rules) {
          if (new RegExp(rule.pattern).test(xml)) matched.push(rule.key);
        }

        // Counted from the open and close tokens, because nesting is not a
        // thing a regular expression can answer and a loose one would call two
        // sibling groups a nest.
        let depth = 0;
        let deepest = 0;
        for (const token of xml.match(/<\/?p:grpSp[\s>/]/g) ?? []) {
          if (token.startsWith('</')) depth -= 1;
          else {
            depth += 1;
            if (depth > deepest) deepest = depth;
          }
        }
        const transformed = (
          xml.match(/<p:grpSpPr>\s*<a:xfrm[^>]*(?:rot="(?!0")|flip[HV]="1")/g) ?? []
        ).length;

        out.push({
          slide: index + 1,
          matched,
          groupDepth: deepest,
          transformedGroups: transformed,
        });
      }
      return out;
    },
    { url: deckUrl, rules: patterns },
  );
}

/* ----------------------------------------------------------------- the run */

export async function runGate2(options: GateOptions): Promise<GateRun> {
  const probes = fidelityProbes();
  const harness = await openHarness();
  await injectHarness(harness.page);

  const facts: SlideFacts[] = [];
  try {
    for (const probe of probes) {
      for (const raw of await scanDeck(harness.page, `/${probe.path}`, MARKUP_PATTERNS)) {
        facts.push({ ...raw, key: slideKey(probe.id, raw.slide), deck: probe.id });
      }
    }

    const coverage: FeatureCoverage[] = GATE2_FEATURES.map((feature) => ({
      key: feature.key,
      label: feature.label,
      slides: facts.filter((slide) => feature.holds(slide)).map((slide) => slide.key),
    }));

    // The union, so nothing is cherry-picked: every committed slide that
    // carries a feature Gate 2 names is drawn and scored here.
    const chosen = new Set(coverage.flatMap((entry) => entry.slides));
    const byKey = new Map(facts.map((slide) => [slide.key, slide]));
    const featuresOf = (key: string): string[] =>
      coverage.filter((entry) => entry.slides.includes(key)).map((entry) => entry.key);

    const grids = new Map<string, ReturnType<typeof decodeGridSet>>();
    const oracleGridFor = (deck: string, key: string): Grid => {
      let set = grids.get(deck);
      if (set === undefined) {
        const file = `${deck}.ppt.grids`;
        set = decodeGridSet(new Uint8Array(readFileSync(join(ORACLE, 'grids', file))), file);
        grids.set(deck, set);
      }
      const grid = set.get(key);
      if (grid === undefined) {
        throw new FidelityError('FID_ORACLE_MISSING', `no oracle grid for ${key}`, key);
      }
      return grid;
    };

    const slides: SlideReport[] = [];
    const notDrawn: { key: string; reason: string }[] = [];

    for (const key of [...chosen].sort()) {
      const slide = byKey.get(key);
      if (slide === undefined) throw new FidelityError('FID_ORACLE_MISSING', `no facts for ${key}`);
      const probe = probes.find((entry) => entry.id === slide.deck);
      if (probe === undefined) {
        throw new FidelityError('FID_ORACLE_MISSING', `${slide.deck} is not a probe`);
      }

      let raster;
      let markup;
      try {
        raster = await renderSlide(harness.page, `/${probe.path}`, slide.slide - 1);
        markup = await slideMarkup(harness.page, `/${probe.path}`, slide.slide - 1);
      } catch (error) {
        const message = error instanceof Error ? error.message : String(error);
        notDrawn.push({ key, reason: (message.split('\n')[0] ?? message).trim() });
        continue;
      }

      const theirs = oracleGridFor(slide.deck, key);
      const difference = differenceOf(raster.grid, theirs);
      const score = scoreOf(difference);

      slides.push({
        key,
        deck: slide.deck,
        slide: slide.slide,
        features: featuresOf(key),
        meanBp: score.meanBp,
        maxD: score.maxD,
        hist: score.hist,
        regions: regionsOf(difference, theirs.width, theirs.height, REGION_FLOOR).slice(0, 5),
        oursSvg: markup.markup,
        theirsSvg: gridSvg(theirs),
        theirsPng: capturedPng(options.capture, key),
        heatSvg: heatmapSvg(difference, theirs.width, theirs.height, theirs.cell),
        cell: theirs.cell,
      });
    }

    const uncovered = coverage.filter((entry) => entry.slides.length === 0);
    return {
      slides,
      coverage,
      notDrawn,
      scanned: facts.length,
      ok: uncovered.length === 0 && notDrawn.length === 0,
    };
  } finally {
    await harness.close();
  }
}

/** The capture's own PNG for a slide, when the directory holds one. */
function capturedPng(capture: string | null, key: string): string | null {
  if (capture === null) return null;
  const file = join(capture, `${key}.png`);
  try {
    readFileSync(file);
    return file;
  } catch {
    return null;
  }
}
