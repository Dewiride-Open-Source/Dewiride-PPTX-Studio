/**
 * Gate 3's verification: the hundred-slide deck, through the page, at every zoom, offline.
 *
 * What is rasterised is the studio page's own DOM - `apps/studio/dist`, served over loopback
 * and driven through its automation hook - not the `<img>` path `tools/fidelity` scores by, so
 * what is measured is what a person would see. The deck is fetched once; then the context goes
 * offline and every render, zoom and screenshot has to come from what the page already holds.
 *
 * Gated, with no tolerance in it: a hundred slides, every one drawn at every zoom, the SVG at
 * any zoom identical to 100 % apart from what the stroke rule owns, no page error, no request
 * the gate does not recognise and none at all once offline, and every raster the one recorded.
 * The scores against PowerPoint's own export at each width are reported and gate nothing.
 */

import { createHash } from 'node:crypto';
import { existsSync, writeFileSync } from 'node:fs';
import { basename, dirname, join } from 'node:path';

import { chromium, type Page } from 'playwright';

import { startServer } from '../bench/serve.ts';
import { boxOf, regionShapes } from '../fidelity/blame.ts';
import { FidelityError } from '../fidelity/errors.ts';
import { claimFixtures } from '../fidelity/fixtures.ts';
import { gridSvg, heatmapSvg } from '../fidelity/heatmap.ts';
import type { Grid } from '../fidelity/metric/reduce.ts';
import { differenceOf, regionsOf, scoreOf, type Region } from '../fidelity/metric/score.ts';
import { openOracle, ORACLE_DIR, type OpenedOracle } from '../fidelity/oracle.ts';
import { PINNED_ARGS } from '../fidelity/raster/browser.ts';
import {
  assertSameEnvironment,
  currentEnvId,
  familiesOf,
  probeEnvironment,
  type Environment,
} from '../fidelity/raster/fonts.ts';
import { geometryOf, injectReduce, oracleGrid, type Geometry } from '../fidelity/raster/render.ts';
import {
  assertBaselineCovers,
  assertExpectedCount,
  assertMayRecord,
  baselineGaps,
  baselineJson,
  readBaseline,
  type RecordArgs,
} from '../fidelity/record.ts';
import { fidelityProbes, slideKey } from '../ground-truth/render/fidelity/probes.ts';
import { repoPath } from '../repo/root.ts';

const EMU_PER_POINT = 12700;

import {
  cellAt,
  GATE_DECK_ID,
  gateHolds,
  integerClip,
  invarianceOf,
  maskZoom,
  renameIdPrefix,
  requestVerdict,
  STAGE_ZOOMS,
  STRIP_ZOOM,
  widthAt,
  zoomKey,
  type Box,
  type LoggedRequest,
  type RequestVerdict,
  type ZoomEntry,
} from './checks.ts';

/** The floor a region has to reach to be worth ranking, in levels of 255. */
const REGION_FLOOR = 16;

const NAMES_PER_REGION = 2;

/** How many slides a zoom shows side by side with PowerPoint in the report. */
const WORST_PER_ZOOM = 3;

/** Wide enough for 400 % of a 960-point slide, and a strip of a hundred thumbnails in four rows. */
const VIEWPORT = { width: 4000, height: 2600 };

/** Where the zoomed baselines live, beside the zoomed oracle. */
export const ZOOM_DIR = join(ORACLE_DIR, 'zoom');

/**
 * The page's stage and strip pinned at whole page pixels, so a clip is exact and nothing is
 * resampled. Nothing inside an `<svg>` is touched; the strip's labels are hidden so a row is
 * exactly a thumbnail tall.
 */
const GATE_STYLE = `
[data-role=stage] { position: absolute !important; left: 0 !important; top: 0 !important;
  z-index: 10; max-width: none !important; max-height: none !important; overflow: visible !important;
  border: 0 !important; margin: 0 !important; padding: 0 !important; background: #fff; }
[data-role=strip] { position: absolute !important; left: 0 !important; top: 0 !important; z-index: 11;
  flex-wrap: wrap; gap: 8px; width: 3968px; overflow: visible !important; padding: 0 !important;
  margin: 0 !important; background: #fff; }
[data-role=strip] .thumb { padding: 0 !important; border: 0 !important; margin: 0 !important;
  border-radius: 0 !important; box-shadow: none !important; }
[data-role=strip] .thumb .n { display: none; }
html.gate3-strip [data-role=stage] { visibility: hidden; }
html:not(.gate3-strip) [data-role=strip] { visibility: hidden; }
`;

/** `@pptx-studio/text` on `globalThis.pptx`, which is what `probeEnvironment` measures through. */
const TEXT_PROBE = `import * as text from '/packages/text/dist/index.js'; globalThis.pptx = { text };`;

export interface Gate3Options {
  readonly out: string;
  readonly port: number;
  readonly headed: boolean;
  readonly record: RecordArgs;
  /** A deck outside the corpus, measured and never scored; `null` is the gate. */
  readonly deck: string | null;
}

/** What the page reported once the deck was open. */
interface DeckOpened {
  readonly slides: number;
  readonly size: { readonly cx: number; readonly cy: number };
  readonly timings: {
    readonly readMs: number;
    readonly parseMs: number;
    readonly firstStageMs: number;
    readonly stripMs: number;
    readonly stripCpuMs: number;
    readonly heapBytes: number | null;
  };
  readonly failed: readonly string[];
}

interface Shown {
  readonly mountMs: number;
  readonly width: number;
  readonly height: number;
}

interface ShapeFrame {
  readonly id: number;
  readonly name: string;
  readonly x: number;
  readonly y: number;
  readonly cx: number;
  readonly cy: number;
  readonly rot: number;
}

/** The page's automation hook, as narrowly as the gate uses it. */
interface StudioHook {
  openFromUrl(url: string): Promise<DeckOpened>;
  show(index: number, zoom: number): Promise<Shown>;
  stageSvg(): string;
  thumbSvg(index: number): string;
  stageShapes(): readonly ShapeFrame[];
  stageBox(): Box;
  thumbBox(index: number): Box;
}

/** A region, and the shapes its bounding box crosses. */
export interface NamedRegion extends Region {
  readonly shapes: readonly string[];
}

export interface ScoredSlide {
  readonly key: string;
  readonly slide: number;
  readonly meanBp: number;
  readonly maxD: number;
  readonly hist: readonly number[];
  readonly regions: readonly NamedRegion[];
  readonly rasterSha256: string;
  readonly svgSha256: string;
  readonly mountMs: number;
  readonly screenshotMs: number;
}

/** One of the worst slides at a zoom, with the pictures the report shows. */
export interface WorstSlide {
  readonly key: string;
  readonly meanBp: number;
  readonly oursPng: Uint8Array;
  readonly theirsSvg: string;
  readonly heatSvg: string;
}

export interface ZoomColumn {
  readonly zoom: number;
  readonly width: number;
  readonly cell: number;
  /** The stage, or the strip for the thumbnail zoom. */
  readonly surface: 'stage' | 'strip';
  readonly slides: readonly ScoredSlide[];
  readonly meanBp: number;
  readonly worst: readonly WorstSlide[];
}

export interface Gate3Run {
  /** A gate, or a measurement of a deck outside the corpus that scores nothing. */
  readonly mode: 'gate' | 'measurement';
  readonly deckPath: string;
  readonly slides: number;
  readonly size: { readonly cx: number; readonly cy: number };
  readonly zooms: readonly ZoomColumn[];
  readonly notDrawn: readonly { readonly key: string; readonly reason: string }[];
  /** `(slide key, zoom)` pairs whose masked SVG is not the one at 100 %. */
  readonly breaks: readonly { readonly key: string; readonly zoom: number }[];
  readonly pageErrors: readonly string[];
  readonly requests: RequestVerdict;
  readonly timings: DeckOpened['timings'];
  readonly environment: Environment | null;
  readonly baseline: {
    readonly changed: readonly string[];
    readonly unrecorded: readonly string[];
    readonly vanished: readonly string[];
    readonly recorded: string | null;
  } | null;
  readonly ok: boolean;
}

/* ----------------------------------------------------------------- the page */

/** The page's own scroll and class list, as narrowly as the gate touches them. */
interface PageGlobals {
  pptxStudio: StudioHook;
  scrollTo(x: number, y: number): void;
  document: { documentElement: { classList: { toggle(name: string, force: boolean): void } } };
}

/**
 * The automation hook, one `page.evaluate` per call. Every body is self-contained because
 * Playwright ships it by source; a box is read and the page scrolled home in the same call, so
 * the page coordinates it carries are the viewport's when the screenshot is clipped.
 */
function hook(page: Page): {
  open(url: string): Promise<DeckOpened>;
  show(index: number, zoom: number): Promise<Shown>;
  stageSvg(): Promise<string>;
  thumbSvg(index: number): Promise<string>;
  stageShapes(): Promise<readonly ShapeFrame[]>;
  stageBox(): Promise<Box>;
  thumbBox(index: number): Promise<Box>;
  stripPass(on: boolean): Promise<void>;
} {
  return {
    open: (url) =>
      page.evaluate((u) => (globalThis as unknown as PageGlobals).pptxStudio.openFromUrl(u), url),
    show: (index, zoom) =>
      page.evaluate(([i, z]) => (globalThis as unknown as PageGlobals).pptxStudio.show(i, z), [
        index,
        zoom,
      ] as const),
    stageSvg: () =>
      page.evaluate(() => (globalThis as unknown as PageGlobals).pptxStudio.stageSvg()),
    thumbSvg: (index) =>
      page.evaluate((i) => (globalThis as unknown as PageGlobals).pptxStudio.thumbSvg(i), index),
    stageShapes: () =>
      page.evaluate(() => (globalThis as unknown as PageGlobals).pptxStudio.stageShapes()),
    stageBox: () =>
      page.evaluate(() => {
        const g = globalThis as unknown as PageGlobals;
        const box = g.pptxStudio.stageBox();
        g.scrollTo(0, 0);
        return box;
      }),
    thumbBox: (index) =>
      page.evaluate((i) => {
        const g = globalThis as unknown as PageGlobals;
        const box = g.pptxStudio.thumbBox(i);
        g.scrollTo(0, 0);
        return box;
      }, index),
    stripPass: (on) =>
      page.evaluate((flag) => {
        (globalThis as unknown as PageGlobals).document.documentElement.classList.toggle(
          'gate3-strip',
          flag,
        );
      }, on),
  };
}

function sha256(text: string): string {
  return createHash('sha256').update(text).digest('hex');
}

function dataUrl(png: Uint8Array): string {
  return `data:image/png;base64,${Buffer.from(png).toString('base64')}`;
}

/** The three worst so far at a zoom, so a hundred screenshots are never held at once. */
function keepWorst(worst: WorstSlide[], candidate: WorstSlide): void {
  worst.push(candidate);
  worst.sort((a, b) => a.meanBp - b.meanBp);
  worst.splice(WORST_PER_ZOOM);
}

/* ------------------------------------------------------------------ the run */

export async function runGate3(options: Gate3Options): Promise<Gate3Run> {
  if (!existsSync(repoPath('apps/studio/dist/main.js'))) {
    throw new FidelityError(
      'FID_ENV_UNKNOWN',
      'apps/studio/dist/main.js is missing; run pnpm build',
    );
  }
  const measurement = options.deck !== null;
  const probe = fidelityProbes().find((entry) => entry.id === GATE_DECK_ID);
  if (probe === undefined) {
    throw new FidelityError('FID_ORACLE_MISSING', `${GATE_DECK_ID} is not a committed deck`);
  }
  const deckPath = options.deck === null ? `/${probe.path}` : `/decks/${basename(options.deck)}`;
  const deckId =
    options.deck === null ? GATE_DECK_ID : basename(options.deck).replace(/\.pptx$/i, '');
  const args = options.record;

  const envId = currentEnvId();
  const expectedPath = join(ZOOM_DIR, `expected.${envId}.json`);
  const expected = measurement ? null : readBaseline(expectedPath);
  if (!measurement) {
    if (expected === null && !args.record) {
      throw new FidelityError(
        'FID_ENV_UNKNOWN',
        `no zoom/expected.${envId}.json; this machine has never recorded a Gate 3 baseline`,
        envId,
      );
    }
    assertMayRecord(args, {
      inCi: process.env['CI'] !== undefined,
      hasBaseline: expected !== null,
    });
  }
  const oracle: OpenedOracle | null = measurement ? null : openOracle();

  const server = startServer({
    port: options.port,
    decks: options.deck === null ? null : dirname(options.deck),
  });
  const browser = await chromium.launch({ args: [...PINNED_ARGS], headless: !options.headed });
  const context = await browser.newContext({
    viewport: VIEWPORT,
    deviceScaleFactor: 1,
    colorScheme: 'light',
  });
  const page = await context.newPage();

  const pageErrors: string[] = [];
  page.on('pageerror', (error) => pageErrors.push(error.message));
  page.on('console', (message) => {
    if (message.type() === 'error') pageErrors.push(message.text());
  });
  let offline = false;
  const log: LoggedRequest[] = [];

  try {
    await page.goto(`${server.url}/`, { waitUntil: 'load' });
    await page.addScriptTag({ type: 'module', content: TEXT_PROBE });
    await page.waitForFunction(() => 'pptx' in globalThis);
    await injectReduce(page);
    await page.addStyleTag({ content: GATE_STYLE });
    // From here every request is the page's own: the harness has nothing left to fetch.
    page.on('request', (request) => log.push({ url: request.url(), afterOffline: offline }));

    const studio = hook(page);
    const opened = await studio.open(deckPath);
    await context.setOffline(true);
    offline = true;

    const notDrawn: { key: string; reason: string }[] = opened.failed.map((entry) => {
      const colon = entry.indexOf(':');
      return {
        key: `${slideKey(deckId, Number(entry.slice(0, colon)))} thumbnail`,
        reason: entry.slice(colon + 1).trim(),
      };
    });
    const breaks: { key: string; zoom: number }[] = [];
    const digests = new Map<string, { rasterSha256: string; svgSha256: string }>();
    const families = new Set<string>();
    const masked = new Map<number, ZoomEntry[]>();
    const zooms: ZoomColumn[] = [];

    /** Screenshot the box, reduce it as PowerPoint's PNGs are reduced, and check it is what was drawn. */
    const raster = async (
      box: Box,
      what: string,
      geometry: Geometry,
    ): Promise<{ grid: Grid; rasterSha256: string; png: Uint8Array; ms: number }> => {
      const clip = integerClip(box, what);
      if (clip.x + clip.width > VIEWPORT.width || clip.y + clip.height > VIEWPORT.height) {
        throw new FidelityError('FID_RASTER_GEOMETRY', `${what} does not fit the viewport`, what);
      }
      const started = performance.now();
      const png = new Uint8Array(
        await page.screenshot({ clip, animations: 'disabled', caret: 'hide' }),
      );
      const ms = performance.now() - started;
      const ours = await oracleGrid(page, dataUrl(png), geometry);
      if (
        ours.natural.width !== geometry.drawWidth ||
        ours.natural.height !== geometry.drawHeight
      ) {
        throw new FidelityError(
          'FID_SVG_INTRINSIC_SIZE',
          `${what} screenshot is ${String(ours.natural.width)}x${String(ours.natural.height)}, ` +
            `not the ${String(geometry.drawWidth)}x${String(geometry.drawHeight)} drawn`,
          what,
        );
      }
      if (args.record) {
        const again = await oracleGrid(
          page,
          dataUrl(
            new Uint8Array(await page.screenshot({ clip, animations: 'disabled', caret: 'hide' })),
          ),
          geometry,
        );
        if (again.rasterSha256 !== ours.rasterSha256) {
          throw new FidelityError(
            'FID_RASTER_NONDETERMINISTIC',
            `${what} rasterised twice in one run and did not agree with itself`,
            what,
          );
        }
      }
      return { grid: ours.grid, rasterSha256: ours.rasterSha256, png, ms };
    };

    /** Score one raster against PowerPoint's grid at that width; only the gate has one. */
    const score = (
      key: string,
      width: number,
      ours: Grid,
      geometry: Geometry,
      shapes: readonly ShapeFrame[],
    ): {
      meanBp: number;
      maxD: number;
      hist: readonly number[];
      regions: NamedRegion[];
      theirs: Grid;
      difference: readonly number[];
    } => {
      if (oracle === null)
        throw new FidelityError('FID_ORACLE_MISSING', 'a measurement scores nothing');
      const theirs =
        width === oracle.oracle.rasterWidth ? oracle.gridFor(key) : oracle.zoomGridFor(key, width);
      if (
        theirs.cell !== geometry.cell ||
        theirs.width * theirs.cell !== geometry.padWidth ||
        theirs.height * theirs.cell !== geometry.padHeight
      ) {
        throw new FidelityError(
          'FID_GRID_MISMATCH',
          `${key} at ${String(width)} px: the oracle grid is ${String(theirs.width)}x${String(theirs.height)} ` +
            `of ${String(theirs.cell)} px and ours is ${String(geometry.padWidth / geometry.cell)}x` +
            `${String(geometry.padHeight / geometry.cell)} of ${String(geometry.cell)} px`,
          key,
        );
      }
      const scale = width / opened.size.cx;
      const boxes = shapes.map((placed) =>
        boxOf({
          ...placed,
          x: placed.x * scale,
          y: placed.y * scale,
          cx: placed.cx * scale,
          cy: placed.cy * scale,
        }),
      );
      const difference = differenceOf(ours, theirs);
      const slideScore = scoreOf(difference);
      const regions = regionsOf(difference, theirs.width, theirs.height, REGION_FLOOR)
        .slice(0, 5)
        .map((region) => ({
          ...region,
          shapes: regionShapes(region, theirs.cell, boxes, NAMES_PER_REGION),
        }));
      return {
        meanBp: slideScore.meanBp,
        maxD: slideScore.maxD,
        hist: slideScore.hist,
        regions,
        theirs,
        difference,
      };
    };

    // The stage, zoom by zoom, so every (slide, zoom) is a fresh mount and never a resize.
    for (const zoom of STAGE_ZOOMS) {
      const slides: ScoredSlide[] = [];
      const worst: WorstSlide[] = [];
      let width = 0;
      let cell = 0;
      for (let index = 0; index < opened.slides; index++) {
        const key = slideKey(deckId, index + 1);
        let shown: Shown;
        try {
          shown = await studio.show(index, zoom);
        } catch (error) {
          const message = error instanceof Error ? error.message : String(error);
          notDrawn.push({
            key: `${key} at ${String(zoom * 100)} %`,
            reason: (message.split('\n')[0] ?? message).trim(),
          });
          continue;
        }
        width = shown.width;
        const svg = await studio.stageSvg();
        const entries = masked.get(index) ?? [];
        entries.push({ zoom, masked: maskZoom(svg) });
        masked.set(index, entries);
        if (zoom === 1) for (const family of familiesOf(svg)) families.add(family);
        if (measurement) {
          slides.push({
            key,
            slide: index + 1,
            meanBp: 0,
            maxD: 0,
            hist: [],
            regions: [],
            rasterSha256: '',
            svgSha256: sha256(svg),
            mountMs: shown.mountMs,
            screenshotMs: 0,
          });
          continue;
        }
        cell = cellAt(width);
        const geometry = geometryOf(opened.size.cx, opened.size.cy, { width, cell });
        const what = zoomKey(key, width);
        const drawn = await raster(await studio.stageBox(), what, geometry);
        const scored = score(key, width, drawn.grid, geometry, await studio.stageShapes());
        digests.set(what, { rasterSha256: drawn.rasterSha256, svgSha256: sha256(svg) });
        slides.push({
          key,
          slide: index + 1,
          meanBp: scored.meanBp,
          maxD: scored.maxD,
          hist: scored.hist,
          regions: scored.regions,
          rasterSha256: drawn.rasterSha256,
          svgSha256: sha256(svg),
          mountMs: shown.mountMs,
          screenshotMs: drawn.ms,
        });
        keepWorst(worst, {
          key,
          meanBp: scored.meanBp,
          oursPng: drawn.png,
          theirsSvg: gridSvg(scored.theirs),
          heatSvg: heatmapSvg(
            scored.difference,
            scored.theirs.width,
            scored.theirs.height,
            scored.theirs.cell,
          ),
        });
      }
      zooms.push({ zoom, width, cell, surface: 'stage', slides, meanBp: meanOf(slides), worst });
    }

    // The strip: the same slides at an eighth, drawn by the same page in another mount.
    await studio.stripPass(true);
    {
      const slides: ScoredSlide[] = [];
      const worst: WorstSlide[] = [];
      let width = widthAt(STRIP_ZOOM, opened.size.cx / EMU_PER_POINT);
      let cell = 0;
      for (let index = 0; index < opened.slides; index++) {
        const key = slideKey(deckId, index + 1);
        let svg: string;
        try {
          svg = await studio.thumbSvg(index);
        } catch {
          continue;
        }
        const entries = masked.get(index) ?? [];
        entries.push({
          zoom: STRIP_ZOOM,
          masked: maskZoom(renameIdPrefix(svg, `thumb${String(index)}`, `slide${String(index)}`)),
        });
        masked.set(index, entries);
        if (measurement) {
          slides.push({
            key,
            slide: index + 1,
            meanBp: 0,
            maxD: 0,
            hist: [],
            regions: [],
            rasterSha256: '',
            svgSha256: sha256(svg),
            mountMs: 0,
            screenshotMs: 0,
          });
          continue;
        }
        const box = await studio.thumbBox(index);
        width = box.width;
        cell = cellAt(width);
        const geometry = geometryOf(opened.size.cx, opened.size.cy, { width, cell });
        const what = zoomKey(key, width);
        const drawn = await raster(box, what, geometry);
        const scored = score(key, width, drawn.grid, geometry, []);
        digests.set(what, { rasterSha256: drawn.rasterSha256, svgSha256: sha256(svg) });
        slides.push({
          key,
          slide: index + 1,
          meanBp: scored.meanBp,
          maxD: scored.maxD,
          hist: scored.hist,
          regions: scored.regions,
          rasterSha256: drawn.rasterSha256,
          svgSha256: sha256(svg),
          mountMs: 0,
          screenshotMs: drawn.ms,
        });
        keepWorst(worst, {
          key,
          meanBp: scored.meanBp,
          oursPng: drawn.png,
          theirsSvg: gridSvg(scored.theirs),
          heatSvg: heatmapSvg(
            scored.difference,
            scored.theirs.width,
            scored.theirs.height,
            scored.theirs.cell,
          ),
        });
      }
      zooms.push({
        zoom: STRIP_ZOOM,
        width,
        cell,
        surface: 'strip',
        slides,
        meanBp: meanOf(slides),
        worst,
      });
    }
    await studio.stripPass(false);

    for (const [index, entries] of masked) {
      for (const zoom of invarianceOf(entries))
        breaks.push({ key: slideKey(deckId, index + 1), zoom });
    }

    const environment = await probeEnvironment(page, [...families].sort(), envId);
    const requests = requestVerdict(log, server.url, deckPath);

    let baseline: Gate3Run['baseline'] = null;
    if (!measurement) {
      const gaps = baselineGaps(Object.keys(expected?.slides ?? {}), digests.keys());
      if (expected !== null) {
        assertSameEnvironment(environment, expected.environment);
        if (!args.record) assertBaselineCovers(gaps);
      }
      const changed = [...digests]
        .filter(([key, digest]) => {
          const was = expected?.slides[key];
          return was !== undefined && was.rasterSha256 !== digest.rasterSha256;
        })
        .map(([key]) => key);
      let recorded: string | null = null;
      if (args.record) {
        assertExpectedCount(args, changed.length + gaps.unrecorded.length + gaps.vanished.length);
        writeFileSync(
          expectedPath,
          baselineJson(
            { environment, slides: Object.fromEntries(digests) },
            'tools/gate3/run-gate3.ts',
          ),
        );
        claimFixtures([
          {
            id: `gate3-expected-${envId}`,
            path: `render/fidelity/zoom/expected.${envId}.json`,
            tags: ['fidelity', 'baseline', 'render', 'zoom', 'gate3'],
            description:
              `The digest of the studio page's own raster of every ${GATE_DECK_ID} slide at every zoom ` +
              `on ${envId}, and the font environment it was recorded in. This is Gate 3's regression ` +
              'lock: a slide that rasterises differently here at any zoom has changed, with no tolerance.',
            recipe: {
              tool: 'tools/gate3/run-gate3.ts',
              args: args.bootstrap ? ['--record', '--bootstrap'] : ['--record'],
            },
            addedIn: '3.11',
          },
        ]);
        recorded = expectedPath;
      }
      baseline = { changed, unrecorded: gaps.unrecorded, vanished: gaps.vanished, recorded };
    }

    // A record run has just written what it compared with, so its differences are not failures.
    const gated = baseline !== null && baseline.recorded === null ? baseline : null;
    const ok =
      !measurement &&
      gateHolds({
        slides: opened.slides,
        notDrawn: notDrawn.length,
        pageErrors: pageErrors.length,
        breaks: breaks.length,
        changed: gated?.changed.length ?? 0,
        vanished: gated?.vanished.length ?? 0,
        offenders: requests.other.length,
        afterOffline: requests.afterOffline.length,
      });

    return {
      mode: measurement ? 'measurement' : 'gate',
      deckPath,
      slides: opened.slides,
      size: opened.size,
      zooms,
      notDrawn,
      breaks,
      pageErrors,
      requests,
      timings: opened.timings,
      environment,
      baseline,
      ok,
    };
  } finally {
    await browser.close();
    await server.close();
  }
}

function meanOf(slides: readonly ScoredSlide[]): number {
  if (slides.length === 0) return 0;
  return Math.round(slides.reduce((sum, slide) => sum + slide.meanBp, 0) / slides.length);
}
