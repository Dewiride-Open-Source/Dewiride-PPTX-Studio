/**
 * Experiment C6, step 4 - render the probe decks ourselves and compare.
 *
 * ```
 * pnpm build
 * node tools/ground-truth/render/verify-render.ts <dir>
 * ```
 *
 * The other sub-phases end by exporting a file and opening it in the real
 * PowerPoint. A renderer cannot be checked that way round: the question is not
 * whether PowerPoint accepts what we wrote, it is whether what we draw is what
 * PowerPoint draws. So this closes the loop the other way.
 *
 * Every C6 deck has already been opened by PowerPoint and exported as a BMP at
 * 1920 x 1080, and every finding has a sample point on one of those bitmaps.
 * This loads the same packages through `opc` and `model`, renders each slide
 * with `render-svg`, rasterises the SVG in Chromium at the identical size, and
 * reads the identical pixels. Two renderers, one deck, one list of coordinates.
 *
 * ## What a disagreement means, and what it does not
 *
 * A sample is a single pixel of a shape's interior chosen to be far from any
 * edge, so antialiasing is not in play and a mismatch is a real difference in
 * what was drawn. What this cannot check is anything it has no sample for -
 * there is no sample on a shadow's penumbra or a dash's gap - and it is a
 * comparison at 54 points, not a fidelity score. The scored, whole-image
 * version is sub-phase 3.9.
 *
 * Reaches into `packages/*&#47;dist` for the same reason `check-roundtrip.ts` does:
 * nothing links the workspace packages into `tools/`.
 */

import { readFileSync, writeFileSync } from 'node:fs';
import { join, resolve } from 'node:path';
import { pathToFileURL } from 'node:url';

import { REPO_ROOT as ROOT } from '../../repo/root.ts';

import { chromium } from 'playwright';

/*
 * The body of `page.evaluate` runs in Chromium, not in Node, so the browser
 * globals it uses are not in scope for the type checker - `tools/` is compiled
 * with `types: ["node"]` and no DOM library, deliberately. As in
 * `tools/ground-truth/fonts/embedding/verify-in-browser.ts`, declare exactly the surface this one file uses
 * rather than widening the project's lib: narrow enough that a body which
 * starts using something else fails to compile instead of becoming `any`.
 */
interface Ctx2D {
  fillStyle: string;
  fillRect(x: number, y: number, w: number, h: number): void;
  drawImage(image: BrowserImage, x: number, y: number, w: number, h: number): void;
  getImageData(x: number, y: number, w: number, h: number): { data: Uint8ClampedArray };
}
interface Canvas {
  width: number;
  height: number;
  getContext(id: '2d'): Ctx2D | null;
}
interface BrowserImage {
  src: string;
  decode(): Promise<void>;
}
declare const document: {
  querySelector(selector: string): unknown;
  createElement(tag: 'canvas'): Canvas;
};
declare const Image: new () => BrowserImage;
declare const XMLSerializer: new () => { serializeToString(node: unknown): string };

const dirArg = process.argv[2];
if (dirArg === undefined) throw new Error('usage: verify-render.ts <dir>');
const dir: string = dirArg;

interface Sample {
  readonly id: string;
  readonly probe: string;
  readonly x: number;
  readonly y: number;
  readonly asks: string;
  readonly rgb: string;
  readonly block: string;
  /** The mean of a 21x21 block, which is what the comparison runs on. */
  readonly mean: string;
}

interface Fixture {
  readonly exportPixels: { readonly w: number; readonly h: number };
  readonly slidePoints: { readonly w: number; readonly h: number };
  readonly probes: readonly { id: string; deck: string; slide: number }[];
  readonly samples: readonly Sample[];
  readonly decks: readonly { deck: string; repaired: boolean | null }[];
}

const fixture = JSON.parse(
  readFileSync(join(ROOT, 'corpus/ground-truth/transforms.json'), 'utf8'),
) as Fixture;
const inputs = JSON.parse(readFileSync(join(dir, 'transform-inputs.json'), 'utf8')) as {
  decks: readonly { deck: string; file: string }[];
};

async function load<T>(pkg: string): Promise<T> {
  return (await import(pathToFileURL(resolve(ROOT, `packages/${pkg}/dist/index.js`)).href)) as T;
}

const opc = await load<{
  PartStore: { open: (bytes: Uint8Array) => unknown };
}>('opc');
const model = await load<{
  loadDocument: (store: unknown) => {
    slides: readonly unknown[];
    slideSize: { cx: number; cy: number };
    problems: readonly { message: string }[];
  };
}>('model');
const renderSvg = await load<{
  renderSlide: (sheet: unknown, size: unknown, options?: unknown) => string;
}>('render-svg');

const PX_PER_PT = fixture.exportPixels.w / fixture.slidePoints.w;

const browser = await chromium.launch();
const page = await browser.newPage({
  viewport: { width: fixture.exportPixels.w, height: fixture.exportPixels.h },
  deviceScaleFactor: 1,
});

interface Row {
  readonly sample: string;
  readonly probe: string;
  readonly powerpoint: string;
  readonly ours: string;
  readonly agree: boolean;
  readonly asks: string;
}

const rows: Row[] = [];
const errors: string[] = [];

for (const deck of inputs.decks) {
  const repaired = fixture.decks.find((d) => d.deck === deck.deck)?.repaired;
  // A repaired deck is not a measurement, and it is not a comparison either:
  // PowerPoint drew a file it rewrote and we would be drawing the original.
  if (repaired === true) continue;

  const bytes = new Uint8Array(readFileSync(join(dir, deck.file)));
  let document;
  try {
    const store = opc.PartStore.open(bytes);
    document = model.loadDocument(store);
  } catch (error) {
    errors.push(`${deck.deck}: ${String(error)}`);
    continue;
  }

  for (const [index, slide] of document.slides.entries()) {
    const wanted = fixture.samples.filter((sample) => {
      const probe = fixture.probes.find((p) => p.id === sample.probe);
      return probe !== undefined && probe.deck === deck.deck && probe.slide === index + 1;
    });
    if (wanted.length === 0) continue;

    let svg: string;
    try {
      svg = renderSvg.renderSlide(slide, document.slideSize, {
        width: fixture.exportPixels.w,
        height: fixture.exportPixels.h,
        idPrefix: `${deck.deck}-${String(index)}`,
      });
    } catch (error) {
      errors.push(`${deck.deck} slide ${String(index + 1)}: ${String(error)}`);
      continue;
    }

    // A white ground, because PowerPoint's export has one and these decks
    // declare a white background rather than none.
    await page.setContent(
      `<!doctype html><html><body style="margin:0;background:#fff">${svg}</body></html>`,
    );

    const points = wanted.map((sample) => ({
      x: Math.round(sample.x * PX_PER_PT),
      y: Math.round(sample.y * PX_PER_PT),
    }));
    const got = await sampleInPage(page, points);

    wanted.forEach((sample, at) => {
      const ours = got[at] ?? 'FFFFFF';
      rows.push({
        sample: sample.id,
        probe: sample.probe,
        powerpoint: sample.mean,
        ours,
        agree: near(sample.mean, ours),
        asks: sample.asks,
      });
    });
  }
}

await browser.close();

/**
 * Two colours close enough to be the same paint.
 *
 * Eight levels out of 255. Loose enough for the two rasterisers to disagree
 * about a half-tone's antialiasing and tight enough that a wrong colour, a
 * wrong gradient position or a missing shape is a failure.
 */
function near(a: string, b: string, tol = 8): boolean {
  for (let i = 0; i < 3; i++) {
    const av = parseInt(a.slice(i * 2, i * 2 + 2), 16);
    const bv = parseInt(b.slice(i * 2, i * 2 + 2), 16);
    if (Math.abs(av - bv) > tol) return false;
  }
  return true;
}

/**
 * The colours at a handful of points, read inside the page.
 *
 * Only the sampled pixels cross the bridge. The first version of this shipped
 * the whole `ImageData` back - eight million numbers per slide, over CDP, fifty
 * times - and did not finish in ten minutes. The comparison needs 54 pixels.
 */
async function sampleInPage(
  target: typeof page,
  points: readonly { x: number; y: number }[],
): Promise<readonly string[]> {
  return target.evaluate(
    async ({ at, width, height }) => {
      const svg = document.querySelector('svg');
      if (svg === null) throw new Error('nothing rendered');
      const source = new XMLSerializer().serializeToString(svg);
      const image = new Image();
      image.src = `data:image/svg+xml;charset=utf-8,${encodeURIComponent(source)}`;
      await image.decode();

      const canvas = document.createElement('canvas');
      canvas.width = width;
      canvas.height = height;
      const context = canvas.getContext('2d');
      if (context === null) throw new Error('no 2d context');
      // The slide's own background may be transparent where nothing is drawn,
      // and PowerPoint's export has a white page behind it.
      context.fillStyle = '#FFFFFF';
      context.fillRect(0, 0, width, height);
      context.drawImage(image, 0, 0, width, height);

      const hexOf = (n: number): string => n.toString(16).padStart(2, '0').toUpperCase();
      // The same 21 x 21 block the analysis averaged on PowerPoint's bitmap.
      return at.map((point) => {
        const block = context.getImageData(point.x - 10, point.y - 10, 21, 21).data;
        let red = 0;
        let green = 0;
        let blue = 0;
        for (let i = 0; i < block.length; i += 4) {
          red += block[i] ?? 0;
          green += block[i + 1] ?? 0;
          blue += block[i + 2] ?? 0;
        }
        const n = block.length / 4;
        return `${hexOf(Math.round(red / n))}${hexOf(Math.round(green / n))}${hexOf(Math.round(blue / n))}`;
      });
    },
    { at: [...points], width: fixture.exportPixels.w, height: fixture.exportPixels.h },
  );
}

/* -------------------------------------------------------------------------- */

const agreed = rows.filter((r) => r.agree).length;
console.log('');
console.log(`${String(agreed)}/${String(rows.length)} samples agree with PowerPoint`);
console.log('');
for (const row of rows) {
  if (row.agree) continue;
  console.log(
    `  ${row.sample.padEnd(20)} PowerPoint ${row.powerpoint}  ours ${row.ours}   (${row.asks})`,
  );
}
for (const error of errors) console.log(`  ! ${error}`);

writeFileSync(
  join(dir, 'verify-render.json'),
  `${JSON.stringify({ agreed, total: rows.length, rows, errors }, null, 2)}\n`,
);

if (errors.length > 0 || agreed < rows.length) process.exitCode = 1;
