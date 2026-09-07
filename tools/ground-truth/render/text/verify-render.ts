/**
 * Experiment T8, step 5 - lay the probe decks out ourselves and compare.
 *
 * ```
 * pnpm build
 * node tools/ground-truth/render/text/verify-render.ts <dir>
 * ```
 *
 * The other sub-phases end by exporting a file and opening it in the real
 * PowerPoint. A renderer cannot be checked that way round: the question is not
 * whether PowerPoint accepts what we wrote, it is whether what we draw is where
 * PowerPoint drew it. So this closes the loop the other way.
 *
 * Every T8 deck has already been opened by PowerPoint, which reported the box of
 * every line through its own object model. This loads the same packages through
 * `opc` and `model`, lays each shape's text out with `render-svg`, and compares
 * line for line: the top, the left and the height, in slide points.
 *
 * ## What a disagreement means
 *
 * A line's box is where the glyphs go, so a millimetre here is a millimetre on
 * the slide. What this cannot check is anything PowerPoint does not report - it
 * gives no baseline, which is why the baseline is measured from the EMF instead
 * and scored in `analyse.ts`.
 *
 * Reaches into `packages/*&#47;dist` for the same reason `check-roundtrip.ts` does:
 * nothing links the workspace packages into `tools/`.
 */

import { readFileSync } from 'node:fs';
import { join, resolve } from 'node:path';

import { chromium } from 'playwright';

import { readEmf } from '../../lib/emf.ts';
import { REPO_ROOT as ROOT } from '../../../repo/root.ts';

const dirArg = process.argv[2];
if (dirArg === undefined) throw new Error('usage: verify-render.ts <dir>');
const dir: string = resolve(dirArg);

interface LineReading {
  readonly index: number;
  readonly left: number | null;
  readonly top: number | null;
  readonly width: number | null;
  readonly height: number | null;
  readonly text: string | null;
}

interface SlideReading {
  readonly slide: number;
  readonly probe: string;
  readonly emf: string | null;
  readonly shapes: readonly {
    readonly name: string;
    readonly lines: readonly LineReading[];
  }[];
}

interface DeckReading {
  readonly deck: string;
  readonly file: string;
  readonly slides: readonly SlideReading[];
}

const readings = JSON.parse(readFileSync(join(dir, 'text-render-readings.json'), 'utf8')) as {
  decks: readonly DeckReading[];
};

/** Every deck, every slide, every line PowerPoint reported. */
const wanted = readings.decks.flatMap((deck) =>
  deck.slides.map((slide) => ({
    deck: deck.deck,
    file: deck.file,
    slide: slide.slide,
    probe: slide.probe,
    lines: slide.shapes.find((shape) => shape.name === slide.probe)?.lines ?? [],
    // `EMR_SETTEXTCOLOR` is `0x00bbggrr`; every text record carries the colour
    // in force when PowerPoint drew it.
    colors:
      slide.emf === null
        ? []
        : readEmf(new Uint8Array(readFileSync(join(dir, slide.emf)))).texts.map((text) =>
            text.colorBgr === null
              ? null
              : ((text.colorBgr & 0xff) << 16) |
                (text.colorBgr & 0xff00) |
                ((text.colorBgr >> 16) & 0xff),
          ),
  })),
);

const files = [...new Set(wanted.map((row) => row.file))];
const decks = new Map(
  files.map((file) => [file, [...new Uint8Array(readFileSync(join(dir, file)))]] as const),
);

/**
 * The packages, served to the page from a fake origin.
 *
 * A `file://` page cannot import another `file://` module, so the built
 * packages are routed instead: `/pkg/<name>/index.js` is that package's
 * `dist/index.js` and `/pkg/<name>/<rest>` is the chunk beside it, which is
 * where a relative import inside `dist` lands on its own.
 */
const PACKAGES = ['opc', 'xml', 'paint', 'geometry', 'text', 'model', 'render-svg'] as const;

const ORIGIN = 'https://pptx-studio.invalid';

const IMPORT_MAP = {
  imports: {
    ...Object.fromEntries(
      PACKAGES.map((name) => [`@pptx-studio/${name}`, `/pkg/${name}/index.js`]),
    ),
    '@pptx-studio/geometry/presets/': '/pkg/geometry/presets/',
    fflate: '/pkg/fflate/index.js',
  },
};

const PAGE =
  '<!doctype html><meta charset="utf-8">' +
  `<script type="importmap">${JSON.stringify(IMPORT_MAP)}</script>`;

function served(pathname: string): { body: Buffer; type: string } | null {
  const match = /^\/pkg\/([^/]+)\/(.+)$/.exec(pathname);
  if (match === null) return null;
  const [, name, rest] = match;
  if (name === undefined || rest === undefined) return null;
  const file =
    name === 'fflate'
      ? join(ROOT, 'packages', 'opc', 'node_modules', 'fflate', 'esm', 'browser.js')
      : join(ROOT, 'packages', name, 'dist', rest);
  try {
    return { body: readFileSync(file), type: 'text/javascript' };
  } catch {
    return null;
  }
}

const browser = await chromium.launch();
try {
  const page = await browser.newPage();
  page.on('console', (message) => {
    if (message.type() === 'error') console.log('  page:', message.text());
  });
  await page.route(
    (url) => url.origin === ORIGIN,
    async (route) => {
      const url = new URL(route.request().url());
      if (url.pathname === '/') {
        await route.fulfill({ body: PAGE, contentType: 'text/html' });
        return;
      }
      const file = served(url.pathname);
      if (file === null) {
        console.log('  404', url.pathname);
        await route.fulfill({ status: 404, body: url.pathname });
        return;
      }
      await route.fulfill({ body: file.body, contentType: file.type });
    },
  );
  await page.goto(`${ORIGIN}/`);

  const measured = (await page.evaluate(
    async ({
      urls,
      probes,
    }: {
      urls: Record<string, readonly number[]>;
      probes: readonly { file: string; slide: number; probe: string }[];
    }) => {
      const load = (name: string): Promise<unknown> =>
        import(/* @vite-ignore */ `/pkg/${name}/index.js`);
      const opc = (await load('opc')) as {
        PartStore: { open: (bytes: Uint8Array) => unknown };
      };
      const model = (await load('model')) as {
        loadDocument: (store: unknown) => { slides: readonly unknown[] };
      };
      const render = (await load('render-svg')) as {
        layoutSlide: (
          sheet: unknown,
        ) => readonly { shape: { name: string }; frame: { x: number; y: number } }[];
        createTextEngine: (options?: unknown) => unknown;
        textBlockOf: (
          placed: unknown,
          engine: unknown,
        ) => {
          turnDeg: number;
          lines: readonly {
            topPt: number;
            leftPt: number;
            heightPt: number;
            pieces: readonly { color: { r: number; g: number; b: number } | null }[];
          }[];
        } | null;
      };

      const EMU_PER_POINT = 12700;
      const loaded = new Map<string, { slides: readonly unknown[] }>();
      const out: Record<
        string,
        {
          turnDeg: number;
          lines: { top: number; left: number; height: number }[];
          colors: (number | null)[];
        }
      > = {};

      for (const entry of probes) {
        let document = loaded.get(entry.file);
        if (document === undefined) {
          document = model.loadDocument(
            opc.PartStore.open(Uint8Array.from(urls[entry.file] ?? [])),
          );
          loaded.set(entry.file, document);
        }
        const sheet = document.slides[entry.slide - 1];
        if (sheet === undefined) continue;
        const engine = render.createTextEngine();
        for (const placed of render.layoutSlide(sheet)) {
          if (placed.shape.name !== entry.probe) continue;
          const block = render.textBlockOf(placed, engine);
          if (block === null) continue;
          // Slide points, which is what PowerPoint's own line boxes are in.
          out[entry.probe] = {
            turnDeg: block.turnDeg,
            lines: block.lines.map((line) => ({
              top: placed.frame.y / EMU_PER_POINT + line.topPt,
              left: placed.frame.x / EMU_PER_POINT + line.leftPt,
              height: line.heightPt,
            })),
            colors: block.lines.flatMap((line) =>
              line.pieces.map((piece) =>
                piece.color === null
                  ? null
                  : (Math.round(piece.color.r * 255) << 16) |
                    (Math.round(piece.color.g * 255) << 8) |
                    Math.round(piece.color.b * 255),
              ),
            ),
          };
        }
      }
      return out;
    },
    {
      urls: Object.fromEntries(decks),
      probes: wanted.map((row) => ({ file: row.file, slide: row.slide, probe: row.probe })),
    },
  )) as Record<
    string,
    {
      turnDeg: number;
      lines: { top: number; left: number; height: number }[];
      colors: (number | null)[];
    }
  >;

  /** A line's box is where the glyphs go, so this is a tenth of a point. */
  const TOLERANCE = 0.1;

  /**
   * How far a centred or right-aligned line may sit from PowerPoint's.
   *
   * Its left edge is the column less the line's own advance, so it carries the
   * measurer's disagreement with PowerPoint whole - 3.2 measured that at a
   * median of 0.15% and a p95 of 0.70% of the string, which on a 128pt line is
   * a fifth of a point. This is not a tolerance on the layout; it is the
   * browser-agreement number arriving where it always was going to.
   */
  const ADVANCE_TOLERANCE = 1;

  let compared = 0;
  let agreed = 0;
  let turned = 0;
  let anchored = 0;
  let drift = 0;
  let colours = 0;
  let colourAgreed = 0;
  const worst: { probe: string; field: string; ours: number; theirs: number }[] = [];

  for (const row of wanted) {
    const ours = measured[row.probe];
    if (ours === undefined) continue;
    // A turned block's `BoundTop` is the axis-aligned box of the turned line,
    // which is a different quantity; the turn itself is measured exactly from
    // the EMF and scored in `analyse.ts`.
    if (ours.turnDeg !== 0) {
      turned += 1;
      continue;
    }
    // A run PowerPoint drew in a colour it stated is a run whose `a:solidFill`
    // reached the piece; without this the whole colour path is untested, and it
    // was wrong.
    row.colors.forEach((theirColour, index) => {
      const mine = ours.colors[index];
      if (theirColour === null || mine === undefined || mine === null) return;
      colours += 1;
      if (mine === theirColour) colourAgreed += 1;
      else {
        worst.push({ probe: `${row.probe} colour`, field: 'rgb', ours: mine, theirs: theirColour });
      }
    });

    const theirs = row.lines;
    if (theirs.length !== ours.lines.length) {
      worst.push({
        probe: row.probe,
        field: 'lines',
        ours: ours.lines.length,
        theirs: theirs.length,
      });
      continue;
    }
    theirs.forEach((line, index) => {
      const mine = ours.lines[index];
      if (mine === undefined || line.top === null || line.height === null || line.left === null) {
        return;
      }
      for (const [field, a, b] of [
        ['top', mine.top, line.top],
        ['left', mine.left, line.left],
        ['height', mine.height, line.height],
      ] as const) {
        compared += 1;
        const off = Math.abs(a - b);
        if (off <= TOLERANCE) {
          agreed += 1;
          continue;
        }
        if (field === 'left') {
          drift = Math.max(drift, off);
          if (off <= ADVANCE_TOLERANCE) {
            anchored += 1;
            continue;
          }
        }
        worst.push({ probe: `${row.probe}#${String(index)}`, field, ours: a, theirs: b });
      }
    });
  }

  for (const row of worst.slice(0, 25)) {
    console.log(
      `  ${row.probe.padEnd(28)} ${row.field.padEnd(7)} ours ${row.ours.toFixed(3).padStart(10)}  PowerPoint ${row.theirs.toFixed(3).padStart(10)}`,
    );
  }
  if (worst.length > 25) console.log(`  ... and ${String(worst.length - 25)} more`);
  console.log(
    `\ntext render: ${String(agreed)}/${String(compared)} line measurement(s) within ${String(TOLERANCE)}pt across ${String(Object.keys(measured).length - turned)} unturned probe(s), ` +
      `${String(turned)} turned one(s) left to the EMF`,
  );
  console.log(
    `             ${String(colourAgreed)}/${String(colours)} run colour(s) match the colour PowerPoint drew`,
  );
  if (anchored > 0) {
    console.log(
      `             ${String(anchored)} anchored line(s) inside the measurer's own disagreement, worst ${drift.toFixed(3)}pt`,
    );
  }
  if (agreed + anchored !== compared || colourAgreed !== colours) process.exitCode = 1;
} finally {
  await browser.close();
}
