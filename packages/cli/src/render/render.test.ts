import { mkdtempSync, readFileSync, readdirSync, writeFileSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { beforeAll, describe, expect, it } from 'vitest';

import { repoPath } from '../../../../tools/repo/root.ts';
import { buildFont } from '../../../../tools/ground-truth/lib/truetype.ts';
import { main, type Streams } from '../main.js';
import { isRenderError } from './errors.js';
import {
  DEFAULT_WIDTH,
  renderDeck,
  runRender,
  type RenderDeckOptions,
  type RenderOptions,
} from './render.js';

/**
 * `pptx-studio render`, over the corpus.
 *
 * Text is drawn in fonts built for the test and written to a temporary
 * directory, with the platform's own font directories switched off. That is not
 * only hygiene: a suite that measured whatever this machine happens to have
 * installed would report a different width on the next machine, and the widths
 * are the thing under test.
 */

const FAMILIES = ['Arial', 'Calibri', 'Times New Roman', 'Courier New'] as const;

let fontDir: string;

beforeAll(() => {
  fontDir = mkdtempSync(join(tmpdir(), 'pptx-studio-render-'));
  for (const family of FAMILIES) {
    writeFileSync(join(fontDir, `${family}.ttf`), buildFont({ familyName: family }).bytes);
  }
});

/** The verb's shape: what `main` builds and `runRender` takes. */
function options(over: Partial<RenderOptions> = {}): RenderOptions {
  return {
    slide: null,
    width: DEFAULT_WIDTH,
    out: null,
    fontDirs: [fontDir],
    systemFonts: false,
    text: true,
    json: false,
    quiet: false,
    ...over,
  };
}

/** The library's shape: only what a drawing has an opinion about. */
function drawing(over: RenderDeckOptions = {}): RenderDeckOptions {
  return { fontDirs: [fontDir], systemFonts: false, ...over };
}

/** The code of the `RenderError` a call throws, as a string to assert on. */
function codeOf(run: () => unknown): string {
  try {
    run();
  } catch (error) {
    return isRenderError(error) ? error.code : `not a RenderError: ${String(error)}`;
  }
  return 'nothing was thrown';
}

function deck(name: string): Uint8Array {
  return new Uint8Array(readFileSync(repoPath('corpus/decks', name)));
}

function streams(): { streams: Streams; out: () => string; err: () => string } {
  const outParts: string[] = [];
  const errParts: string[] = [];
  return {
    streams: { out: (t) => outParts.push(t), err: (t) => errParts.push(t) },
    out: () => outParts.join(''),
    err: () => errParts.join(''),
  };
}

describe('drawing a deck', () => {
  it('renders every slide as its own SVG document', () => {
    const result = renderDeck(deck('a01-minimal.pptx'), options());
    expect(result.slides.length).toBeGreaterThan(0);
    for (const slide of result.slides) {
      expect(slide.svg.startsWith('<svg')).toBe(true);
      expect(slide.svg).toContain('xmlns="http://www.w3.org/2000/svg"');
      expect(slide.svg).toContain('viewBox="0 0 ');
      expect(slide.svg.endsWith('</svg>')).toBe(true);
    }
  });

  it('defaults to every slide, 1920 wide, with text drawn', () => {
    // One call pins three defaults, and the font directory is the temporary one
    // rather than this machine's.
    const result = renderDeck(deck('a07-text-cascade.pptx'), drawing());
    expect(result.slides.length).toBe(3);
    expect([result.width, result.height]).toEqual([DEFAULT_WIDTH, 1080]);
    expect(result.facesIndexed).toBe(FAMILIES.length);
    expect(result.slides[0]?.svg).toContain('<text');
  });

  it('takes the shape a web route hands it: a width, and nothing about files', () => {
    const result = renderDeck(deck('a01-minimal.pptx'), { width: 640, text: false });
    expect([result.width, result.height]).toEqual([640, 360]);
    expect(result.slides.length).toBe(1);
  });

  it("takes none of the command's own fields", () => {
    // An accepted `out` would be a field the function ignores, so the unused
    // directive is what fails if the two shapes merge again.
    // @ts-expect-error `out` is the command's, not the library's
    renderDeck(deck('a01-minimal.pptx'), { text: false, out: null });
  });

  it('refuses a width that is not a positive whole number of pixels', () => {
    const bad = [0, -1, 1.5, Number.NaN];
    for (const width of bad) {
      expect(codeOf(() => renderDeck(deck('a01-minimal.pptx'), drawing({ width })))).toBe(
        'CLI_WIDTH',
      );
    }
  });

  it('takes its height from the deck aspect, not from the width it was asked for', () => {
    // a01 is 16:9, so the numbers are known rather than merely self-consistent -
    // comparing two renders to each other passes for `height = width` too.
    const wide = renderDeck(deck('a01-minimal.pptx'), options({ width: 1920 }));
    const small = renderDeck(deck('a01-minimal.pptx'), options({ width: 480 }));
    expect([wide.width, wide.height]).toEqual([1920, 1080]);
    expect([small.width, small.height]).toEqual([480, 270]);
    expect(wide.height).not.toBe(wide.width);
  });

  it('draws text, which is the whole reason this needs fonts', () => {
    const withText = renderDeck(deck('a07-text-cascade.pptx'), options());
    const without = renderDeck(deck('a07-text-cascade.pptx'), options({ text: false }));
    expect(withText.slides[0]?.svg).toContain('<text');
    expect(without.slides[0]?.svg).not.toContain('<text');
    expect(withText.slides[0]!.svg.length).toBeGreaterThan(without.slides[0]!.svg.length * 4);
  });

  it('asks no font questions at all with text off', () => {
    // No font option of any kind: `indexFonts` is not reached, so the default
    // `systemFonts` never gets to scan anything.
    const result = renderDeck(deck('a07-text-cascade.pptx'), { text: false });
    expect(result.fonts).toEqual([]);
    expect(result.facesIndexed).toBe(0);
    expect(result.fontDirectories).toEqual([]);
    expect(result.slides[0]?.svg).not.toContain('<text');
  });

  it('says which face drew each typeface the deck named', () => {
    const result = renderDeck(deck('a07-text-cascade.pptx'), options());
    expect(result.fonts.length).toBeGreaterThan(0);
    for (const font of result.fonts) expect(FAMILIES).toContain(font.drawn);
  });

  it('embeds a picture rather than linking one, so the file stands alone', () => {
    const result = renderDeck(deck('a03-fills.pptx'), options());
    const svg = result.slides.map((slide) => slide.svg).join('');
    expect(svg).toContain('data:image/');
    expect(svg).not.toMatch(/href="(?!data:)[^"]*\.(png|jpe?g)"/);
  });

  it('refuses a slide the deck does not have', () => {
    expect(() => renderDeck(deck('a01-minimal.pptx'), options({ slide: 99 }))).toThrowError(
      /the deck has \d+ slide/,
    );
  });

  it('shares no id between two slides, so both can go in one page', () => {
    const result = renderDeck(deck('a03-fills.pptx'), options());
    const ids = result.slides.map(
      (slide) => new Set([...slide.svg.matchAll(/\sid="([^"]+)"/g)].map((m) => m[1])),
    );
    const total = ids.reduce((sum, set) => sum + set.size, 0);
    expect(total, 'no slide in a03-fills declares an id, so this proves nothing').toBeGreaterThan(
      0,
    );
    expect(new Set(ids.flatMap((set) => [...set])).size).toBe(total);
  });
});

describe('the command line', () => {
  it('writes one file per slide into a directory', () => {
    const out = mkdtempSync(join(tmpdir(), 'pptx-studio-out-'));
    const s = streams();
    const code = main(
      [
        'render',
        repoPath('corpus/decks/a03-fills.pptx'),
        '--out',
        out,
        '--font-dir',
        fontDir,
        '--no-system-fonts',
      ],
      s.streams,
    );
    expect(s.err()).toBe('');
    expect(code).toBe(0);
    const written = readdirSync(out).sort();
    expect(written.length).toBeGreaterThan(1);
    expect(written[0]).toMatch(/^slide-0*1\.svg$/);
    expect(readFileSync(join(out, written[0]!), 'utf8').startsWith('<svg')).toBe(true);
  });

  it('writes one named file for one slide', () => {
    const out = join(mkdtempSync(join(tmpdir(), 'pptx-studio-out-')), 'first.svg');
    const s = streams();
    const code = main(
      [
        'render',
        repoPath('corpus/decks/a01-minimal.pptx'),
        '--slide',
        '1',
        '--out',
        out,
        '--font-dir',
        fontDir,
        '--no-system-fonts',
      ],
      s.streams,
    );
    expect(code).toBe(0);
    expect(readFileSync(out, 'utf8').startsWith('<svg')).toBe(true);
  });

  it('refuses to write many slides to one .svg path', () => {
    const out = join(mkdtempSync(join(tmpdir(), 'pptx-studio-out-')), 'all.svg');
    const s = streams();
    const code = main(
      [
        'render',
        repoPath('corpus/decks/a03-fills.pptx'),
        '--out',
        out,
        '--font-dir',
        fontDir,
        '--no-system-fonts',
      ],
      s.streams,
    );
    expect(code).toBe(1);
    expect(s.err()).toContain('CLI_OUTPUT_PATH');
  });

  it('writes the markup to stdout when no --out is given', () => {
    const s = streams();
    const code = main(
      [
        'render',
        repoPath('corpus/decks/a01-minimal.pptx'),
        '--slide',
        '1',
        '--font-dir',
        fontDir,
        '--no-system-fonts',
      ],
      s.streams,
    );
    expect(code).toBe(0);
    expect(s.out().startsWith('<svg')).toBe(true);
  });

  it('reports the fonts it used as JSON', () => {
    const s = streams();
    const code = main(
      [
        'render',
        repoPath('corpus/decks/a07-text-cascade.pptx'),
        '--json',
        '--font-dir',
        fontDir,
        '--no-system-fonts',
      ],
      s.streams,
    );
    expect(code).toBe(0);
    const report = JSON.parse(s.out()) as {
      fonts: { asked: string; drawn: string }[];
      facesIndexed: number;
      slides: { number: number; bytes: number }[];
    };
    expect(report.facesIndexed).toBe(FAMILIES.length);
    expect(report.fonts.length).toBeGreaterThan(0);
    expect(report.slides.every((slide) => slide.bytes > 0)).toBe(true);
  });

  it('names the code and the typeface when no face can stand in', () => {
    const empty = mkdtempSync(join(tmpdir(), 'pptx-studio-nofonts-'));
    const s = streams();
    const code = main(
      [
        'render',
        repoPath('corpus/decks/a07-text-cascade.pptx'),
        '--font-dir',
        empty,
        '--no-system-fonts',
      ],
      s.streams,
    );
    expect(code).toBe(1);
    expect(s.err()).toContain('CLI_NO_FACE');
  });

  it('names the code when a --font-dir is not there', () => {
    const s = streams();
    const code = main(
      [
        'render',
        repoPath('corpus/decks/a01-minimal.pptx'),
        '--font-dir',
        join(tmpdir(), 'nope-3f8a'),
      ],
      s.streams,
    );
    expect(code).toBe(1);
    expect(s.err()).toContain('CLI_FONT_DIR');
  });

  it('refuses a --slide that is not a positive integer', () => {
    const s = streams();
    expect(
      main(['render', repoPath('corpus/decks/a01-minimal.pptx'), '--slide', '0'], s.streams),
    ).toBe(2);
    expect(s.err()).toContain('--slide');
  });

  it('names the typefaces in its summary when the caller left text alone', () => {
    // `main` always spells `text` out, so only a caller that omits it can catch
    // a summary that reads the field instead of its default.
    const out = mkdtempSync(join(tmpdir(), 'pptx-studio-out-'));
    const lines: string[] = [];
    const code = runRender(
      repoPath('corpus/decks/a07-text-cascade.pptx'),
      { out, json: false, quiet: false, fontDirs: [fontDir], systemFonts: false },
      (text) => lines.push(text),
    );
    expect(code).toBe(0);
    expect(lines.join('')).toContain('typeface(s) from');
  });

  it('says what the default width is, and which options the library shares', () => {
    const s = streams();
    main(['--help'], s.streams);
    expect(s.out()).toContain(`(default ${String(DEFAULT_WIDTH)})`);
    expect(s.out()).toContain('import { renderDeck } from "@pptx-studio/cli"');
  });

  it('no longer says render is unbuilt', () => {
    const s = streams();
    main(['--help'], s.streams);
    expect(s.out()).toContain('render           draw slides as SVG');
    expect(s.out()).not.toMatch(/render\s+render slides to SVG or PNG/);
  });
});
