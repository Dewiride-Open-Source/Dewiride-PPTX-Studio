import { mkdtempSync, readFileSync, readdirSync, rmSync, writeFileSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';
import { afterAll, beforeAll, describe, expect, it } from 'vitest';
import type { RoundTripReport } from '@pptx-studio/writer';
import { main, type Streams } from './main.js';
import { formatRoundTrip, type RoundTripStats } from './roundtrip.js';

/**
 * `pptx-studio roundtrip`, run in-process against the corpus.
 *
 * The comparison itself is `@pptx-studio/writer`'s and is tested there and in
 * `tools/corpus/roundtrip.test.ts`. What this file is for is the part only the
 * command line decides: **which files make it exit non-zero**, and whether what
 * it prints is any use to the person reading it.
 *
 * The corpus answers the first from both ends with no fixtures of our own:
 * every deck in `corpus/decks` opens in PowerPoint and must round-trip, and so
 * must every package in `corpus/reject` that PowerPoint will not open - because
 * handing back a broken file unchanged is preservation working, not failing.
 */

const ROOT = resolve(fileURLToPath(import.meta.url), '../../../..');
let directory: string;

function streams(): { streams: Streams; out: () => string; err: () => string } {
  const outParts: string[] = [];
  const errParts: string[] = [];
  return {
    streams: {
      out: (text) => outParts.push(text),
      err: (text) => errParts.push(text),
    },
    out: () => outParts.join(''),
    err: () => errParts.join(''),
  };
}

function decks(collection: string): string[] {
  const dir = join(ROOT, 'corpus', collection);
  return readdirSync(dir)
    .filter((name) => name.endsWith('.pptx') || name.endsWith('.pptm'))
    .map((name) => join(dir, name));
}

beforeAll(() => {
  directory = mkdtempSync(join(tmpdir(), 'pptx-roundtrip-'));
});

afterAll(() => {
  rmSync(directory, { recursive: true, force: true });
});

describe('pptx-studio roundtrip', () => {
  // A minute, rather than vitest's default five seconds. Fifty-one decks read,
  // exported and then compared part by part is real work - about three seconds
  // alone and more with the rest of the suite running beside it - and a timeout
  // tuned to a quiet machine is a test that fails on a busy one.
  it('exits 0 on every deck the corpus says PowerPoint opens', () => {
    const failed: string[] = [];
    for (const path of [...decks('decks'), ...decks('authored'), ...decks('written')]) {
      const s = streams();
      if (main(['roundtrip', path], s.streams) !== 0) failed.push(path + '\n' + s.out());
    }
    expect(failed).toEqual([]);
  }, 60_000);

  it('reports what it compared and how', () => {
    const s = streams();
    const path = join(ROOT, 'corpus', 'decks', 'a01-minimal.pptx');
    expect(main(['roundtrip', path], s.streams)).toBe(0);

    const text = s.out();
    expect(text).toContain('read      18272 bytes, 18 entries');
    expect(text).toContain('0 part(s) re-serialized');
    expect(text).toContain('parts identical');
    expect(text).toContain('xml,');
    expect(text).toContain('rels,');
    expect(text).toContain('no differences');
  });

  it('exits 0 on every package PowerPoint refuses, which is the right answer', () => {
    // Worth an assertion because it looks like a bug and is not. `corpus/reject`
    // holds seventeen packages PowerPoint will not open, and `validate` exits 1
    // on all of them - but round-tripping one is supposed to hand back the same
    // broken package, defect intact. A `roundtrip` that failed here would be
    // saying it had *changed* the file, which is the one thing it must not do.
    const wrong: string[] = [];
    for (const path of decks('reject')) {
      const s = streams();
      if (main(['roundtrip', path], s.streams) !== 0) wrong.push(path + '\n' + s.err() + s.out());
    }
    expect(wrong).toEqual([]);
  }, 30_000);

  it('shows only the differences under --quiet', () => {
    const s = streams();
    const path = join(ROOT, 'corpus', 'decks', 'a01-minimal.pptx');
    expect(main(['roundtrip', path, '--quiet'], s.streams)).toBe(0);
    expect(s.out()).toBe('  no differences\n');
  });

  it('saves the package it wrote when asked, so it can be opened in PowerPoint', () => {
    // Until 1.5 scripts the COM loop there is no substitute for opening the
    // file by hand, and no way to do that without this flag.
    const path = join(ROOT, 'corpus', 'decks', 'a01-minimal.pptx');
    const out = join(directory, 'written.pptx');
    const s = streams();
    expect(main(['roundtrip', path, '--write', out], s.streams)).toBe(0);

    // Not byte-identical to the input, and that is this sub-phase's whole
    // premise rather than a disappointment: `passthroughEntry` does not carry
    // an entry's DOS timestamp or external attributes, because those describe
    // where a file came from and not what is in it. What must hold is that the
    // saved file is the same deck, so it is round-tripped in turn.
    const written = readFileSync(out);
    expect(written.length).toBeGreaterThan(0);
    expect(Buffer.compare(written, readFileSync(path))).not.toBe(0);

    const again = streams();
    expect(main(['roundtrip', out], again.streams)).toBe(0);
    expect(again.out()).toContain('no differences');
  });

  it('emits a comparison a script can read, with a digest per part', () => {
    const s = streams();
    main(['roundtrip', join(ROOT, 'corpus', 'decks', 'a01-minimal.pptx'), '--json'], s.streams);

    const report = JSON.parse(s.out()) as {
      ok: boolean;
      counts: { xml: number; binary: number; relationships: number; same: number };
      differences: unknown[];
      parts: { part: string; how: string; same: boolean; digest: string }[];
    };
    expect(report.ok).toBe(true);
    expect(report.differences).toEqual([]);
    expect(report.parts).toHaveLength(
      report.counts.xml + report.counts.binary + report.counts.relationships,
    );
    for (const part of report.parts) {
      expect(part.same, part.part).toBe(true);
      expect(part.digest, part.part).toMatch(/^[0-9a-f]{64}$/);
    }
  });

  it('writes the report to a file when told to, and prints nothing', () => {
    const out = join(directory, 'report.txt');
    const s = streams();
    expect(
      main(
        ['roundtrip', join(ROOT, 'corpus', 'decks', 'a01-minimal.pptx'), '--out', out],
        s.streams,
      ),
    ).toBe(0);
    expect(s.out()).toBe('');
    expect(readFileSync(out, 'utf8')).toContain('no differences');
  });

  it('says a file is not a package rather than reporting it as a difference', () => {
    // Three ways to fail and one exit code, so stderr has to carry which. This
    // is the first of them; the other two are an export the firewall refused and
    // a comparison that found something, and each needs different work next.
    const path = join(directory, 'not-a-deck.pptx');
    writeFileSync(path, 'PK and then nothing that follows');

    const s = streams();
    expect(main(['roundtrip', path], s.streams)).toBe(1);
    expect(s.err()).toContain('ERR_');
    expect(s.out()).toBe('');
  });

  it('needs a path, and says which command needs it', () => {
    const s = streams();
    expect(main(['roundtrip'], s.streams)).toBe(2);
    expect(s.err()).toContain('roundtrip needs a path to a .pptx');
  });

  it('is no longer listed as unbuilt', () => {
    const s = streams();
    main(['--help'], s.streams);
    expect(s.out()).toContain('roundtrip        read a deck, write it back');
    expect(s.out()).not.toContain('(sub-phase 1.4)');
  });
});

describe('the report a difference produces', () => {
  /**
   * The one output nothing else in this repository can exercise.
   *
   * No deck in the corpus makes this writer produce a difference - that is the
   * result of the whole phase - so the branch that renders one has no natural
   * test, and it is the branch a person reads at the worst moment. The
   * comparison here is synthetic on purpose: the finding of differences is
   * tested in `@pptx-studio/writer`, and what is tested here is only that the
   * printing of them is legible.
   */
  const stats: RoundTripStats = {
    bytesIn: 1000,
    bytesOut: 990,
    entriesIn: 12,
    entriesOut: 12,
    rewritten: 1,
    streamed: 10,
  };

  const comparison: RoundTripReport = {
    ok: false,
    parts: [],
    relabelled: [{ source: '/ppt/slides/slide1.xml', from: 'rId2', to: 'rId9', target: 'image' }],
    counts: { xml: 8, binary: 1, relationships: 2, same: 10 },
    differences: [
      {
        kind: 'xml',
        part: '/ppt/slides/slide1.xml',
        detail: 'the canonical XML differs at character 812\n    original: a\n    written:  b',
      },
      { kind: 'binary', part: '/ppt/media/image1.png', detail: 'the bytes differ' },
    ],
  };

  it('names every difference, its kind and its part, on one line', () => {
    const lines = formatRoundTrip('deck.pptx', stats, comparison, false).split('\n');
    // Asserted by shape rather than by column, so that changing the width of
    // the kind column is a formatting change and not a test failure.
    const heading = (kind: string, part: string): string | undefined =>
      lines.find((line) => line.trimStart().startsWith(kind) && line.endsWith(part));

    expect(lines).toContain('  2 difference(s)');
    expect(heading('xml', '/ppt/slides/slide1.xml')).toBeDefined();
    expect(heading('binary', '/ppt/media/image1.png')).toBeDefined();
  });

  it('indents a multi-line detail under the part it belongs to', () => {
    // A canonical-XML difference carries an excerpt of both sides, so the
    // detail is three lines and has to stay visibly attached to its heading.
    const text = formatRoundTrip('deck.pptx', stats, comparison, false);
    expect(text).toContain('    the canonical XML differs at character 812');
    expect(text).toContain('        original: a');
  });

  it('mentions renamed ids, which mean the file came from somewhere else', () => {
    const text = formatRoundTrip('deck.pptx', stats, comparison, false);
    expect(text).toContain('1 relationship id(s) renamed');
    expect(text).toContain('/ppt/slides/slide1.xml  rId2 -> rId9');
  });

  it('drops the tally and the renames under --quiet', () => {
    const text = formatRoundTrip('deck.pptx', stats, comparison, true);
    expect(text).not.toContain('roundtrip deck.pptx');
    expect(text).not.toContain('renamed');
    expect(text).toContain('2 difference(s)');
  });
});
