import { existsSync, mkdtempSync, readFileSync, rmSync, writeFileSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import {
  deflatedEntry,
  passthroughEntry,
  readZip,
  writeZip,
  type ZipEntryInput,
} from '@pptx-studio/opc';
import type { BisectResult } from '@pptx-studio/writer';
import { afterAll, beforeAll, describe, expect, it } from 'vitest';
import { REPO_ROOT } from '../../../tools/repo/root.ts';
import { formatBisect, oracleScriptPath, type BisectStats } from './bisect.js';
import { main, type Streams } from './main.js';

/**
 * `pptx-studio bisect`, run in-process.
 *
 * The search is `@pptx-studio/writer`'s and is tested there and in
 * `tools/corpus/roundtrip/bisect.test.ts`. What is left for this file is the part only the
 * command line owns: which files make it exit non-zero, how an oracle's exit
 * code becomes a verdict, and whether what it prints is any use.
 *
 * Nothing here runs the PowerPoint oracle. It needs Office, a Windows COM
 * server and about a second and a half per candidate, none of which belong in a
 * test suite - so what is asserted about it is the one thing that can break
 * silently and would not be noticed until someone installed the package: that
 * the script it shells out to is where the code thinks it is.
 */

const ROOT = REPO_ROOT;
const DECK = join(ROOT, 'corpus', 'decks', 'a31-embedded-fonts.pptx');

let directory: string;
let original: string;
let broken: string;

const dec = (bytes: Uint8Array): string => new TextDecoder().decode(bytes);
const enc = (text: string): Uint8Array => new TextEncoder().encode(text);

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

/** The corpus deck with its fntdata content type gone, plus harmless noise. */
function damage(bytes: Uint8Array): Uint8Array {
  const archive = readZip(bytes);
  const entries: ZipEntryInput[] = archive.entries.map((entry) => {
    if (entry.name === '[Content_Types].xml') {
      return deflatedEntry(
        entry.name,
        enc(dec(archive.read(entry)).replace(/<Default Extension="fntdata"[^>]*>/, '')),
      );
    }
    if (entry.name.startsWith('ppt/slides/slide')) {
      return deflatedEntry(
        entry.name,
        enc(dec(archive.read(entry)).replaceAll(/ name="([^"]*)"/g, ' name="$1 (renamed)"')),
      );
    }
    return passthroughEntry(archive, entry);
  });
  return writeZip(entries);
}

beforeAll(() => {
  directory = mkdtempSync(join(tmpdir(), 'pptx-bisect-cli-'));
  original = join(directory, 'original.pptx');
  broken = join(directory, 'broken.pptx');
  const bytes = new Uint8Array(readFileSync(DECK));
  writeFileSync(original, bytes);
  writeFileSync(broken, damage(bytes));
});

afterAll(() => {
  rmSync(directory, { recursive: true, force: true });
});

describe('pptx-studio bisect', () => {
  it('narrows a broken deck to one change, and exits 1', () => {
    const s = streams();
    expect(main(['bisect', original, broken], s.streams)).toBe(1);

    const text = s.out();
    expect(text).toContain('oracle    validate');
    expect(text).toContain('1 change(s) in 1 entry(s)');
    expect(text).toContain('removed  /[Content_Types].xml');
    expect(text).toContain('<Default Extension="fntdata"');
    // ...and it says how much it threw away, which is the claim.
    expect(text).toMatch(/from a delta of \d\d+/);
  });

  it('finds nothing to bisect between a deck and our own export of it', () => {
    // The one-argument form. Exits 0 because there is nothing wrong, which is
    // the answer sub-phases 1.3 and 1.4 already earned and this one inherits.
    const s = streams();
    expect(main(['bisect', original], s.streams)).toBe(0);
    expect(s.out()).toContain('the same entries, byte for byte');
  });

  it('says so when the second deck is not broken either', () => {
    const s = streams();
    const other = join(directory, 'renamed-only.pptx');
    const bytes = new Uint8Array(readFileSync(DECK));
    const archive = readZip(bytes);
    writeFileSync(
      other,
      writeZip(
        archive.entries.map((entry) =>
          entry.name === 'ppt/slides/slide1.xml'
            ? deflatedEntry(
                entry.name,
                enc(dec(archive.read(entry)).replaceAll(' name="', ' name="x')),
              )
            : passthroughEntry(archive, entry),
        ),
      ),
    );

    const s2 = streams();
    expect(main(['bisect', original, other], s2.streams)).toBe(0);
    expect(s2.out()).toContain('nothing to look for');
    expect(s.out()).toBe('');
  });

  it('emits a result a script can read', () => {
    const s = streams();
    expect(main(['bisect', original, broken, '--json'], s.streams)).toBe(1);

    const report = JSON.parse(s.out()) as {
      outcome: string;
      runs: number;
      delta: number;
      minimal: { entry: string; kind: string; where: string }[];
    };
    expect(report.outcome).toBe('localized');
    expect(report.minimal).toHaveLength(1);
    expect(report.minimal[0]!.entry).toBe('[Content_Types].xml');
    expect(report.delta).toBeGreaterThan(report.minimal.length);
    expect(report.runs).toBeGreaterThan(0);
  });

  it('saves the smallest failing package when asked', () => {
    // The deliverable: a package small enough to look at, that still fails.
    const out = join(directory, 'minimal.pptx');
    const s = streams();
    expect(main(['bisect', original, broken, '--write', out], s.streams)).toBe(1);

    const saved = new Uint8Array(readFileSync(out));
    const archive = readZip(saved);
    expect(dec(archive.read(archive.get('[Content_Types].xml')!))).not.toContain('fntdata');
    // The noise is gone: the slide came back the way the original wrote it.
    expect(dec(archive.read(archive.get('ppt/slides/slide1.xml')!))).not.toContain('(renamed)');
  });

  it('writes the report to a file when told to, and prints nothing', () => {
    const out = join(directory, 'report.txt');
    const s = streams();
    expect(main(['bisect', original, broken, '--out', out], s.streams)).toBe(1);
    expect(s.out()).toBe('');
    expect(readFileSync(out, 'utf8')).toContain('[Content_Types].xml');
  });

  it('shows a line per oracle run under --progress', () => {
    const s = streams();
    main(['bisect', original, broken, '--progress'], s.streams);
    expect(s.out()).toContain('run    1  0/');
    expect(s.out()).toContain('change(s)');
  });
});

describe('choosing an oracle', () => {
  it('turns a command exit code into a verdict', () => {
    // 0 passes, 1 fails, anything else is not an answer. Here the command sees
    // the candidate and always calls it broken, so nothing can be reduced away
    // and the whole delta comes back as the answer.
    const s = streams();
    const code = main(
      [
        'bisect',
        original,
        broken,
        '--oracle',
        'command',
        '--command',
        'node -e "process.exit(require(\'fs\').existsSync(process.argv[1]) ? 1 : 0)" {}',
      ],
      s.streams,
    );
    expect(code).toBe(1);
    expect(s.out()).toContain('oracle    command');
  }, 60_000);

  it('treats an exit code that is neither 0 nor 1 as no answer at all', () => {
    // The distinction that keeps a broken harness from being reported as a
    // broken deck. A command that cannot run exits 127 or 2; reading that as
    // "fails" would blame whichever changes happened to be applied.
    const s = streams();
    main(
      [
        'bisect',
        original,
        broken,
        '--oracle',
        'command',
        '--command',
        'node -e "process.exit(2)" {}',
      ],
      s.streams,
    );
    expect(s.out()).toContain('unresolved');
  }, 60_000);

  it('refuses an oracle it does not have', () => {
    const s = streams();
    expect(main(['bisect', original, broken, '--oracle', 'vibes'], s.streams)).toBe(2);
    expect(s.err()).toContain('--oracle wants validate, powerpoint or command');
  });

  it('refuses --oracle command with nothing to run', () => {
    const s = streams();
    expect(main(['bisect', original, broken, '--oracle', 'command'], s.streams)).toBe(2);
    expect(s.err()).toContain('needs --command');
  });

  it('refuses a ceiling that is not a positive number', () => {
    const s = streams();
    expect(main(['bisect', original, broken, '--max-runs', 'lots'], s.streams)).toBe(2);
    expect(s.err()).toContain('positive integers');
  });

  it('ships the PowerShell the PowerPoint oracle shells out to', () => {
    // Nothing else would notice this breaking. The path is resolved relative to
    // the module, so it survives the move from `src/` to `dist/` - but only as
    // long as `scripts/` stays in the package's `files`, and a missing file
    // would show up as every candidate being "unresolved" on a user's machine.
    expect(existsSync(oracleScriptPath())).toBe(true);
    expect(readFileSync(oracleScriptPath(), 'utf8')).toContain('Open2007');
  });

  it('is no longer listed as unbuilt', () => {
    const s = streams();
    main(['--help'], s.streams);
    expect(s.out()).toContain('bisect           narrow a broken deck');
    expect(s.out()).not.toContain('(sub-phase 1.5)');
  });
});

describe('the reports a green corpus cannot produce', () => {
  /**
   * Three of the four outcomes never happen on a working deck, and they are the
   * three a person reads when something has gone wrong. So they are rendered
   * from synthetic results here - the finding of them is tested in the writer,
   * and what is tested here is only that they are legible.
   */
  const stats: BisectStats = {
    original: 'a.pptx',
    broken: 'b.pptx',
    bytesOriginal: 1000,
    bytesBroken: 900,
    entriesOriginal: 12,
    entriesBroken: 12,
    oracle: 'powerpoint',
  };

  const base: BisectResult = {
    outcome: 'localized',
    changes: [],
    minimal: [],
    bytes: new Uint8Array(0),
    runs: 9,
    cached: 2,
    unresolved: 0,
    exhausted: false,
    truncated: false,
  };

  it('tells you to look at the original when the original is the problem', () => {
    const text = formatBisect(stats, { ...base, outcome: 'original-fails' }, false);
    expect(text).toContain('the *original* fails the oracle');
    expect(text).toContain('look at the original package first');
  });

  it('says the answer is not minimal when the ceiling stopped it', () => {
    const text = formatBisect(stats, { ...base, exhausted: true }, false);
    expect(text).toContain('stopped at the ceiling');
    expect(text).toContain('--max-runs');
  });

  it('says when the delta was too big to decompose', () => {
    const text = formatBisect(stats, { ...base, truncated: true }, false);
    expect(text).toContain('too large to decompose');
  });

  it('counts the runs it did not have to make', () => {
    expect(formatBisect(stats, base, false)).toContain('9 run(s), 2 from cache');
  });

  it('drops the tally under --quiet', () => {
    const text = formatBisect(stats, { ...base, outcome: 'identical' }, true);
    expect(text).not.toContain('bisect a.pptx');
    expect(text).toContain('byte for byte');
  });
});
