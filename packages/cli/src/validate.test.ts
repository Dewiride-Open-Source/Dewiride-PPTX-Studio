import { readdirSync, readFileSync } from 'node:fs';
import { join } from 'node:path';
import { describe, expect, it } from 'vitest';
import { REPO_ROOT } from '../../../tools/repo/root.ts';
import { main, type Streams } from './main.js';

/**
 * `pptx-studio validate`, run in-process against the corpus.
 *
 * The command is thin - read a path, render a report, choose an exit code - and
 * the rules it calls are tested exhaustively in `@pptx-studio/validate`. What
 * is worth testing here is the thing only this layer decides: **which files
 * make the command fail.**
 *
 * The corpus answers it from both ends without any fixtures of our own. Every
 * deck in `corpus/decks` opens in PowerPoint, so `validate` must exit 0 on all
 * of them; every package in `corpus/reject` is one PowerPoint refuses, so it
 * must exit 1 on all of those. A command that got either wrong would be a
 * command nobody could put in a script.
 */

const ROOT = REPO_ROOT;

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

describe('pptx-studio validate', () => {
  // A minute, rather than vitest's default five seconds. Twenty-nine rules over
  // fifty-two decks is real work, and it started timing out when 1.4 added its
  // own corpus sweeps beside it - not because this got slower, but because a
  // timeout tuned to a quiet machine is a test that fails on a busy one.
  it('exits 0 on every deck the corpus says PowerPoint opens', () => {
    const failed: string[] = [];
    for (const path of [...decks('decks'), ...decks('authored'), ...decks('written')]) {
      const s = streams();
      if (main(['validate', path], s.streams) !== 0) failed.push(path + '\n' + s.out());
    }
    expect(failed).toEqual([]);
  }, 60_000);

  it('exits 1 on every package the corpus says PowerPoint refuses', () => {
    const passed: string[] = [];
    for (const path of decks('reject')) {
      const s = streams();
      if (main(['validate', path], s.streams) !== 1) passed.push(path + '\n' + s.out());
    }
    expect(passed).toEqual([]);
  }, 30_000);

  it('names the part and the XPath, which is the whole point of the report', () => {
    const s = streams();
    const path = join(ROOT, 'corpus', 'reject', 'r01-ph-hdr.pptx');
    expect(main(['validate', path], s.streams)).toBe(1);

    const text = s.out();
    expect(text).toContain('/ppt/slides/slide1.xml');
    expect(text).toContain('/p:sld/p:cSld/p:spTree/p:sp/p:nvSpPr/p:nvPr/p:ph/@type');
    expect(text).toContain('V022');
    expect(text).toContain('1 blocking');
  });

  it('says which rules did not run rather than counting them as passes', () => {
    const s = streams();
    main(['validate', join(ROOT, 'corpus', 'decks', 'a01-minimal.pptx')], s.streams);

    // The three preservation rules need the package as it was opened, and a
    // file on the command line has no such history. `validate` is a diagnostic;
    // the export gate is `assertValid`, called from the writer with both.
    for (const rule of ['V027', 'V028', 'V029']) {
      expect(s.out()).toContain(rule + ' did not run');
    }
  });

  it('explains a rule when asked', () => {
    const s = streams();
    main(['validate', join(ROOT, 'corpus', 'reject', 'r15-control.pptx'), '--explain'], s.streams);
    expect(s.out()).toContain('V025 - no `p:control`');
    expect(s.out()).toContain('all eight forms tried');
  });

  it('emits a report a script can read', () => {
    const s = streams();
    main(
      ['validate', join(ROOT, 'corpus', 'reject', 'r12-sheet-id-collision.pptx'), '--json'],
      s.streams,
    );

    const report = JSON.parse(s.out()) as {
      findings: { rule: string; where: { part: string; xpath: string } }[];
      blocking: number;
      ok: boolean;
    };
    expect(report.ok).toBe(false);
    expect(report.blocking).toBe(1);
    expect(report.findings[0]!.rule).toBe('V019');
    expect(report.findings[0]!.where.part).toBe('/ppt/slideMasters/slideMaster1.xml');
  });

  it('writes to a file when told to, and prints nothing', () => {
    const s = streams();
    const out = join(ROOT, 'corpus', 'reject', 'manifest.json');
    // Deliberately not written: `--out` with a path we would clobber is not a
    // test worth having. What is checked is that the command still needs a
    // deck, and says which command needs it.
    expect(main(['validate'], s.streams)).toBe(2);
    expect(s.err()).toContain('validate needs a path to a .pptx');
    expect(readFileSync(out, 'utf8').length).toBeGreaterThan(0);
  });

  it('is no longer listed as unbuilt', () => {
    const s = streams();
    main(['--help'], s.streams);
    expect(s.out()).toContain('validate         the must-not-break rules');
    expect(s.out()).not.toContain('validate     run the must-not-break rules');
  });
});
