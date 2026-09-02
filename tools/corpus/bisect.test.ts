import { readFileSync, readdirSync } from 'node:fs';
import { join, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';
import {
  deflatedEntry,
  passthroughEntry,
  readZip,
  writeZip,
  type ZipEntryInput,
} from '@pptx-studio/opc';
import { validatePackage } from '@pptx-studio/validate';
import { bisectPackages, flattenChanges, roundTripPackage } from '@pptx-studio/writer';
import { describe, expect, it } from 'vitest';

/**
 * Sub-phase 1.5's verification, in the plan's own words: *deliberately break a
 * deck three ways and confirm bisect localizes each*.
 *
 * The three are one per category the firewall recognises, because a bisector
 * that only worked on parts would look fine on two of them:
 *
 *   1. the **package stream** - `[Content_Types].xml`, which is not a part
 *   2. a **part** - `slideMaster1.xml`, with two children swapped
 *   3. a **relationship part** - one id renumbered so a reference dangles
 *
 * Each break is buried in three dozen harmless changes, which is the only
 * version of this test worth running. Handed a deck with one difference, a
 * bisector returns that difference without ever making a decision; what is
 * being checked here is that it throws the other thirty-odd away.
 *
 * The oracle is `validate` rather than PowerPoint, so this runs in CI on a
 * machine with no Office on it. The first of the three has been confirmed
 * against the real thing by hand - see ADR 0013 - and the point of `--oracle
 * powerpoint` is the failures no rule covers, which by definition cannot be
 * asserted here.
 */

const ROOT = resolve(fileURLToPath(import.meta.url), '../../..');
const DECK = join(ROOT, 'corpus', 'decks', 'a31-embedded-fonts.pptx');

const dec = (bytes: Uint8Array): string => new TextDecoder().decode(bytes);
const enc = (text: string): Uint8Array => new TextEncoder().encode(text);

const CONTENT_TYPES = '[Content_Types].xml';
const MASTER = 'ppt/slideMasters/slideMaster1.xml';
const PRESENTATION_RELS = 'ppt/_rels/presentation.xml.rels';

/** Entries whose shape names the noise renames. Everything with a `p:cNvPr`. */
const NOISY = [
  'ppt/slides/slide1.xml',
  'ppt/slides/slide2.xml',
  'ppt/slides/slide3.xml',
  'ppt/slideLayouts/slideLayout1.xml',
  'ppt/slideLayouts/slideLayout2.xml',
  MASTER,
];

type Edit = (was: string) => string;

/** Rewrite entries, leaving every other one byte-identical. */
function rewrite(bytes: Uint8Array, edits: Record<string, Edit>): Uint8Array {
  const archive = readZip(bytes);
  const entries: ZipEntryInput[] = archive.entries.map((entry) => {
    const edit = edits[entry.name];
    if (edit === undefined) return passthroughEntry(archive, entry);
    const before = dec(archive.read(entry));
    const after = edit(before);
    if (after === before) throw new Error('the edit for ' + entry.name + ' changed nothing');
    return deflatedEntry(entry.name, enc(after));
  });
  return writeZip(entries);
}

/**
 * Rename every shape, everywhere.
 *
 * Harmless: `p:cNvPr/@name` is free text and no rule looks at it. It is the
 * `@id` beside it that has to be unique, and that is left alone.
 */
const renameShapes: Edit = (was) => was.replaceAll(/ name="([^"]*)"/g, ' name="$1 (renamed)"');

function noisy(bytes: Uint8Array, extra: Record<string, Edit> = {}): Uint8Array {
  const edits: Record<string, Edit> = {};
  for (const name of NOISY) edits[name] = renameShapes;
  for (const [name, edit] of Object.entries(extra)) {
    const noise = edits[name];
    edits[name] = noise === undefined ? edit : (was) => edit(noise(was));
  }
  return rewrite(bytes, edits);
}

const original = new Uint8Array(readFileSync(DECK));

/** Which rules fire, fatally, on a package. */
function fatalRules(bytes: Uint8Array): string[] {
  return [
    ...new Set(
      validatePackage({ bytes })
        .findings.filter((finding) => finding.severity === 'fatal')
        .map((finding) => finding.rule),
    ),
  ].sort();
}

/** Entry names whose inflated bytes differ between two packages. */
function differingEntries(a: Uint8Array, b: Uint8Array): string[] {
  const left = readZip(a);
  const right = readZip(b);
  const names = new Set([...left.entries, ...right.entries].map((entry) => entry.name));
  const out: string[] = [];
  for (const name of names) {
    const one = left.get(name);
    const two = right.get(name);
    if (one === undefined || two === undefined) {
      out.push(name);
      continue;
    }
    if (dec(left.read(one)) !== dec(right.read(two))) out.push(name);
  }
  return out.sort();
}

const validateOracle = (bytes: Uint8Array) =>
  validatePackage({ bytes }).ok ? ('passes' as const) : ('fails' as const);

describe('the three deliberate breakages', () => {
  it('finds the missing content-type Default, which is not in any part', () => {
    // The canonical "PowerPoint found a problem with content" bug, and the one
    // break here that has been put to the real PowerPoint: with `OpenAndRepair`
    // off it will not open this package, and with it on it opens it repaired.
    // See ADR 0013.
    const broken = noisy(original, {
      [CONTENT_TYPES]: (was) => was.replace(/<Default Extension="fntdata"[^>]*>/, ''),
    });

    expect(fatalRules(original)).toEqual([]);
    expect(fatalRules(broken)).toContain('V002');

    const result = bisectPackages(original, broken, { oracle: validateOracle });

    expect(result.outcome).toBe('localized');
    expect(result.minimal.map((change) => change.entry)).toEqual([CONTENT_TYPES]);
    // ...and the delta it threw away was thirty times the size of the answer.
    expect(flattenChanges(result.changes).length).toBeGreaterThan(30);
    // Both rules, not just the famous one. Deleting the `Default` does not only
    // break the fntdata invariant `V002` names - it leaves five `.fntdata`
    // parts resolving to no content type at all, which is `V001`. Worth pinning
    // rather than narrowing: one edit, two rules, and a report that showed only
    // the second would understate what is wrong with the file.
    expect(fatalRules(result.bytes)).toEqual(['V001', 'V002']);
  });

  it('finds two swapped children in the slide master', () => {
    // The regression the plan cites by name: Microsoft's own Open XML SDK
    // shipped a build that merely swapped two elements in `slideMaster1.xml`,
    // and PowerPoint refused the file.
    const broken = noisy(original, {
      [MASTER]: (was) =>
        was.replace(
          '<a:off x="457200" y="274638"/><a:ext cx="11277600" cy="700088"/>',
          '<a:ext cx="11277600" cy="700088"/><a:off x="457200" y="274638"/>',
        ),
    });

    expect(fatalRules(broken)).toContain('V010');

    const result = bisectPackages(original, broken, { oracle: validateOracle });

    expect(result.outcome).toBe('localized');
    expect(new Set(result.minimal.map((change) => change.entry))).toEqual(new Set([MASTER]));
    // Two changes, and that is the correct answer rather than a failure to
    // reduce: a swap is only out of order once *both* halves are in place.
    // Putting either element back leaves the other where the schema allows it.
    expect(result.minimal).toHaveLength(2);
    for (const change of result.minimal) {
      expect(change.where).toContain('a:xfrm');
    }
    expect(fatalRules(result.bytes)).toEqual(['V010']);
  });

  it('finds the one relationship id that no longer resolves', () => {
    // `p:embeddedFont` still says `r:id="rId9"`; the `.rels` now calls that
    // relationship something else. Dangling is fatal - orphan is harmless - so
    // this is the direction that matters.
    const broken = noisy(original, {
      [PRESENTATION_RELS]: (was) => was.replace('Id="rId9"', 'Id="rId99"'),
    });

    expect(fatalRules(broken)).toContain('V006');

    const result = bisectPackages(original, broken, { oracle: validateOracle });

    expect(result.outcome).toBe('localized');
    expect(result.minimal).toHaveLength(1);
    expect(result.minimal[0]!.entry).toBe(PRESENTATION_RELS);
    expect(result.minimal[0]!.kind).toBe('attribute');
    expect(result.minimal[0]!.where).toMatch(/@Id$/);
    expect(result.minimal[0]!.was).toBe(' Id="rId9"');
    expect(result.minimal[0]!.text).toBe(' Id="rId99"');
  });
});

describe('what the search costs, and what it is worth', () => {
  it('reads the answer out of a delta of dozens in a couple of dozen runs', () => {
    const broken = noisy(original, {
      [PRESENTATION_RELS]: (was) => was.replace('Id="rId9"', 'Id="rId99"'),
    });
    const result = bisectPackages(original, broken, { oracle: validateOracle });

    // Seven entries differ and the delta runs to dozens of changes; the answer
    // is one attribute. The run count is pinned because it is the whole claim:
    // a linear scan of the leaves would cost one run each, and `ddmin` is what
    // keeps it logarithmic in the common case where one change is at fault.
    expect(result.changes).toHaveLength(7);
    expect(flattenChanges(result.changes).length).toBeGreaterThan(30);
    expect(result.runs).toBeLessThan(30);
    expect(result.unresolved).toBe(0);
  });

  it('hands back a package that still fails, and fails for the stated reason', () => {
    // The deliverable is not the report. It is this: the smallest package that
    // still shows the problem, ready to open, to keep, or to send to somebody.
    const broken = noisy(original, {
      [CONTENT_TYPES]: (was) => was.replace(/<Default Extension="fntdata"[^>]*>/, ''),
    });
    const result = bisectPackages(original, broken, { oracle: validateOracle });

    expect(validatePackage({ bytes: result.bytes }).ok).toBe(false);

    // The property worth asserting is not that the file got smaller - archive
    // size moves with compression level and says nothing - but that it differs
    // from the original in exactly one entry. That is what "smallest package
    // that still fails" means, and it is checkable entry by entry.
    expect(differingEntries(original, result.bytes)).toEqual([CONTENT_TYPES]);
    const archive = readZip(result.bytes);
    expect(dec(archive.read(archive.get('ppt/slides/slide1.xml')!))).not.toContain('(renamed)');
    // ...where the broken package it came from differed in seven.
    expect(differingEntries(original, broken)).toHaveLength(7);
  });
});

describe('bisecting a deck against our own export of it', () => {
  it('finds nothing to bisect, on every deck in the corpus', () => {
    // The one-argument form of the command, which is the workflow it exists
    // for: read a deck, write it back, and ask what we changed. The answer, on
    // all fifty-one, is nothing - which is sub-phases 1.3 and 1.4 restated at
    // the granularity this sub-phase works in, and the baseline that makes a
    // real bisection meaningful when one is ever needed.
    const failed: string[] = [];
    for (const collection of ['decks', 'authored', 'written']) {
      const dir = join(ROOT, 'corpus', collection);
      for (const name of readdirSync(dir)) {
        if (!name.endsWith('.pptx') && !name.endsWith('.pptm')) continue;
        const bytes = new Uint8Array(readFileSync(join(dir, name)));
        const result = bisectPackages(bytes, roundTripPackage(bytes).exported.bytes, {
          oracle: validateOracle,
        });
        if (result.outcome !== 'identical') failed.push(name + ': ' + result.outcome);
      }
    }
    expect(failed).toEqual([]);
  }, 120_000);
});
