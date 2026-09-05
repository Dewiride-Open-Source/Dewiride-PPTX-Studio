import { createHash } from 'node:crypto';
import { readFileSync } from 'node:fs';
import { join, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';
import { censusPackage } from '@pptx-studio/census';
import { describe, expect, it } from 'vitest';
import { AUTHORED_DECKS } from './decks.ts';
import { readZip } from '../../ground-truth/zip.ts';

const CORPUS = resolve(fileURLToPath(import.meta.url), '../../../../corpus/authored');

/**
 * `C-CENSUS` for Tier B, plus the four constraints the roster says are enforced
 * rather than trusted.
 *
 * There is no `C-REGEN` here and there cannot be. Tier A's decks rebuild
 * byte-for-byte, so a test can hold the generator to its own output; these are
 * bytes PowerPoint wrote, and re-running the authoring script within the same
 * second produces a different file because `dcterms:created` is a timestamp.
 * What is checkable is that the committed bytes are the ones the manifest
 * claims, that the census matches what was reviewed, and that nothing the roster
 * forbids got in.
 *
 * The privacy test is the one that earns its place. The roster originally said
 * to scrub through `BuiltInDocumentProperties`, which cannot be reached from
 * PowerShell at all on this machine - the collection enumerates but every member
 * access throws NullReferenceException, so a scrub written that way runs, reports
 * nothing, and changes nothing. The first build of `b01` went out carrying the
 * author's name twice. This test asserts the *result* rather than the method, so
 * it keeps holding if the method has to change again.
 */

function load(id: string): Uint8Array {
  return new Uint8Array(readFileSync(join(CORPUS, id + '.pptx')));
}

/**
 * Strings that must not appear anywhere in a committed deck. Everything here is
 * either the authoring account, its organisation, or the machine - the three
 * things `docProps` picks up without being asked.
 */
const MUST_NOT_APPEAR = ['Jagdish', 'Kumawat', 'Dewiride Technologies', 'AzureAD', 'JAGDIS~1'];

describe.each(AUTHORED_DECKS.map((deck) => [deck.id, deck] as const))('%s', (id, deck) => {
  it('reports exactly the census its entry declares', () => {
    const census = censusPackage(load(id));
    const found: Record<string, number> = {};
    for (const feature of census.features) found[feature.key] = feature.count;

    expect(found).toEqual(deck.features);
  });

  it('is a package with nothing wrong with it', () => {
    const census = censusPackage(load(id));

    expect(census.problems).toEqual([]);
    expect(census.relationships.dangling).toEqual([]);
    expect(census.relationships.unreachableParts).toEqual([]);
  });

  it('is the deck the manifest claims it is', () => {
    const manifest = JSON.parse(readFileSync(join(CORPUS, 'manifest.json'), 'utf8')) as {
      entries: { id: string; sha256: string; bytes: number; slides: number }[];
    };
    const claimed = manifest.entries.find((row) => row.id === id);
    const bytes = load(id);

    expect(claimed, 'no manifest entry for ' + id).toBeDefined();
    expect(claimed?.sha256).toBe(createHash('sha256').update(bytes).digest('hex'));
    expect(claimed?.bytes).toBe(bytes.byteLength);
    expect(claimed?.slides).toBe(deck.slides);
  });

  it('carries no author, company or machine name', () => {
    // The whole archive, not just docProps. A name can reach `docProps/app.xml`
    // through `TitlesOfParts`, a slide through a shape name, and `core.xml`
    // through two separate elements; sweeping the decompressed bytes catches
    // all of those without having to guess which one is next.
    const parts = readZip(load(id));
    const text = new TextDecoder().decode(load(id));
    const found: string[] = [];

    for (const needle of MUST_NOT_APPEAR) {
      if (text.includes(needle)) found.push(needle + ' (raw archive)');
      for (const part of parts) {
        if (!part.name.endsWith('.xml') && !part.name.endsWith('.rels')) continue;
        if (new TextDecoder().decode(part.bytes).includes(needle)) {
          found.push(needle + ' in ' + part.name);
        }
      }
    }

    expect(found).toEqual([]);
  });

  it('has an empty dc:creator and cp:lastModifiedBy', () => {
    // The positive form of the test above. `RemoveDocumentInformation` empties
    // these two elements rather than removing them, so "absent" would pass for
    // the wrong reason and is not what is asserted.
    const core = readZip(load(id)).find((part) => part.name === 'docProps/core.xml');
    expect(core, 'no docProps/core.xml').toBeDefined();
    const xml = new TextDecoder().decode(core?.bytes ?? new Uint8Array());

    expect(xml).toContain('<dc:creator></dc:creator>');
    expect(xml).toContain('<cp:lastModifiedBy></cp:lastModifiedBy>');
  });

  it('embeds no fonts', () => {
    // `SaveAs`'s third argument is `EmbedTrueTypeFonts` and it is msoFalse on
    // every call in the authoring script. This checks the consequence: a
    // `ppt/fonts/` part would be a licensing problem, because the faces on this
    // machine are Microsoft's and are not ours to redistribute.
    const names = readZip(load(id)).map((part) => part.name);

    expect(names.filter((name) => name.startsWith('ppt/fonts/'))).toEqual([]);
    expect(names.filter((name) => name.endsWith('.fntdata'))).toEqual([]);
  });
});

describe('Tier B as a whole', () => {
  it('has unique, kebab-case, sorted ids', () => {
    const ids = AUTHORED_DECKS.map((deck) => deck.id);

    expect(new Set(ids).size).toBe(ids.length);
    expect([...ids].sort((a, b) => a.localeCompare(b))).toEqual(ids);
    for (const id of ids) expect(id).toMatch(/^b\d\d-[a-z0-9-]+$/);
  });

  it('was written by one producer on one build', () => {
    // A collection assembled from two PowerPoint builds would be a corpus whose
    // lexical claims are about nothing in particular. If Tier B is ever
    // re-authored it is re-authored whole.
    const manifest = JSON.parse(readFileSync(join(CORPUS, 'manifest.json'), 'utf8')) as {
      measuredOn: { build: string };
      entries: { producer: { build: string } }[];
    };

    for (const entry of manifest.entries) {
      expect(entry.producer.build).toBe(manifest.measuredOn.build);
    }
  });

  it('claims no recipe, because none could be honoured', () => {
    const manifest = JSON.parse(readFileSync(join(CORPUS, 'manifest.json'), 'utf8')) as {
      entries: (Record<string, unknown> & { id: string })[];
    };

    for (const entry of manifest.entries) {
      expect(entry['recipe'], entry.id + ' must not claim a recipe').toBeUndefined();
      expect(entry['sourceNote'], entry.id + ' must say why').toBeDefined();
    }
  });

  it('shares the floor b01-blank establishes', () => {
    // Every Tier B deck inherits PowerPoint's chassis, so the interesting part
    // of any entry is what exceeds this. If the floor ever moves - a new build
    // emitting a twelfth layout, say - this fails once rather than nine times
    // with no explanation.
    const floor = AUTHORED_DECKS.find((deck) => deck.id === 'b01-blank')?.features ?? {};

    for (const deck of AUTHORED_DECKS) {
      for (const [key, count] of Object.entries(floor)) {
        // b04 replaces a placeholder with a graphic frame and b02 adds ten
        // slides, so counts move in both directions; the keys are what is
        // shared, not their values.
        expect(Object.keys(deck.features), deck.id + ' is missing ' + key).toContain(key);
        expect(typeof count).toBe('number');
      }
    }
  });
});
