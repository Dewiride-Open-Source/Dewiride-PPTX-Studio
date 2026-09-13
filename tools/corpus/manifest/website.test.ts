import { describe, expect, it } from 'vitest';

import type { CorpusFile } from './check.ts';
import { checkWebsiteCopies, readWebsiteManifest, type WebsiteDeck } from './website.ts';

const SHA_A = 'a'.repeat(64);
const SHA_B = 'b'.repeat(64);
const MANIFEST = 'website/public/decks/manifest.json';

function deck(overrides: Partial<WebsiteDeck> = {}): WebsiteDeck {
  return {
    file: 'a01-example.pptx',
    sha256: SHA_A,
    license: 'CC0-1.0',
    source: 'self-authored',
    about: 'One shape.',
    ...overrides,
  };
}

const file = (path: string, sha256: string): CorpusFile => ({ path, bytes: 1024, sha256 });

const CLEAN = {
  manifestPath: MANIFEST,
  decks: [deck()],
  copies: [file('website/public/decks/a01-example.pptx', SHA_A)],
  corpus: [file('corpus/decks/a01-example.pptx', SHA_A)],
};

describe('C020-website-copy', () => {
  it('passes a copy that is byte-identical to a corpus deck and claimed with CC0 self-authored', () => {
    expect(checkWebsiteCopies(CLEAN)).toEqual([]);
  });

  it('fires when the copy differs from every corpus file', () => {
    const found = checkWebsiteCopies({
      ...CLEAN,
      decks: [deck({ sha256: SHA_B })],
      copies: [file('website/public/decks/a01-example.pptx', SHA_B)],
    });
    expect(found.map((one) => one.rule)).toEqual(['C020-website-copy']);
    expect(found[0]?.message).toContain('byte-identical to no file under corpus/');
  });

  it('fires when the manifest digest is not the file digest', () => {
    const found = checkWebsiteCopies({ ...CLEAN, decks: [deck({ sha256: SHA_B })] });
    expect(found).toHaveLength(1);
    expect(found[0]?.message).toContain(`sha256 is ${SHA_B}`);
  });

  it('fires on a licence or source a public site may not serve', () => {
    for (const overrides of [{ license: 'CC-BY-4.0' }, { source: 'third-party' }]) {
      const found = checkWebsiteCopies({ ...CLEAN, decks: [deck(overrides)] });
      expect(found).toHaveLength(1);
      expect(found[0]?.message).toContain('a public site serves');
    }
  });

  it('fires on a deck the manifest does not claim, and on a claim with no file', () => {
    const orphan = checkWebsiteCopies({
      ...CLEAN,
      copies: [...CLEAN.copies, file('website/public/decks/b02-extra.pptx', SHA_B)],
    });
    expect(orphan.map((one) => one.where)).toEqual(['website/public/decks/b02-extra.pptx']);

    const missing = checkWebsiteCopies({ ...CLEAN, copies: [] });
    expect(missing).toHaveLength(1);
    expect(missing[0]?.message).toContain('is not under website/public/decks/');
  });

  it('reads a manifest and refuses one with a missing field', () => {
    expect(readWebsiteManifest({ decks: [deck()] }, MANIFEST)).toEqual([deck()]);
    expect(() => readWebsiteManifest({ decks: [{ ...deck(), about: '' }] }, MANIFEST)).toThrow(
      /decks\/0\/about/,
    );
    expect(() => readWebsiteManifest({}, MANIFEST)).toThrow(/expected \{ decks/);
  });
});
