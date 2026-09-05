#!/usr/bin/env node
/**
 * Write `corpus/authored/manifest.json` from the committed Tier B bytes.
 *
 *   node tools/corpus/tiers/b-authored/write-manifest.ts
 *   npx prettier --write corpus/authored/manifest.json
 *
 * The Prettier line is not optional, for the same reason it is not optional for
 * Tier A: `JSON.stringify(…, null, 2)` puts every array element on its own line
 * and Prettier collapses short ones, so `pnpm format:check` fails on a freshly
 * written manifest. Nothing downstream cares which form it is in - `C010` and
 * `C011` hash the `.pptx` files, not the manifest.
 *
 * Unlike `build-probes.ts --manifest`, this does NOT build anything. It reads
 * what is already committed under `corpus/authored/` and describes it. That
 * asymmetry is the whole difference between the two tiers: Tier A's manifest is
 * a statement about a recipe, and this one is a statement about bytes that
 * cannot be regenerated. Running this after re-running `build-tier-b.ps1` is
 * therefore a deliberate act of replacing the corpus, not a refresh - every
 * `sha256` will change even if nothing about the decks did, because
 * `dcterms:created` is a timestamp.
 */

import { createHash } from 'node:crypto';
import { readFileSync, writeFileSync } from 'node:fs';
import { join } from 'node:path';

import { repoPath } from '../../../repo/root.ts';
import { AUTHORED_DECKS } from './decks.ts';

const CORPUS = repoPath('corpus/authored');

/**
 * The machine and build every one of these came off. Recorded per manifest
 * rather than per entry because it is the same for all nine, and because a
 * later re-authoring on a different build should replace the whole collection
 * rather than leave a mixture nobody can reason about.
 */
const PRODUCER = {
  application: 'Microsoft PowerPoint',
  version: '16.0',
  build: '20326',
  platform: 'Windows 11 Enterprise 10.0.26200',
};

/**
 * The same producer as one string, for `C018`.
 *
 * Derived rather than typed out, because `corpus/written` declares this exact
 * spelling as its own XML serializer and `C-LEX` counts the two as one. Two
 * literals that must match are two literals that eventually will not.
 */
const SERIALIZER =
  PRODUCER.application + ' ' + PRODUCER.version + '.' + String(PRODUCER.build ?? '');

const entries = AUTHORED_DECKS.map((deck) => {
  const path = deck.id + '.pptx';
  const bytes = readFileSync(join(CORPUS, path));
  return {
    id: deck.id,
    kind: 'deck',
    storage: 'committed',
    path,
    license: 'CC0-1.0',
    source: 'self-authored',
    producer: PRODUCER,
    sha256: createHash('sha256').update(bytes).digest('hex'),
    bytes: bytes.byteLength,
    slides: deck.slides,
    format: 'pptx',
    features: deck.features,
    description: deck.description,
    sourceNote:
      'Authored by tools/corpus/tiers/b-authored/build-tier-b.ps1 against the producer above. No recipe: ' +
      'PowerPoint stamps dcterms:created into docProps/core.xml on every save, so re-running the ' +
      'script produces different bytes within the same second, and the monthly build changes the ' +
      'markup underneath. Personal information was removed at author time with ' +
      'RemoveDocumentInformation(ppRDIRemovePersonalInformation) before the save, never by editing ' +
      'the saved file - which would have destroyed the exact Microsoft serialization these decks ' +
      'exist to capture.',
    addedIn: '1.1',
  };
});

const manifest = {
  $comment:
    'Tier B of the sub-phase 1.1 corpus: decks Microsoft PowerPoint wrote, committed as bytes and ' +
    'NOT reproducible. The authoring is auditable instead of the output being reproducible - ' +
    'tools/corpus/tiers/b-authored/build-tier-b.ps1 is the provenance, line by line. This collection is ' +
    'separate from corpus/decks because build-probes.ts --manifest rewrites that one in full.',
  manifestVersion: 1,
  collection: 'authored',
  generatedIn: '1.1',
  adr: 'docs/adr/phase-1-round-trip/0009-the-corpus.md',
  generator: 'tools/corpus/tiers/b-authored/write-manifest.ts',
  // C018. Both layers are PowerPoint here, and the string must match the one
  // corpus/written declares for its XML: Tier C's parts are these parts, so
  // C-LEX has to count them once.
  serializers: { xml: SERIALIZER, container: SERIALIZER },
  targetCount: 9,
  measuredOn: PRODUCER,
  entries: [...entries].sort((a, b) => a.id.localeCompare(b.id)),
};

const path = join(CORPUS, 'manifest.json');
writeFileSync(path, JSON.stringify(manifest, null, 2) + '\n');
console.log(
  'wrote ' +
    path +
    '  (' +
    String(entries.length) +
    ' entries, ' +
    String(entries.reduce((n, e) => n + e.bytes, 0)) +
    ' bytes)',
);
