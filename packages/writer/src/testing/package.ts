import { CONTENT_TYPE, PartStore, REL_TYPE } from '@pptx-studio/opc';

/**
 * A package, not a presentation.
 *
 * The writer's unit tests are about the container - what is reachable, what a
 * sweep may remove, whether an untouched entry comes back out unchanged - and
 * none of that needs a deck PowerPoint would open. Building one here anyway
 * would mean a second copy of the minimal-deck knowledge that
 * `packages/validate/src/testing/deck.ts` already holds, and two copies of
 * "what a valid presentation looks like" drift until one of them quietly stops
 * being valid.
 *
 * So these fixtures are the smallest thing `PartStore.write` accepts, the tests
 * here pass `validate: false`, and the firewall is exercised against the real
 * corpus in `tools/corpus/export.test.ts` - where there are fifty-two decks
 * that genuinely open, and a filesystem to read them from.
 */

export interface FixtureOptions {
  /** Media parts to add, each with one relationship from `/doc.xml`. */
  readonly media?: readonly string[];
  /** Media parts to add with **no** relationship pointing at them. */
  readonly orphanMedia?: readonly string[];
}

export const MAIN_PART = '/doc.xml';

/** Distinct, deterministic bytes, so a mix-up shows up as a mismatch. */
export function bytesFor(name: string): Uint8Array {
  const out = new Uint8Array(64);
  for (let i = 0; i < out.length; i++) out[i] = (name.charCodeAt(i % name.length) + i) & 0xff;
  return out;
}

/** The archive bytes of a fixture package. */
export function fixtureBytes(options: FixtureOptions = {}): Uint8Array {
  const store = PartStore.create();
  store.addPart(MAIN_PART, CONTENT_TYPE.xml, new TextEncoder().encode('<doc/>'));
  store.rootRelationships().addTo(REL_TYPE.officeDocument, MAIN_PART);

  for (const part of options.media ?? []) {
    store.addPart(part, CONTENT_TYPE.png, bytesFor(part));
    store.relationships(MAIN_PART).addTo(REL_TYPE.image, part);
  }
  for (const part of options.orphanMedia ?? []) {
    store.addPart(part, CONTENT_TYPE.png, bytesFor(part));
  }
  return store.write();
}

/** A fixture opened twice: the package to edit, and the baseline to judge it against. */
export function fixture(options: FixtureOptions = {}): {
  store: PartStore;
  baseline: PartStore;
  baselineBytes: Uint8Array;
} {
  const bytes = fixtureBytes(options);
  return {
    store: PartStore.open(bytes),
    baseline: PartStore.open(bytes),
    baselineBytes: bytes,
  };
}
