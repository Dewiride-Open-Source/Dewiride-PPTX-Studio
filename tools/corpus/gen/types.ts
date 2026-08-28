import type { ProbePackage } from './package.ts';

/**
 * One Tier A corpus deck.
 *
 * `features` is a **literal**, not a computation. The temptation is to have the
 * builder tally what it emitted and call that the expectation, which is how
 * `tools/bench/deck.ts` works and is right for a deck whose contents a recipe
 * decides. Here it would be circular: a probe deck exists to state what markup
 * is in it, and a map derived from the markup restates the markup rather than
 * checking it.
 *
 * So the map is written down, reviewed, and locked by `probes.test.ts`, which
 * runs the census - an independent token scanner - over the built bytes and
 * demands exact agreement. Changing the markup fails the test until a person
 * updates the map, which is the point.
 *
 * Authoring a deck is therefore: write the map by hand, run
 * `npx vitest run tools/corpus/gen/probes.test.ts -t <id>`, read the diff,
 * decide whether what moved should have moved, and freeze. There is no
 * `--census` flag on the generator to shortcut it, and there should not be:
 * `build-probes.ts` runs under bare `node`, which cannot resolve
 * `@pptx-studio/census` - nothing links workspace packages into `tools/`, and
 * the resolution the tests get comes from Vitest's alias table. A flag that
 * printed the answer would also make it too easy to paste the map without
 * reading it, which is the one thing this design is trying to prevent.
 */
export interface ProbeDeck {
  /** Kebab-case, matching the manifest entry id and the file name. */
  readonly id: string;
  /** `dc:title`. */
  readonly title: string;
  /** The manifest's `description`: what this deck is the probe for. */
  readonly description: string;
  /** Census keys to exact counts. Keys absent from the map must count zero. */
  readonly features: Readonly<Record<string, number>>;
  /**
   * The file extension, without the dot. Defaults to `pptx`.
   *
   * There is exactly one reason for this to be anything else: PowerPoint
   * checks a package's file extension against the content type of its main
   * part and refuses the pair when they disagree, so a deck typed
   * `ms-powerpoint.presentation.macroEnabled.main+xml` has to be written
   * `.pptm`. `a32-macros` is that deck and is expected to stay the only one -
   * `.ppsx`, `.potx` and the rest differ from `.pptx` in the same one content
   * type and would be probes for the same fact.
   *
   * It also reaches `C016-gitattributes`, which wants a `binary` line for every
   * committed extension, and the manifest's `format` field.
   */
  readonly extension?: string;
  readonly build: () => ProbePackage;
}

/** `<id>.<extension>`, the only name a probe deck is ever written under. */
export function outputName(deck: ProbeDeck): string {
  return deck.id + '.' + (deck.extension ?? 'pptx');
}
