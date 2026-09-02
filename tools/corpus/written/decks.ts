/**
 * Tier C's roster, as reviewed literals.
 *
 * One deck, and the shape still mirrors `gen/decks/*.ts` and `authored/decks.ts`
 * so that all three tiers state their census the same way and
 * `decks.test.ts` can assert exact equality in both directions.
 *
 * What Tier C is for
 * ------------------
 * Tier A proves our generator can emit a package PowerPoint accepts. Tier B
 * proves we can read one PowerPoint wrote. Neither exercises the writer this
 * project actually ships, which is the one that has to hand a user back their
 * own file at the end of an edit - so `c01-opc-writer` is `b01-blank` read
 * through `PartStore.open` and written straight back out by `PartStore.write`,
 * with nothing touched in between.
 *
 * That is a smaller claim than it sounds and a more load-bearing one. The
 * architecture's first bet is that a part nobody edited is re-emitted
 * byte-for-byte, and this is the committed evidence for it at package scale:
 * thirty-seven entries in, thirty-seven out, in the same order, every part
 * byte-identical after inflation, and no entry recompressed.
 *
 * Its census contributes nothing
 * ------------------------------
 * Deliberately. `c01`'s features map is `b01`'s, key for key, because the parts
 * are the same bytes - so under `C-COV`, a rule about feature coverage, this
 * deck is pure redundancy. Everything it does contribute is at the layer
 * `C-COV` cannot see: the ZIP headers, which are ours and not Microsoft's.
 * Recording that here rather than letting a reader work it out is the point of
 * having the literal at all.
 */

export interface WrittenDeck {
  readonly id: string;
  /**
   * The committed deck this one is written from, workspace-relative.
   *
   * A Tier C deck is a function of another corpus entry rather than of a
   * markup module, which is what makes it different from both other tiers and
   * why the source is named here rather than being implied by the recipe.
   */
  readonly source: string;
  /**
   * SHA-256 of that source at the time this deck was built.
   *
   * Pinned because the dependency is otherwise invisible: re-authoring Tier B
   * changes `b01-blank.pptx`, and without this the only symptom would be
   * `C-REGEN` failing on `c01` with nothing to say about why. With it, the
   * build refuses before it writes and names the deck that moved.
   */
  readonly sourceSha256: string;
  readonly slides: number;
  readonly description: string;
  readonly features: Readonly<Record<string, number>>;
}

export const WRITTEN_DECKS: readonly WrittenDeck[] = [
  {
    id: 'c01-opc-writer',
    source: 'corpus/authored/b01-blank.pptx',
    sourceSha256: '2a173ed55d222e863779e0b73178a5759053770b8d92ab481c6578ce16a9ccd1',
    slides: 1,
    description:
      'b01-blank read through PartStore.open and written straight back out, with nothing touched ' +
      'in between - the committed evidence that a part nobody edited is re-emitted byte-for-byte. ' +
      'All 37 entries survive in their original order, every part is byte-identical after ' +
      'inflation, and no entry is recompressed, because an untouched part is passed through as the ' +
      "DEFLATE stream it already was. Its census is b01's exactly and contributes nothing to C-COV; " +
      'what it contributes is a third producer of ZIP headers, and the diff against its source is ' +
      'the whole of what our writer normalises away: version-made-by 45 becomes 20, the ' +
      'general-purpose flags 0x0006 that PowerPoint sets on every deflated entry become 0, and the ' +
      'five 0xA220 growth-hint extra fields - 520 bytes on each of the first two entries and 264 on ' +
      'three more, 1832 bytes in all - are dropped, which is the entire size difference. Writing it ' +
      'again is byte-identical, and writing the output again is a fixpoint.',
    features: {
      placeholder: 65,
      shape: 65,
      field: 24,
      presetGeom: 5,
      gradientFill: 3,
      shadow: 1,
      thumbnail: 1,
    },
  },
];
