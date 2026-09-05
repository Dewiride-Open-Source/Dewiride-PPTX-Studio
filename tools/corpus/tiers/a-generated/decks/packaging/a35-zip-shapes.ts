import { png } from '../../assets/png.ts';
import {
  grid,
  picture,
  prstGeom,
  scheme,
  shape,
  solidFill,
  type Cell,
} from '../../markup/shapes.ts';
import { REL } from '../../markup/chassis.ts';
import { textLine, txBody } from '../../markup/text.ts';
import type { ProbeDeck } from '../../markup/types.ts';

/**
 * The archive, rather than the markup: mixed compression methods, the
 * general-purpose flag bits, and the growth hint.
 *
 * ## The one deck that deflates, and why the other forty-one do not
 *
 * ADR 0009 committed the bytes of every corpus deck rather than the recipes,
 * because a recipe's hash is the hash of a *build* - `node:zlib` and `fflate`
 * both move, and a routine upgrade would invalidate every pinned hash with no
 * source change. Storing every entry removes that: a stored archive's bytes are
 * a pure function of the XML above them, so `C-REGEN` is stable for the life of
 * the project.
 *
 * What that costs is real: no Tier A deck then exercises DEFLATE at all, and
 * the ZIP reader's decompression path - budgets, ratios, the running inflated
 * counter from sub-phase 0.2 - never sees a corpus fixture. This deck is the
 * exception, and making the exception the deck that is *about* compression
 * keeps the dependency somewhere it can be reasoned about. **`a35` is the one
 * deck whose committed hash can change on a toolchain upgrade with no source
 * change**, and the fix when that happens is to check that nothing but the
 * deflate streams moved and re-pin it.
 *
 * ## What PowerPoint writes, measured
 *
 * From `corpus/ground-truth/powerpoint-conventions.json`, experiment E8, build
 * 16.0.20326:
 *
 * ```
 * generalPurposeFlag      0x0006
 * utf8FlagSet             false
 * compressionMethod       8
 * dosTime                 0x0000
 * dosDate                 0x0021
 * firstExtraFieldId       0xa220
 * firstExtraFieldLength   520
 * hasDirectoryEntries     false
 * ```
 *
 * Two of those are worth spelling out.
 *
 * **`0x0006` is bits 1 and 2, and they mean nothing.** For method 8 they are
 * the deflate level hint - `00` normal, `01` maximum, `10` fast, `11` super
 * fast - and no decompressor consults them, because the stream says what it is.
 * PowerPoint sets `11` on every entry. A reader that validated them, or a
 * writer that copied them onto a **stored** entry where they have no meaning at
 * all, would both be wrong in ways nothing would notice; this deck does the
 * second on purpose.
 *
 * **`0xa220` is Microsoft's growth hint**, and it is padding. Header id, data
 * length, the signature `0xa028`, a padding value, then that many zero bytes,
 * sitting in the **local** header between the file name and the payload so a
 * part can be rewritten slightly larger in place without moving every entry
 * after it. PowerPoint writes 520 bytes of it on the first entry. It is an
 * extra field, so a conformant reader skips it by its declared length and never
 * looks inside - which is exactly why it is worth having in the corpus: it is
 * several hundred bytes between the local header and the data that a reader
 * computing the payload offset from the name length alone will walk straight
 * into.
 *
 * ## What this archive does
 *
 * | entry                      | method   | flags    | extra |
 * | -------------------------- | -------- | -------- | ----- |
 * | `[Content_Types].xml`      | deflate  | `0x0006` | 520   |
 * | `_rels/.rels`              | deflate  | `0x0006` | -     |
 * | `ppt/presentation.xml`     | deflate  | -        | 64    |
 * | `ppt/slides/slide1.xml`    | **store**| -        | -     |
 * | `ppt/slides/slide2.xml`    | **store**| `0x0006` | -     |
 * | `ppt/slides/slide3.xml`    | deflate  | `0x0800` | -     |
 * | `ppt/theme/theme1.xml`     | deflate  | `0x0002` | -     |
 * | `ppt/tableStyles.xml`      | deflate  | `0x0004` | -     |
 * | `docProps/app.xml`         | **store**| -        | 128   |
 * | `ppt/media/image2.png`     | **store**| -        | -     |
 * | everything else            | deflate  | -        | -     |
 *
 * `0x0800` is bit 11, the language-encoding flag, and it is the only flag here
 * with consequences: it declares the entry **name** to be UTF-8 rather than
 * CP437. The name it is set on is pure ASCII, where the two encodings agree, so
 * the flag is true and redundant at once - which is the case a reader has to
 * get right before it can get a non-ASCII name right, and the case PowerPoint
 * never produces because it never sets the bit.
 *
 * `ppt/media/image2.png` is stored for a different reason: it is **noise**, so
 * DEFLATE makes it bigger. The writer notices and stores it anyway. That branch
 * exists in every ZIP writer and is exercised by no other deck.
 *
 * Bit 3 - sizes in a trailing data descriptor - is deliberately absent. This
 * writer does not emit one, so setting the bit would produce an archive that
 * lies about itself, and an archive that lies is `corpus/reject/`'s business.
 */

/** A 24x24 image of pure noise, from a fixed sequence. Deflate makes it larger. */
function noisePng(): Uint8Array {
  // A linear congruential generator, so the bytes are the same on every
  // machine and in every run. `Math.random()` here would break `C-REGEN` in a
  // way nothing else would notice.
  let state = 0x2545f491;
  const next = (): number => {
    state = (Math.imul(state, 1664525) + 1013904223) >>> 0;
    return (state >>> 8) & 0xffffff;
  };
  return png(24, 24, () => next());
}

/** A compressible image: flat colour, so deflate wins by a wide margin. */
function flatPng(): Uint8Array {
  return png(24, 24, () => 0x2980b9);
}

const cell = grid(2, 2);

function box(id: number, name: string, c: Cell, accent: string, lines: readonly string[]): string {
  return shape({
    id,
    name,
    x: c.x,
    y: c.y,
    cx: c.cx,
    cy: c.cy,
    geometry: prstGeom('roundRect'),
    fill: solidFill(scheme(accent, '<a:lumMod val="60000"/><a:lumOff val="40000"/>')),
    textBody: txBody({
      bodyPr: '<a:bodyPr wrap="square" anchor="ctr"/>',
      paras: lines
        .map((line, i) => textLine(line, { sz: 1000, b: i === 0 }, { algn: 'ctr' }))
        .join(''),
    }),
  });
}

/** PowerPoint's own value, on every entry it writes. Bits 1 and 2. */
const DEFLATE_HINT = 0x0006;
/** Bit 11: the entry name is UTF-8 rather than CP437. */
const UTF8_NAME = 0x0800;

export const a35ZipShapes: ProbeDeck = {
  id: 'a35-zip-shapes',
  title: 'PPTX Studio corpus: a35 zip shapes',
  description:
    'The only corpus deck that deflates, and therefore the only one whose committed hash depends on ' +
    'which zlib built it - deliberately the deck that is about compression, so the dependency sits ' +
    'where it can be reasoned about. Mixed methods: four stored entries among deflated ones, one of ' +
    'them stored because it is noise and DEFLATE made it larger. The general-purpose flag bits ' +
    'PowerPoint writes (0x0006, the meaningless deflate-level hint, including on a stored entry ' +
    'where it means even less), bit 1 and bit 2 alone, and bit 11 - the UTF-8 name flag PowerPoint ' +
    'never sets - on an ASCII name where it is true and redundant at once. And Microsoft&apos;s ' +
    '0xa220 growth hint at 520 bytes, the length measured from a PowerPoint package, plus 64 and 128 ' +
    'on two more entries, because it is padding in the local header that a reader computing the ' +
    'payload offset from the name length alone walks straight into.',
  features: {
    shape: 18,
    placeholder: 6,
    presetGeom: 14,
    gradientFill: 2,
    picture: 2,
  },
  build: () => ({
    title: 'PPTX Studio corpus: a35 zip shapes',
    deflate: true,
    parts: [
      {
        name: 'ppt/media/image1.png',
        bytes: flatPng(),
        contentType: { kind: 'default', extension: 'png', type: 'image/png' },
      },
      { name: 'ppt/media/image2.png', bytes: noisePng() },
    ],
    zip: {
      // PowerPoint's own shape, on the entry PowerPoint puts it on.
      '[Content_Types].xml': { flags: DEFLATE_HINT, growthHint: 520 },
      '_rels/.rels': { flags: DEFLATE_HINT },
      'ppt/presentation.xml': { growthHint: 64 },
      'ppt/slides/slide1.xml': { store: true },
      // Flags that mean nothing for method 0, on an entry that uses method 0.
      'ppt/slides/slide2.xml': { store: true, flags: DEFLATE_HINT },
      'ppt/slides/slide3.xml': { flags: UTF8_NAME },
      'ppt/theme/theme1.xml': { flags: 0x0002 },
      'ppt/tableStyles.xml': { flags: 0x0004 },
      'docProps/app.xml': { store: true, growthHint: 128 },
      // Noise. The writer tries deflate, finds it bigger, and stores it.
      'ppt/media/image2.png': { store: true },
    },
    slides: [
      {
        title: 'a35 — two methods in one archive',
        rels: [
          { id: 'rId2', type: REL + 'image', target: '../media/image1.png' },
          { id: 'rId3', type: REL + 'image', target: '../media/image2.png' },
        ],
        body:
          picture({
            id: 10,
            name: 'A compressible image',
            relId: 'rId2',
            x: cell(0).x,
            y: cell(0).y,
            cx: 914400,
            cy: 914400,
            description: 'Flat colour: DEFLATE wins, so this part is method 8',
          }) +
          box(
            11,
            'This slide is stored',
            { ...cell(0), x: cell(0).x + 1005840, cx: cell(0).cx - 1005840 },
            'accent1',
            ['ppt/slides/slide1.xml', 'method 0, in an archive where', 'most entries are method 8'],
          ) +
          picture({
            id: 12,
            name: 'An incompressible image',
            relId: 'rId3',
            x: cell(1).x,
            y: cell(1).y,
            cx: 914400,
            cy: 914400,
            description: 'Noise: DEFLATE makes it larger, so this part is method 0',
          }) +
          box(
            15,
            'Noise is stored too',
            { ...cell(1), x: cell(1).x + 1005840, cx: cell(1).cx - 1005840 },
            'accent2',
            [
              'ppt/media/image2.png',
              'DEFLATE makes it larger, so the writer',
              'stores it. Every ZIP writer has this branch.',
            ],
          ) +
          box(13, 'And most things deflate', cell(2), 'accent3', [
            'The masters, the layouts, the theme.',
            'This is the only deck in the corpus',
            'where any of that is true.',
          ]) +
          box(14, 'What it costs', cell(3), 'accent4', [
            'C-REGEN for this deck depends on zlib.',
            'For the other forty-one it does not,',
            'which is why they all store.',
          ]),
      },
      {
        title: 'a35 — the general-purpose flag bits',
        body:
          box(10, 'What PowerPoint writes', cell(0), 'accent5', [
            '0x0006 — bits 1 and 2',
            'the deflate level hint: 11, super fast.',
            'No decompressor reads them.',
          ]) +
          box(11, 'On a stored entry', cell(1), 'accent6', [
            'slide2.xml is method 0 and carries 0x0006.',
            'The bits mean nothing for method 8',
            'and less than nothing for method 0.',
          ]) +
          box(12, 'Bit 11, which does mean something', cell(2), 'accent1', [
            '0x0800 on slide3.xml',
            'the entry name is UTF-8, not CP437.',
            'PowerPoint never sets it.',
          ]) +
          box(13, 'True and redundant', cell(3), 'accent2', [
            'The name it is set on is pure ASCII,',
            'where the two encodings agree —',
            'the case to get right first.',
          ]),
      },
      {
        title: 'a35 — the 0xa220 growth hint',
        body:
          box(10, 'What it is', cell(0), 'accent3', [
            'header id 0xa220, a length,',
            'signature 0xa028, a padding value,',
            'then that many zero bytes',
          ]) +
          box(11, 'Where it is', cell(1), 'accent4', [
            'The local header only, between the',
            'file name and the payload. The central',
            'directory entry has no extra field.',
          ]) +
          box(12, 'Why it is there', cell(2), 'accent5', [
            'So a part can be rewritten slightly larger',
            'in place, without moving every entry',
            'after it in the archive.',
          ]) +
          box(13, 'Three lengths', cell(3), 'accent6', [
            '520 on [Content_Types].xml, as measured;',
            '64 on presentation.xml; 128 on app.xml.',
            'The length is per entry, not a constant.',
          ]),
      },
    ],
  }),
};
