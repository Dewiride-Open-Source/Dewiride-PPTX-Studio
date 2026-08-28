import { REL } from '../package.ts';
import { probeJpeg } from '../jpeg.ts';
import { grid, picture, prstGeom, scheme, shape, solidFill, type Cell } from '../shapes.ts';
import { textLine, txBody } from '../text.ts';
import type { ProbeDeck } from '../types.ts';

/**
 * `docProps/thumbnail.jpeg` - the one part of a package no part points at.
 *
 * ## What PowerPoint writes, measured
 *
 * Saving a one-slide deck through COM on 2026-08-28, build 16.0.20326 wrote a
 * 1203-byte `docProps/thumbnail.jpeg`, typed by a **`Default`** rather than an
 * `Override`:
 *
 * ```xml
 * <Default Extension="jpeg" ContentType="image/jpeg"/>
 * ```
 *
 * and reached by one relationship in `_rels/.rels`:
 *
 * ```xml
 * <Relationship Id="rId2"
 *   Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/thumbnail"
 *   Target="docProps/thumbnail.jpeg"/>
 * ```
 *
 * Two things about that measurement are worth more than the part itself.
 *
 * **The relationship is on the package, not on a part.** `_rels/.rels` has four
 * children - `officeDocument`, `core-properties`, `extended-properties` and
 * this - and nothing under `ppt/` mentions the thumbnail at all. It is
 * therefore the one part where "reachable" and "referenced by the document"
 * come apart, and the case that tells a mark-and-sweep media collector whether
 * it walks package relationships or only part ones. Sub-phase 1.3's collector
 * that started from `ppt/presentation.xml` would delete it.
 *
 * **PowerPoint writes `_rels/.rels` out of rId order**: the four came back as
 * `rId3`, `rId2`, `rId1`, `rId4`. `CT_Relationships` is a sequence of
 * `Relationship`, so that order is part of the file, and a writer that
 * regenerates the part sorted by id changes bytes it was not asked to change.
 * This deck does **not** reproduce it - the chassis writes its own three first
 * and appends - so it stays a recorded measurement for a Tier B deck to carry.
 * `ROSTER.md` says so beside the other conventions.
 *
 * ## The census keys on the name, and that is a choice with an edge
 *
 * `packages/census`'s `thumbnail` rule matches the part name lowercased against
 * exactly `/docprops/thumbnail.jpeg`. OPC does not say a thumbnail must be
 * called that, or be a JPEG: §10.2 defines it by **relationship type**, and a
 * `docProps/thumbnail.png` reached by the same type is as legitimate. So the
 * census would report zero thumbnails for a package that has one.
 *
 * That is defensible - the census is a fast token-and-name scanner, and every
 * producer in the world writes that name - but it is a rule with an edge, and
 * the roster records it rather than the corpus quietly agreeing with it. This
 * deck carries the conventional name so the feature is exercised at all; the
 * off-name case belongs in `corpus/reject/` once that exists, because it is a
 * case where the right answer is "the census is wrong", not "the deck is".
 *
 * ## Two JPEGs, one Default
 *
 * The deck also carries `ppt/media/image1.jpeg` on a slide. The point is the
 * content-type table: **one** `<Default Extension="jpeg"/>` covers both parts,
 * and a second one naming the same extension is a package PowerPoint refuses.
 * Nothing distinguishes the thumbnail in that table; it is an ordinary image
 * part that happens to be reached from the package root.
 *
 * Both are written by `tools/corpus/gen/jpeg.ts`, whose Huffman tables were
 * read out of the `DHT` segments of the thumbnail measured above rather than
 * recalled.
 */

const THUMB_WIDTH = 256;
const THUMB_HEIGHT = 144;

/**
 * A blocky greyscale picture of a slide: a dark title band, two content
 * blocks and a footer rule. Recognisable at 256 x 144, which is the whole
 * requirement for a thumbnail.
 */
function slideThumbnail(): Uint8Array {
  return probeJpeg({
    width: THUMB_WIDTH,
    height: THUMB_HEIGHT,
    level: (bx, by) => {
      if (by <= 1) return 45;
      if (by === 2) return 235;
      if (by >= 4 && by <= 8 && bx >= 2 && bx <= 13) return 125;
      if (by >= 4 && by <= 8 && bx >= 18 && bx <= 29) return 175;
      if (by >= 11 && by <= 14 && bx >= 2 && bx <= 29) return 205;
      return 246;
    },
  });
}

/** A second, visibly different image, so the two parts cannot be confused. */
function checkerboard(): Uint8Array {
  return probeJpeg({
    width: 128,
    height: 128,
    level: (bx, by) => ((bx + by) % 2 === 0 ? 60 : 225),
  });
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
        .map((line, i) => textLine(line, { sz: 1100, b: i === 0 }, { algn: 'ctr' }))
        .join(''),
    }),
  });
}

/** OPC Part 2 §10.2. Not under the ECMA-376 relationship base, which is why it is spelled out. */
const THUMBNAIL_REL =
  'http://schemas.openxmlformats.org/package/2006/relationships/metadata/thumbnail';

export const a33Thumbnail: ProbeDeck = {
  id: 'a33-thumbnail',
  title: 'PPTX Studio corpus: a33 thumbnail',
  description:
    'docProps/thumbnail.jpeg, reached by the OPC metadata/thumbnail relationship from _rels/.rels ' +
    'and by nothing under ppt/ - the one part where reachable and referenced-by-the-document come ' +
    'apart, and the case that decides whether a media collector walks package relationships or only ' +
    'part ones. Typed by a Default on the jpeg extension, which a second jpeg part on a slide shares ' +
    'and which a duplicate Default would make PowerPoint refuse. Both images are written by ' +
    'tools/corpus/gen/jpeg.ts, a baseline greyscale encoder whose Huffman tables were read out of ' +
    'the DHT segments of a thumbnail PowerPoint 16.0.20326 wrote rather than recalled.',
  features: {
    shape: 17,
    placeholder: 6,
    presetGeom: 12,
    gradientFill: 2,
    picture: 1,
    thumbnail: 1,
  },
  build: () => ({
    title: 'PPTX Studio corpus: a33 thumbnail',
    parts: [
      {
        name: 'docProps/thumbnail.jpeg',
        bytes: slideThumbnail(),
        contentType: { kind: 'default', extension: 'jpeg', type: 'image/jpeg' },
      },
      // No `contentType` of its own: the Default above already covers it, and a
      // second `<Default Extension="jpeg"/>` is a whole-package refusal.
      { name: 'ppt/media/image1.jpeg', bytes: checkerboard() },
    ],
    packageRels: [{ id: 'rId4', type: THUMBNAIL_REL, target: 'docProps/thumbnail.jpeg' }],
    slides: [
      {
        title: 'a33 — the part nothing under ppt/ points at',
        body:
          box(10, 'Where it lives', cell(0), 'accent1', [
            'docProps/thumbnail.jpeg',
            'a sibling of core.xml and app.xml',
            'not under ppt/ at all',
          ]) +
          box(11, 'How it is reached', cell(1), 'accent2', [
            '_rels/.rels, one relationship',
            '…/package/2006/relationships/metadata/thumbnail',
            'a package relationship, not a part one',
          ]) +
          box(12, 'How it is typed', cell(2), 'accent3', [
            '<Default Extension="jpeg" ContentType="image/jpeg"/>',
            'a Default, not an Override',
            'shared with every other jpeg in the package',
          ]) +
          box(13, 'What deletes it', cell(3), 'accent4', [
            'a media sweep that starts at presentation.xml',
            'and follows only part relationships',
            'never reaches this part',
          ]),
      },
      {
        title: 'a33 — one Default, two jpeg parts',
        rels: [{ id: 'rId2', type: REL + 'image', target: '../media/image1.jpeg' }],
        body:
          picture({
            id: 10,
            name: 'The other jpeg',
            relId: 'rId2',
            x: cell(0).x,
            y: cell(0).y,
            cx: Math.min(cell(0).cx, cell(0).cy),
            cy: Math.min(cell(0).cx, cell(0).cy),
            description: 'A checkerboard, so the two jpeg parts cannot be confused',
          }) +
          box(11, 'ppt/media/image1.jpeg', cell(1), 'accent5', [
            'an ordinary image part',
            'reached from this slide by rId2',
            'and typed by the same one Default',
          ]) +
          box(12, 'The thumbnail is not special', cell(2), 'accent6', [
            'nothing in [Content_Types].xml',
            'distinguishes it from the checkerboard',
            'only the relationship type does',
          ]) +
          box(13, 'A second Default is fatal', cell(3), 'accent1', [
            'two <Default Extension="jpeg"/> entries',
            'is a whole-package refusal',
            'so the table is deduplicated by extension',
          ]),
      },
      {
        title: 'a33 — what the census keys on',
        body:
          box(10, 'The census rule', cell(0), 'accent2', [
            'part name, lowercased, equal to',
            '/docprops/thumbnail.jpeg',
            'a name match, not a relationship match',
          ]) +
          box(11, 'What OPC actually says', cell(1), 'accent3', [
            'Part 2 §10.2 defines a thumbnail',
            'by relationship type alone.',
            'The name and the format are conventions.',
          ]) +
          box(12, 'So the rule has an edge', cell(2), 'accent4', [
            'docProps/thumbnail.png, same relationship,',
            'is a legitimate thumbnail the census',
            'reports as zero.',
          ]) +
          box(13, 'Where that belongs', cell(3), 'accent5', [
            'corpus/reject/, once it exists:',
            'a case where the right answer is',
            '"the census is wrong", not "the deck is".',
          ]),
      },
    ],
  }),
};
