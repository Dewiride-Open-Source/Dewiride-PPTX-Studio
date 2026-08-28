import { REL } from '../package.ts';
import { png } from '../png.ts';
import { grid, picture, prstGeom, scheme, shape, solidFill, type Cell } from '../shapes.ts';
import { textLine, txBody } from '../text.ts';
import type { ProbeDeck } from '../types.ts';

/**
 * Part names and text at the edges of what OPC and XML 1.0 permit.
 *
 * ## Part names: what is legal, and what PowerPoint actually accepts
 *
 * An OPC part name is a sequence of `/`-separated segments of `pchar`, and
 * `pchar` is more generous than anybody writes: the unreserved set, the
 * sub-delims `! $ & ' ( ) * + , ; =`, plus `:` and `@`, plus any octet as a
 * percent-escape. So `image+2,~!()$@.png` is a perfectly ordinary part name
 * that no producer has ever emitted.
 *
 * ADR 0003 measured which of those PowerPoint accepts, and the answer that
 * matters is the pair:
 *
 * - **`/ppt/media/image%201.png` opens.** A percent-escaped space is a part
 *   name PowerPoint reads and resolves relationships against.
 * - **`/ppt/media/image 1.png`, the literal space, is refused** with
 *   `0x808D1001` - a different error from the usual repair, from the packaging
 *   layer rather than the presentation one.
 *
 * ## PowerPoint applies RFC 3986 normalisation, and it is fatal
 *
 * The first version of this deck named a part `ppt/media/im%61ge-4.png` -
 * `%61` is `a` - and the whole package was refused with `0x808D1005`, a code
 * this project had not seen before and one more from the packaging layer. Nine
 * one-name packages later, on 2026-08-28 against 16.0.20326:
 *
 * | escape | character   | class                | result   |
 * | ------ | ----------- | -------------------- | -------- |
 * | `%2D`  | `-`         | unreserved           | refused  |
 * | `%41`  | `A`         | unreserved           | refused  |
 * | `%5F`  | `_`         | unreserved           | refused  |
 * | `%7E`  | `~`         | unreserved           | refused  |
 * | `%24`  | `$`         | sub-delim            | **opens**|
 * | `%2C`  | `,`         | sub-delim            | **opens**|
 * | `%3A`  | `:`         | gen-delim, and pchar | **opens**|
 * | `%20`  | space       | not pchar at all     | **opens**|
 * | `%23`  | `#`         | not pchar at all     | **opens**|
 *
 * Nine for nine, and the line is exactly RFC 3986 §6.2.2.2: *the percent-encoded
 * octet of an unreserved character must be normalised to that character.* An
 * escape that is **required** is fine; an escape that is **gratuitous** is a
 * refusal. So `%20` opens and `%41` does not, which looks arbitrary until the
 * rule is named.
 *
 * **This project rates that case a warning.** `packages/opc`'s grammar has it -
 * rule `M1.8`, over-encoded unreserved - at `severity: 'warning'`, so
 * `toPartName` accepts it and only `isValidPartName` objects. The one consumer
 * that matters treats it as fatal. That is a severity to fix in sub-phase 1.2
 * rather than a rule to add, and it is the sort of thing only a measurement
 * finds: nothing in the code was wrong, and the number beside it was.
 *
 * Note what the finding does **not** licence. This repository still never
 * decodes a part name: `%2e%2e` is not `..`, and two names that differ only in
 * escaping are two parts. Decoding first, then comparing, is a path-traversal
 * bug wearing a normalisation hat. Rejecting an over-encoding and decoding one
 * are different operations, and only the first is safe.
 *
 * The four names on slide 1 are all accepted by `packages/opc`'s grammar with
 * **zero** violations, which is deliberate. The grammar's warning tier - a
 * literal space, a `#`, any non-ASCII character - is escalated to fatal by
 * `PartStore.write()`, so a fixture using one could be read by this project and
 * never written by it. A corpus deck our own writer cannot re-emit would fail
 * sub-phase 1.4 for a reason that has nothing to do with 1.4.
 *
 * ## Text: XML 1.0 is narrower than Unicode, and wider than UTF-16
 *
 * `Char ::= #x9 | #xA | #xD | [#x20-#xD7FF] | [#xE000-#xFFFD] |
 * [#x10000-#x10FFFF]`. Three facts fall out of that and all three are on
 * slide 2:
 *
 * **Astral characters are one character and two code units.** `𝕏` is U+1D54F,
 * a surrogate pair in UTF-16 and one character in XML. Anything that measures
 * text in `String.length`, breaks lines by index, or truncates to a byte budget
 * splits it and produces a lone surrogate - which is not a legal XML character
 * at all, so the file it writes cannot be parsed. The census counts UTF-16 code
 * units and says so; a renderer must not.
 *
 * **A literal carriage return is not a carriage return.** XML 1.0 §2.11 makes
 * every parser normalise a literal `#xD` - and `#xD#xA` - to a single `#xA`
 * *before* the application sees it. A numeric character reference `&#13;` is
 * **not** normalised. So the two spellings of the same character mean different
 * things, and a serializer that re-emits parsed text as literal CR silently
 * changes the document on every round trip. This is the one place where "keep
 * the bytes" and "keep the characters" are not the same instruction, and it is
 * why sub-phase 0.5 re-emits clean nodes by slicing the original buffer rather
 * than by re-escaping a string.
 *
 * **Combining marks and bidi controls are ordinary text.** U+0301 after a
 * letter is one grapheme in two characters; U+202E is invisible and reverses
 * everything after it. Neither is a special case in the file format and both
 * change what a line-breaker and a caret have to do.
 *
 * What is **not** here: `#x0`-`#x8`, `#xB`, `#xC`, `#xE`-`#x1F`, `#xFFFE` and
 * `#xFFFF`, none of which XML 1.0 permits at all, in any spelling - a numeric
 * character reference to a non-`Char` is as malformed as the literal. Those
 * belong in `corpus/reject/`, not in a deck that has to open.
 *
 * ## Attribute values
 *
 * Slide 3 puts the same content in `@name` and `@descr`, because an attribute
 * is escaped differently from character data - `&quot;` matters there and
 * nowhere else - and because `@descr` is the one attribute in a deck that holds
 * text a human typed. `a19-decorative` established that PowerPoint writes
 * `&quot;` and nothing else; this checks the rest of the range survives it.
 *
 * ## What survives a round trip, measured
 *
 * Re-saved through PowerPoint on 2026-08-28 and read back:
 *
 * **The characters survive.** `𝕏📄𠮷` comes back as U+1D54F, U+1F4C4, U+20BB7 -
 * three astral characters, six UTF-16 code units, written as literal UTF-8
 * rather than as references. The combining acute, U+202E and U+202C, U+FFFD,
 * ZWJ, ZWNJ and NBSP all come back unchanged. Nothing in this range is a
 * special case for PowerPoint, which is the useful answer: it means a renderer
 * has no licence to treat them as one either.
 *
 * **The carriage return does not, and it is not lost - it is promoted.** One
 * run written `before&#13;after` comes back as **two `a:p` elements**. The
 * numeric character reference survives XML 1.0's line-ending normalisation, as
 * §2.11 says it must, and then does not survive PowerPoint's text model, which
 * reads `#xD` inside an `a:t` as a paragraph break. So a control character in
 * run text is **structure**, not text, and a text engine that stored it as a
 * character would disagree with PowerPoint about how many paragraphs a shape
 * has.
 *
 * **The part names do not survive, and neither would ordinary ones.**
 * `image%201.png`, `image+2,~!()$@'.png`, `IMAGE-3.PNG` and `image%234.png`
 * come back as `image1.png` to `image4.png` - PowerPoint renumbers media on
 * every save whatever they were called, which `a25-svg-blips` had already
 * shown. The one detail worth keeping: the mixed-case name comes back
 * `image3.PNG`, so the **extension's** case is preserved while the stem's is
 * not.
 */

/** A percent-escaped space. Measured accepted; the literal space is refused. */
const NAME_ESCAPED_SPACE = 'ppt/media/image%201.png';

/** Sub-delims, `~`, `(`, `)`, `$`, `@` and `+` - all pchar, none ever written. */
const NAME_PUNCTUATION = "ppt/media/image+2,~!()$@'.png";

/** Mixed case. Part names are compared case-insensitively for collisions and stored as written. */
const NAME_MIXED_CASE = 'ppt/media/IMAGE-3.PNG';

/**
 * An escape that is *required*: `%23` is `#`, which is not pchar at all.
 *
 * The counterpart to the table above. This one opens; `im%61ge-4.png`, whose
 * only difference is that `a` did not need escaping, does not.
 */
const NAME_REQUIRED_ESCAPE = 'ppt/media/image%234.png';

const SWATCHES = [0xc0392b, 0x27ae60, 0x2980b9, 0xf1c40f] as const;

/** A flat colour with a corner mark, so a mis-resolved relationship is visible. */
function marked(colour: number): Uint8Array {
  return png(24, 24, (x, y) => (x < 8 && y < 8 ? 0xffffff : colour));
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

/**
 * A shape whose `a:t` content is written **raw**, bypassing `escapeXml`.
 *
 * Every other deck goes through the chassis's escaper, which is right: it is
 * what stops a stray `<` becoming a whole-package refusal. This deck is about
 * what the escaper produces and what it must not touch, so it writes the
 * markup itself and takes responsibility for it.
 */
function rawText(
  id: number,
  name: string,
  c: Cell,
  accent: string,
  label: string,
  raw: string,
): string {
  return shape({
    id,
    name,
    x: c.x,
    y: c.y,
    cx: c.cx,
    cy: c.cy,
    geometry: prstGeom('roundRect'),
    fill: solidFill(scheme(accent, '<a:lumMod val="60000"/><a:lumOff val="40000"/>')),
    textBody:
      '<p:txBody><a:bodyPr wrap="square" anchor="ctr"/><a:lstStyle/>' +
      '<a:p><a:pPr algn="ctr"/><a:r><a:rPr lang="en-GB" sz="1000" b="1" dirty="0"/>' +
      `<a:t>${label}</a:t></a:r></a:p>` +
      '<a:p><a:pPr algn="ctr"/><a:r><a:rPr lang="en-GB" sz="1600" dirty="0"/>' +
      `<a:t>${raw}</a:t></a:r></a:p>` +
      '</p:txBody>',
  });
}

export const a40Unicode: ProbeDeck = {
  id: 'a40-unicode',
  title: 'PPTX Studio corpus: a40 unicode',
  description:
    'Part names and text at the edges. Four media parts named with a percent-escaped space, the ' +
    'pchar punctuation no producer emits, mixed case, and an escape that is required rather than ' +
    'gratuitous. A fifth was built and refused, and bisecting it produced the finding: PowerPoint ' +
    'applies RFC 3986 normalisation and rejects a percent-escape of an unreserved character outright ' +
    'with 0x808D1005, nine probes for nine, while an escape of anything else opens - so %20 is fine ' +
    'and %41 is fatal. This project rates that case a warning. Text with ' +
    'astral characters that are one character and two UTF-16 code units, combining marks, bidi ' +
    'controls, U+FFFD, and the numeric character reference &#13; beside a literal CR - which XML ' +
    '1.0 normalises and the reference does not, so the two spellings of one character mean ' +
    'different things and a re-escaping serializer changes the document every round trip.',
  features: {
    shape: 18,
    placeholder: 6,
    presetGeom: 16,
    gradientFill: 2,
    picture: 4,
  },
  build: () => ({
    title: 'PPTX Studio corpus: a40 unicode',
    parts: [
      {
        name: NAME_ESCAPED_SPACE,
        bytes: marked(SWATCHES[0]),
        contentType: { kind: 'default', extension: 'png', type: 'image/png' },
      },
      { name: NAME_PUNCTUATION, bytes: marked(SWATCHES[1]) },
      { name: NAME_MIXED_CASE, bytes: marked(SWATCHES[2]) },
      { name: NAME_REQUIRED_ESCAPE, bytes: marked(SWATCHES[3]) },
    ],
    slides: [
      {
        title: 'a40 — four part names no producer writes',
        // Targets are relative to `ppt/slides/`, and are spelled exactly as the
        // ZIP entry is. Nothing decodes them, so `%61` here has to be `%61`
        // there; writing `a` would dangle.
        rels: [
          { id: 'rId2', type: REL + 'image', target: '../media/image%201.png' },
          { id: 'rId3', type: REL + 'image', target: "../media/image+2,~!()$@'.png" },
          { id: 'rId4', type: REL + 'image', target: '../media/IMAGE-3.PNG' },
          { id: 'rId5', type: REL + 'image', target: '../media/image%234.png' },
        ],
        body:
          picture({
            id: 10,
            name: 'Escaped space',
            relId: 'rId2',
            x: cell(0).x,
            y: cell(0).y,
            cx: 914400,
            cy: 914400,
            description: 'ppt/media/image%201.png — a percent-escaped space, measured accepted',
          }) +
          box(
            11,
            'Escaped space',
            { ...cell(0), x: cell(0).x + 1005840, cx: cell(0).cx - 1005840 },
            'accent1',
            ['image%201.png', 'opens. The literal space', 'is refused with 0x808D1001.'],
          ) +
          picture({
            id: 12,
            name: 'pchar punctuation',
            relId: 'rId3',
            x: cell(1).x,
            y: cell(1).y,
            cx: 914400,
            cy: 914400,
            description: "ppt/media/image+2,~!()$@'.png — sub-delims, all of them pchar",
          }) +
          box(
            13,
            'Punctuation',
            { ...cell(1), x: cell(1).x + 1005840, cx: cell(1).cx - 1005840 },
            'accent2',
            ["image+2,~!()$@'.png", 'Every one of these is pchar,', 'and none is ever written.'],
          ) +
          picture({
            id: 14,
            name: 'Mixed case',
            relId: 'rId4',
            x: cell(2).x,
            y: cell(2).y,
            cx: 914400,
            cy: 914400,
            description: 'ppt/media/IMAGE-3.PNG — stored as written, compared case-insensitively',
          }) +
          box(
            15,
            'Mixed case',
            { ...cell(2), x: cell(2).x + 1005840, cx: cell(2).cx - 1005840 },
            'accent3',
            ['IMAGE-3.PNG', 'The Default keys on "png" and', 'the extension here is "PNG".'],
          ) +
          picture({
            id: 16,
            name: 'Over-encoded',
            relId: 'rId5',
            x: cell(3).x,
            y: cell(3).y,
            cx: 914400,
            cy: 914400,
            description: 'ppt/media/image%234.png — %23 is "#", an escape that is required',
          }) +
          box(
            17,
            'A required escape',
            { ...cell(3), x: cell(3).x + 1005840, cx: cell(3).cx - 1005840 },
            'accent4',
            [
              'image%234.png',
              '%23 is "#", not pchar, so it must be',
              'escaped. im%61ge-4.png is refused.',
            ],
          ),
      },
      {
        title: 'a40 — text at the edges of XML 1.0',
        body:
          rawText(
            10,
            'Astral characters',
            cell(0),
            'accent5',
            'astral: one character, two code units',
            '&#x1D54F;&#x1F4C4;&#x20BB7;',
          ) +
          rawText(
            11,
            'Combining marks and bidi',
            cell(1),
            'accent6',
            'combining mark, then U+202E',
            'e&#x301;cole &#x202E;reversed&#x202C; back',
          ) +
          rawText(
            12,
            'Carriage returns, two spellings',
            cell(2),
            'accent1',
            'a literal CR is normalised; &amp;#13; is not',
            'before&#13;after',
          ) +
          rawText(
            13,
            'Replacement and zero width',
            cell(3),
            'accent2',
            'U+FFFD, ZWJ, ZWNJ, NBSP',
            'a&#xFFFD;b&#x200D;c&#x200C;d&#xA0;e',
          ),
      },
      {
        title: 'a40 — the same content in an attribute',
        body:
          shape({
            id: 10,
            name: 'quote " ampersand & angle < >',
            descr: 'quote " ampersand & angle < > — every one of these needs escaping in @descr',
            x: cell(0).x,
            y: cell(0).y,
            cx: cell(0).cx,
            cy: cell(0).cy,
            geometry: prstGeom('roundRect'),
            fill: solidFill(scheme('accent3', '<a:lumMod val="60000"/><a:lumOff val="40000"/>')),
            textBody: txBody({
              bodyPr: '<a:bodyPr wrap="square" anchor="ctr"/>',
              paras:
                textLine('@name and @descr', { sz: 1000, b: true }, { algn: 'ctr' }) +
                textLine('quote " ampersand & angle < >', { sz: 1000 }, { algn: 'ctr' }) +
                textLine('&quot; is the fourth entity', { sz: 1000 }, { algn: 'ctr' }),
            }),
          }) +
          shape({
            id: 11,
            name: 'Astral in a name: \u{1D54F}\u{1F4C4}',
            descr: 'Astral characters in an attribute: \u{1D54F} \u{1F4C4} \u{20BB7}',
            x: cell(1).x,
            y: cell(1).y,
            cx: cell(1).cx,
            cy: cell(1).cy,
            geometry: prstGeom('roundRect'),
            fill: solidFill(scheme('accent4', '<a:lumMod val="60000"/><a:lumOff val="40000"/>')),
            textBody: txBody({
              bodyPr: '<a:bodyPr wrap="square" anchor="ctr"/>',
              paras:
                textLine('astral in @name and @descr', { sz: 1000, b: true }, { algn: 'ctr' }) +
                textLine('written as UTF-8, not as a reference', { sz: 1000 }, { algn: 'ctr' }),
            }),
          }) +
          shape({
            id: 12,
            name: 'Bidi in a name: ‮reversed‬',
            descr: 'A right-to-left override inside an attribute value: ‮reversed‬ back',
            x: cell(2).x,
            y: cell(2).y,
            cx: cell(2).cx,
            cy: cell(2).cy,
            geometry: prstGeom('roundRect'),
            fill: solidFill(scheme('accent5', '<a:lumMod val="60000"/><a:lumOff val="40000"/>')),
            textBody: txBody({
              bodyPr: '<a:bodyPr wrap="square" anchor="ctr"/>',
              paras:
                textLine('U+202E in @name', { sz: 1000, b: true }, { algn: 'ctr' }) +
                textLine('invisible, and it reverses the pane', { sz: 1000 }, { algn: 'ctr' }),
            }),
          }) +
          box(13, 'What is deliberately absent', cell(3), 'accent6', [
            'no #x0-#x8, #xB, #xC, #xE-#x1F',
            'no #xFFFE or #xFFFF, in any spelling:',
            'XML 1.0 forbids them, so corpus/reject/',
          ]),
      },
    ],
  }),
};
