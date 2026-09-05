import { buildProbeFont } from '../../../../../ground-truth/fonts/embedding/build-font.ts';
import { writeEot, EOT_VERSION_2_2 } from '../../../../../ground-truth/fonts/format/eot.ts';
import {
  placeholderXml,
  REL,
  TITLE_BOX,
  type ProbeMaster,
  type ProbePart,
  type ProbeRel,
} from '../../markup/chassis.ts';
import { shape } from '../../markup/shapes.ts';
import { textLine, txBody } from '../../markup/text.ts';
import type { ProbeDeck } from '../../markup/types.ts';

/**
 * Embedded fonts: six artifacts, and five of them are silent when missing.
 *
 * Sub-phase 8.7 lists what has to be right for PowerPoint to use an embedded
 * font. This deck writes all of it, and the reason it is worth a whole fixture
 * is that **only one of the six failures is visible**. Get the content type
 * wrong and PowerPoint says "PowerPoint found a problem with content" - that is
 * the loud one, and it is pandoc issue #11492. Get any of the other five wrong
 * and the file opens, renders in a substitute face, and says nothing.
 *
 * | # | artifact                                              | if missing                |
 * | - | ----------------------------------------------------- | ------------------------- |
 * | 1 | `/ppt/fonts/fontN.fntdata`                            | dangling relationship     |
 * | 2 | `<Default Extension="fntdata" ContentType="application/x-fontdata"/>` | **refused** |
 * | 3 | the `…/relationships/font` rel on `ppt/presentation.xml` | dangling                |
 * | 4 | `p:embeddedFont` in the right slot of `CT_Presentation` | refused, or ignored     |
 * | 5 | `@typeface` byte-identical to the runs' `a:latin`     | silently substituted      |
 * | 6 | `p:presentation/@embedTrueTypeFonts="1"`              | **silently ignored**      |
 *
 * Number six is the cruel one. Without that single attribute PowerPoint does
 * not look at `p:embeddedFontLst` at all, however correct it is.
 *
 * ## The font is ours
 *
 * `tools/ground-truth/fonts/embedding/build-font.ts` writes an SFNT from scratch - `head`,
 * `hhea`, `maxp`, `OS/2`, `hmtx`, `cmap`, `loca`, `glyf`, `name`, `post` - for
 * sub-phase 0.7's experiments A and B, and `eot.ts` wraps it. Both are reused
 * here rather than copied, so this deck embeds a font nobody else has any claim
 * on, with `fsType` **0** because we wrote it and Installable Embedding is the
 * honest answer.
 *
 * The four slots of one typeface are four separate `.fntdata` parts. All four
 * wrap the same outlines: the SFNT builder has one weight, and what the slots
 * differ in is the **EOT header** - `Weight` 400 or 700, the italic flag set or
 * clear - which is what sub-phase 8.2's reader actually reads. A deck that
 * needed four genuinely different faces would need four font designs, and what
 * is being probed here is the package, not the outlines.
 *
 * ## `@charset` is a signed byte
 *
 * `p:font/@charset="0"` is ANSI. The one that catches people is Shift-JIS,
 * which is `"-128"` and not `"128"`: `ST_Charset` is `xsd:byte`, so 0x80 is
 * negative. Slide 2's second typeface writes `-128` for that reason, and it is
 * a real value rather than a demonstration - a Japanese font is embedded that
 * way in every deck that has one.
 *
 * `@pitchFamily` packs two nibbles: `(family << 4) | pitch`. Family 2 is
 * Swiss/sans and pitch 2 is variable, so 34. `@panose` is ten bytes as twenty
 * hex characters, copied from the font's own `OS/2.panose` rather than made up.
 *
 * ## The theme indirection, which is what makes the GC wrong
 *
 * Slide 3's text says `a:latin typeface="+mn-lt"`. Nothing on that slide names
 * `ProbeBravo` - the theme does, in `a:minorFont`. So a save-time collector
 * that scans runs for typefaces, finds no mention of ProbeBravo, and drops it
 * as unused has just broken the file. `+mj-lt` and `+mn-lt` have to be resolved
 * through the theme **before** the used-set is computed, and this is the deck
 * that says so.
 *
 * ## And the two rules PowerPoint enforces that no schema states
 *
 * Each `@typeface` in `p:embeddedFontLst` must be **unique**, and every listed
 * font must actually be **used**. So the list is not "the fonts available", it
 * is "the fonts in use", and it has to be garbage-collected on every save.
 */

/** `application/x-fontdata`. Missing this Default is the one loud failure. */
const CT_FNTDATA = 'application/x-fontdata';

const ALPHA = 'PptxStudio Alpha';
const BRAVO = 'PptxStudio Bravo';

interface Slot {
  readonly slot: 'regular' | 'bold' | 'italic' | 'boldItalic';
  readonly weight: number;
  readonly italic: boolean;
}

const SLOTS: readonly Slot[] = [
  { slot: 'regular', weight: 400, italic: false },
  { slot: 'bold', weight: 700, italic: false },
  { slot: 'italic', weight: 400, italic: true },
  { slot: 'boldItalic', weight: 700, italic: true },
];

/**
 * One `.fntdata` part.
 *
 * EOT version 2.2 with NUL-terminated FamilyName and StyleName, which is what
 * experiment B measured PowerPoint writing. `writeEot` keeps Internet
 * Explorer's old invariant that FullName begins with FamilyName, so the style
 * name is appended rather than substituted.
 */
function fntdata(family: string, slot: Slot): Uint8Array {
  const font = buildProbeFont(family);
  const styleName =
    slot.slot === 'regular'
      ? 'Regular'
      : slot.slot === 'bold'
        ? 'Bold'
        : slot.slot === 'italic'
          ? 'Italic'
          : 'Bold Italic';
  return writeEot(font.bytes, {
    familyName: font.familyName,
    styleName,
    versionName: font.versionName,
    fullName: `${font.familyName} ${styleName}`,
    panose: font.panose,
    charset: 1,
    italic: slot.italic,
    weight: slot.weight,
    // `fsType` 0 is Installable Embedding, and it is true: we wrote this font.
    // `writeEot` copies it into offset 0x20, which upstream ttf2eot leaves at
    // zero whatever the source font said - a claim the OpenType specification
    // makes normative and which sub-phase 8.4 refuses to make on a font's behalf.
    fsType: font.fsType,
    unicodeRange: font.unicodeRange,
    codePageRange: font.codePageRange,
    checkSumAdjustment: font.checkSumAdjustment,
    version: EOT_VERSION_2_2,
    nulTerminateNames: true,
  });
}

const PANOSE = [...buildProbeFont(ALPHA).panose]
  .map((byte) => byte.toString(16).padStart(2, '0').toUpperCase())
  .join('');

/** Family 2 (Swiss) in the high nibble, pitch 2 (variable) in the low one. */
const PITCH_FAMILY = (2 << 4) | 2;

// --------------------------------------------------------------- the parts

interface Embedded {
  readonly typeface: string;
  readonly charset: number;
  readonly slots: readonly Slot[];
}

const EMBEDDED: readonly Embedded[] = [
  { typeface: ALPHA, charset: 0, slots: SLOTS },
  // A single slot, and a charset that is negative. See the file comment.
  { typeface: BRAVO, charset: -128, slots: [SLOTS[0] as Slot] },
];

const parts: ProbePart[] = [];
const presentationRels: ProbeRel[] = [];
const entries: string[] = [];

// The chassis has already allocated rId1..rId8 on ppt/presentation.xml - two
// masters' worth of nothing, the three slides, presProps, viewProps, theme and
// tableStyles - so the font relationships start at rId9. rIds are scoped to one
// .rels part, and this is the one place in the deck that has to know that.
let fontNumber = 0;
let relNumber = 8;
for (const font of EMBEDDED) {
  const slots: string[] = [];
  for (const slot of font.slots) {
    fontNumber += 1;
    relNumber += 1;
    const relId = 'rId' + String(relNumber);
    const name = `ppt/fonts/font${String(fontNumber)}.fntdata`;
    parts.push({
      name,
      bytes: fntdata(font.typeface, slot),
      contentType: { kind: 'default', extension: 'fntdata', type: CT_FNTDATA },
    });
    presentationRels.push({
      id: relId,
      type: REL + 'font',
      target: `fonts/font${String(fontNumber)}.fntdata`,
    });
    slots.push(`<p:${slot.slot} r:id="${relId}"/>`);
  }
  entries.push(
    '<p:embeddedFont>' +
      `<p:font typeface="${font.typeface}" panose="${PANOSE}"` +
      ` pitchFamily="${String(PITCH_FAMILY)}" charset="${String(font.charset)}"/>` +
      // Schema order: font, regular, bold, italic, boldItalic.
      slots.join('') +
      '</p:embeddedFont>',
  );
}

const EMBEDDED_FONT_LST = `<p:embeddedFontLst>${entries.join('')}</p:embeddedFontLst>`;

// --------------------------------------------------------------- the slides

/**
 * Text in the embedded face.
 *
 * `@typeface` here has to be **byte-identical** to the `p:font/@typeface` in
 * the list. DrawingML has no weight axis, so a family whose real name is
 * "Roboto Light" is a *typeface*, not Roboto at weight 300, and matching it by
 * stripping a suffix loses the font.
 */
const sample = (
  id: number,
  y: number,
  typeface: string,
  label: string,
  props: { readonly b?: boolean; readonly i?: boolean } = {},
): string =>
  shape({
    id,
    name: label,
    x: 685800,
    y,
    cx: 10820400,
    cy: 838200,
    textBody: txBody({
      bodyPr: '<a:bodyPr wrap="square" anchor="ctr"/>',
      paras: textLine(`${label} — ABCDEFGHIJ abcdefghij 0123456789`, {
        sz: 2000,
        latin: typeface,
        ...props,
      }),
    }),
  });

const caption = (id: number, y: number, lines: readonly string[]): string =>
  shape({
    id,
    name: 'Caption',
    x: 685800,
    y,
    cx: 10820400,
    cy: 1600200,
    textBody: txBody({
      bodyPr: '<a:bodyPr wrap="square" anchor="t"/>',
      paras: lines.map((line) => textLine(line, { sz: 1300 })).join(''),
    }),
  });

/**
 * One master, whose theme's minor latin is the second embedded typeface.
 *
 * That is the whole point of slide 3: `+mn-lt` resolves here, and nowhere in
 * any slide is `PptxStudio Bravo` written down.
 */
const MASTERS: readonly ProbeMaster[] = [
  {
    minorLatin: BRAVO,
    majorLatin: ALPHA,
    layouts: [
      {
        type: 'titleOnly',
        name: 'Title Only',
        hasTitle: true,
        // The chassis default pair, spelled out because a deck that declares
        // `masters` replaces the layouts outright - and a title placeholder on
        // a slide whose layout has none is the orphan case `a02` is for, not
        // something this deck wants to be quietly exercising.
        shapes: placeholderXml({ id: 2, name: 'Title 1', type: 'title', ...TITLE_BOX }),
      },
      { type: 'blank', name: 'Blank' },
    ],
  },
];

export const a31EmbeddedFonts: ProbeDeck = {
  id: 'a31-embedded-fonts',
  title: 'PPTX Studio corpus: a31 embedded fonts',
  description:
    'All six artifacts sub-phase 8.7 needs, around a font this repository authored: five ' +
    'ppt/fonts/*.fntdata parts in EOT 2.2, the Default Extension="fntdata" whose absence is the ' +
    'one loud failure, five font relationships on ppt/presentation.xml, a p:embeddedFontLst in ' +
    'CT_Presentation slot 8 with all four slots on one typeface and one slot on another, a ' +
    'charset of -128 because ST_Charset is a signed byte, and ' +
    'p:presentation/@embedTrueTypeFonts="1", without which PowerPoint ignores the whole list in ' +
    'silence. The theme’s minor latin is the second typeface, so slide 3 uses it through +mn-lt ' +
    'and never names it - which is what breaks a save-time collector that scans runs alone.',
  features: {
    // 6 chassis + four samples, three more and two captions.
    shape: 15,
    placeholder: 6,
    gradientFill: 2,
    presetGeom: 9,
    embeddedFont: 2,
  },
  build: () => ({
    title: 'PPTX Studio corpus: a31 embedded fonts',
    // Artifact 6. One attribute, and the list is dead without it.
    presentationAttributes: ' embedTrueTypeFonts="1"',
    presentationTail: EMBEDDED_FONT_LST,
    presentationRels,
    parts,
    masters: MASTERS,
    slides: [
      {
        title: 'a31 — four slots of one typeface',
        body:
          sample(10, 1600200, ALPHA, 'regular') +
          sample(11, 2514600, ALPHA, 'bold', { b: true }) +
          sample(12, 3429000, ALPHA, 'italic', { i: true }) +
          sample(13, 4343400, ALPHA, 'bold italic', { b: true, i: true }),
      },
      {
        title: 'a31 — a second typeface, with one slot and a negative charset',
        body:
          sample(10, 1600200, BRAVO, 'regular only') +
          caption(11, 2743200, [
            'p:font/@charset is xsd:byte, so Shift-JIS is "-128" and not "128".',
            'This entry writes -128 to keep the signed reading in the corpus.',
            '@pitchFamily is (family << 4) | pitch: 2 and 2 make 34.',
            '@panose is the font’s own OS/2.panose, ten bytes as twenty hex chars.',
            'A typeface may appear in this list exactly once, and must be used.',
          ]),
      },
      {
        title: 'a31 — a font used only through the theme',
        body:
          sample(10, 1600200, '+mn-lt', 'through +mn-lt') +
          sample(11, 2514600, '+mj-lt', 'through +mj-lt') +
          caption(12, 3657600, [
            'Neither run names a typeface. +mn-lt resolves through the theme to',
            'PptxStudio Bravo and +mj-lt to PptxStudio Alpha, and the string',
            '"PptxStudio Bravo" appears nowhere on any slide in this deck.',
            'A save-time collector that scans runs for typefaces concludes it is',
            'unused, drops the entry and the part, and breaks the file - quietly.',
            'Resolve the theme indirection first, then compute the used set.',
          ]),
      },
    ],
  }),
};
