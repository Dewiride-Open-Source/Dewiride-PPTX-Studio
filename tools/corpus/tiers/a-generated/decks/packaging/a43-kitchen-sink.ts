import { grid, prstGeom, scheme, shape, solidFill, type Cell } from '../shapes.ts';
import { textLine, txBody } from '../text.ts';
import type { ProbeDeck } from '../types.ts';
import type { ProbePart, ProbeSlide } from '../package.ts';
import { cBhvr, clickEffect, mainSequence, makeVisible } from './a17-animations.ts';
import { a21Charts } from './a21-charts.ts';
import { a23SmartArt } from './a23-smartart.ts';
import { a26Ole } from './a26-ole.ts';
import { a32Macros } from './a32-macros.ts';

/**
 * Charts, SmartArt, animations, OLE and macros, in one package - and the only
 * deck in Tier A that is deliberately not about one thing.
 *
 * ## Why it breaks the tier's own rule
 *
 * Every other probe isolates a feature so that a failure names it. `ROSTER.md`
 * is explicit about that and it is the right default: a deck holding six
 * features that fails to round-trip tells you a package broke, not what broke
 * it, and `cli bisect` from 1.5 exists because narrowing after the fact is
 * expensive.
 *
 * This deck exists for the opposite reason. **Gate 1 does not ask whether five
 * features survive; it asks whether they survive together**, and that is a
 * different question with its own failure modes, none of which any of the five
 * source decks can reach:
 *
 * - `[Content_Types].xml` carries four `Default` entries (`xlsx`, `emf`, `vml`,
 *   `bin`) and eleven `Override`s from three different feature families, so the
 *   part that a single missing `Default` corrupts is now shared. That is the
 *   canonical repair prompt, and `a26-ole` alone cannot produce a package where
 *   two families compete for it.
 * - `ppt/embeddings/` and `ppt/vbaProject.bin` both end in a name the OLE deck
 *   and the macro deck each type on their own terms - `.xlsx` by `Default`, and
 *   `bin` by a `Default` that also has to be right for any embedding a chart
 *   drags in.
 * - The `p:timing` trees here target `p:graphicFrame` shapes rather than the
 *   `p:sp` shapes `a17-animations` targets. A `p:spTgt/@spid` naming a frame is
 *   the case where renumbering `p:cNvPr/@id` on export silently unhooks an
 *   animation, and no deck in the corpus had one before this.
 * - The package is macro-enabled **and** holds charts, so its main part carries
 *   the `macroEnabled` content type while four chart parts carry theirs. The
 *   two are independent and nothing but this deck says so.
 *
 * ## How it is built
 *
 * By composition, not by re-authoring. The chart parts, the diagram parts and
 * the OLE parts are `a21-charts`, `a23-smartart` and `a26-ole`'s own - taken
 * from their `build()` output, not copied - and the VBA project is
 * `a32-macros`'s. Markup that already passes an independent census and has been
 * opened in PowerPoint is worth more than the same markup typed a second time,
 * and a merge cannot drift from what it merges.
 *
 * Every slide of all three source decks comes across, which is not tidiness: a
 * probe deck's parts are exactly the parts its slides reach, so taking a subset
 * of the slides would leave chart parts and diagram parts unreachable, and
 * `probes.test.ts` asserts the reachability graph has no orphans. Nine slides
 * in, plus one that says what the deck is.
 *
 * ## What the merge has to check
 *
 * That the four part lists do not collide. They do not - `ppt/charts/`,
 * `ppt/diagrams/`, `ppt/embeddings/` + `ppt/media/` + `ppt/drawings/`, and
 * `ppt/vbaProject.bin` - but "they do not" is a claim about four modules that
 * are free to change, so it is asserted at build time rather than believed.
 * `ppt/charts/colors1.xml` and `ppt/diagrams/colors1.xml` are the near miss.
 *
 * Relationship ids need no reconciliation and that is worth saying once here,
 * because it is the thing people expect to be a problem: an rId is an `xsd:ID`
 * scoped to **one** `.rels` part, so `rId2` on the chart slide and `rId2` on
 * the OLE slide are unrelated names in unrelated files. Slides come across
 * whole, so nothing is ever renumbered.
 *
 * `p:cNvPr/@id` is the one that would be a problem, since it is unique within
 * a part - and it is not one either, for the same reason: a slide arrives with
 * its own shapes and nothing is merged onto it.
 */

const cell = grid(3, 2);

/** A click-to-fade entrance on one shape, in the nest PowerPoint writes. */
function entrance(spid: number): string {
  return (
    mainSequence(
      clickEffect({
        id: 5,
        presetID: 10,
        presetClass: 'entr',
        presetSubtype: 0,
        body:
          makeVisible(6, spid) +
          '<p:animEffect transition="in" filter="fade">' +
          cBhvr({ id: 7, cTnAttributes: 'dur="500"', spid }) +
          '</p:animEffect>',
      }),
    ) +
    `<p:bldLst><p:bldP spid="${String(spid)}" grpId="0"/></p:bldLst>` +
    '</p:timing>'
  );
}

/** The first slide of a group, animated; every other slide unchanged. */
function animateFirst(slides: readonly ProbeSlide[], spid: number): readonly ProbeSlide[] {
  return slides.map((slide, index) => (index === 0 ? { ...slide, tail: entrance(spid) } : slide));
}

/**
 * Concatenate part lists, refusing a collision rather than resolving one.
 *
 * A later part silently winning would be the worst outcome available: the
 * package would still build, still open, and hold one deck's chart under
 * another deck's name.
 */
function mergeParts(...lists: readonly (readonly ProbePart[])[]): readonly ProbePart[] {
  const merged: ProbePart[] = [];
  const seen = new Set<string>();
  for (const list of lists) {
    for (const part of list) {
      if (seen.has(part.name)) {
        throw new Error('a43-kitchen-sink: two source decks both define ' + part.name);
      }
      seen.add(part.name);
      merged.push(part);
    }
  }
  return merged;
}

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

const OVERVIEW =
  box(10, 'Charts', cell(0), 'accent1', [
    'Charts',
    'Four c:chartSpace parts, one with',
    'its own style and colour parts.',
    'Slides 2-4.',
  ]) +
  box(11, 'SmartArt', cell(1), 'accent2', [
    'SmartArt',
    'Three data models, two drawing',
    'fallbacks, one shared layout.',
    'Slides 5-7.',
  ]) +
  box(12, 'OLE', cell(2), 'accent3', [
    'OLE',
    'An embedded workbook, an EMF',
    'preview, a VML drawing, a link.',
    'Slides 8-10.',
  ]) +
  box(13, 'Animations', cell(3), 'accent4', [
    'Animations',
    'A p:timing tree on the first slide',
    'of each group, targeting the frame',
    'rather than a p:sp.',
  ]) +
  box(14, 'Macros', cell(4), 'accent5', [
    'Macros',
    'ppt/vbaProject.bin, and the',
    'macroEnabled content type on',
    'the main part. Hence .pptm.',
  ]) +
  box(15, 'What it is for', cell(5), 'accent6', [
    'Gate 1',
    'Not whether each survives, which',
    'five other decks answer, but',
    'whether they survive together.',
  ]);

const CHARTS = a21Charts.build();
const DIAGRAMS = a23SmartArt.build();
const OLE = a26Ole.build();
const MACROS = a32Macros.build();

export const a43KitchenSink: ProbeDeck = {
  id: 'a43-kitchen-sink',
  title: 'PPTX Studio corpus: a43 kitchen sink',
  description:
    'Charts, SmartArt, animations, OLE and macros in one package, and the only Tier A deck that ' +
    'is deliberately not about one thing. Every other probe isolates a feature so a failure names ' +
    'it; this one exists because Gate 1 asks a different question - not whether five features ' +
    'survive a round trip, but whether they survive it together. That has failure modes none of ' +
    'the five source decks can reach: four Default content types from three feature families ' +
    'sharing one [Content_Types].xml, a macroEnabled main part alongside four chart parts, and ' +
    'p:timing trees whose p:spTgt/@spid names a p:graphicFrame rather than a p:sp - the case ' +
    'where renumbering shape ids on export silently unhooks an animation. Built by merging ' +
    'a21-charts, a23-smartart, a26-ole and a32-macros rather than by re-authoring them, so it ' +
    'cannot drift from the decks it is made of.',
  extension: 'pptm',
  addedIn: 'gate-1',
  // Every count here is the sum of what the source decks contribute plus this
  // deck's own overview slide, and each was read out of the census and checked
  // against that arithmetic rather than pasted. The two that are easy to get
  // wrong: `embeddedPackage` is 1, not 2, because `ppt/vbaProject.bin` is
  // counted by `macros` and not as an embedding - the two rules key on
  // different parts and a deck holding both is the only place that shows. And
  // `vml` is 1: it counts the drawing part, not the `p:oleObj/@spid` that
  // names a shape inside it.
  features: {
    // 10 slide titles + 6 overview boxes + 3 OLE captions + 2 on the master + 1
    // on the Title Only layout.
    shape: 22,
    placeholder: 13,
    presetGeom: 15,
    gradientFill: 2,
    // 4 chart + 3 diagram + 3 OLE.
    graphicFrame: 10,
    chart: 4,
    smartArt: 3,
    smartArtDrawing: 2,
    // Two on the MCE slide - one per branch - and one on each of the other two.
    oleObject: 4,
    embeddedPackage: 1,
    picture: 3,
    vml: 1,
    alternateContent: 1,
    animation: 3,
    macros: 1,
  },
  build: () => ({
    title: 'PPTX Studio corpus: a43 kitchen sink',
    macroEnabled: true,
    parts: mergeParts(
      CHARTS.parts ?? [],
      DIAGRAMS.parts ?? [],
      OLE.parts ?? [],
      MACROS.parts ?? [],
    ),
    // The VBA relationship on `ppt/presentation.xml`, under Microsoft's 2006
    // extension base. `?? []` rather than a spread because a source deck that
    // stopped declaring it should make this deck lose the relationship, not
    // silently drop the field and take the chassis default.
    presentationRels: MACROS.presentationRels ?? [],
    slides: [
      {
        title: 'a43 — five features, one package',
        body: OVERVIEW,
      },
      // The frame on the first slide of each group is `p:cNvPr/@id` 10 in all
      // three source decks, which is why one `spid` serves all three calls.
      ...animateFirst(CHARTS.slides, 10),
      ...animateFirst(DIAGRAMS.slides, 10),
      ...animateFirst(OLE.slides, 10),
    ],
  }),
};
