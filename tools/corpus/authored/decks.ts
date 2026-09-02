/**
 * Tier B's roster, as reviewed literals.
 *
 * The shape mirrors `tools/corpus/gen/decks/*.ts` on purpose: each entry states
 * the census it should produce, and `decks.test.ts` asserts exact equality
 * against a fresh `censusPackage()` of the committed bytes, in both directions.
 * A map derived from the file would restate the file; a map somebody typed after
 * looking at the file is a claim that can be wrong, which is the only kind of
 * claim worth checking.
 *
 * What is different from Tier A, and why
 * --------------------------------------
 * No `build`. Tier A's decks are functions returning markup; these are bytes
 * PowerPoint wrote, and the closest thing to a build function is
 * `build-tier-b.ps1`, which is committed beside this file. No `recipe` either -
 * see `schema.ts`'s note on the `authored` collection. Re-running the script
 * produces a different file every time, because `dcterms:created` is a
 * timestamp, so a recipe would be a claim this repository cannot honour.
 *
 * The baseline every entry carries
 * ---------------------------------
 * Tier A has `a01-minimal` as its subtrahend. Tier B's is `b01-blank`, and it is
 * much larger, because PowerPoint's floor is not ours: 37 entries, all eleven
 * layouts whether a slide uses them or not, a `docProps/thumbnail.jpeg`, and a
 * master whose `p:txStyles` and theme carry effects. So every Tier B deck reports
 * at least `{placeholder: 65, shape: 65, field: 24, presetGeom: 5,
 * gradientFill: 3, shadow: 1, thumbnail: 1}` and the interesting part of any
 * entry below is what exceeds that.
 *
 * `field: 24` is the date, footer and slide-number placeholders across eleven
 * layouts and the master. `shadow: 1` comes from the theme's `effectStyleLst`.
 * `thumbnail: 1` is on every one of these and on exactly one Tier A deck, which
 * is `a33-thumbnail` and which had to be written deliberately - PowerPoint emits
 * it unasked.
 */

export interface AuthoredDeck {
  readonly id: string;
  readonly slides: number;
  readonly description: string;
  readonly features: Readonly<Record<string, number>>;
}

export const AUTHORED_DECKS: readonly AuthoredDeck[] = [
  {
    id: 'b01-blank',
    slides: 1,
    description:
      "PowerPoint's floor, and the file every lexical claim in this corpus is checked against. One " +
      'slide on the Title Slide layout with its two placeholders, which is a01-minimal as Microsoft ' +
      'writes it - and Microsoft writes 37 ZIP entries where we write 18, because all eleven layouts ' +
      'ship whether a slide uses them or not and docProps/thumbnail.jpeg is emitted unasked. Three ' +
      'divergences from our own generator live here and nowhere smaller: p:sldSz carries NO @type at ' +
      '12192000x6858000 where our chassis writes type="screen16x9"; _rels/.rels is written rId3, ' +
      'rId2, rId1, rId4; and the thumbnail is the one STORED entry in an archive that deflates ' +
      'everything else. p:presentation also carries saveSubsetFonts="1" and removePersonalInfo="1".',
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
  {
    id: 'b02-layouts',
    slides: 11,
    description:
      'One slide per built-in layout, bound through Slides.AddSlide(index, CustomLayout) so slide i ' +
      'resolves to slideLayout i for all eleven. This is the fixture sub-phase 7.1 needs: the ' +
      '121-case matcher matrix is these eleven layouts crossed with themselves, and the (type, idx) ' +
      'pairs in it have to be the ones PowerPoint really allocates rather than ones we invented. ' +
      'Measured placeholder types per layout: Title Slide ctrTitle+subTitle, Title and Content ' +
      'title+obj, Section Header title+body, Two Content title+obj+obj, Comparison ' +
      'title+body+obj+body+obj, Title Only title, Blank none, Content with Caption title+obj+body, ' +
      'Picture with Caption title+pic+body, Title and Vertical Text title+vertBody, Vertical Title ' +
      'and Text vertTitle+vertBody. Every layout also carries dt, ftr and sldNum, and none of those ' +
      'three is inherited onto a slide unless p:hf turns it on - which is why a slide on Title Slide ' +
      'reports two shapes and not five.',
    features: {
      placeholder: 88,
      shape: 88,
      field: 24,
      presetGeom: 5,
      gradientFill: 3,
      shadow: 1,
      thumbnail: 1,
    },
  },
  {
    id: 'b03-text',
    slides: 3,
    description:
      'Autofit as PowerPoint computes it, which is the one thing sub-phase 3.4 cannot check against a ' +
      'deck we wrote: the discrete fourteen-step ladder is what is being verified, so the fixture has ' +
      'to come from the implementation that defines it. Slide 1 overflows a body placeholder set to ' +
      'msoAutoSizeTextToFitShape, so a:normAutofit arrives with the @fontScale and @lnSpcReduction ' +
      'PowerPoint chose. Slide 2 carries bold and italic runs inside one paragraph, a second indent ' +
      'level, an explicit run colour, and a slide-number field with the ST_Guid PowerPoint allocated ' +
      'rather than one we made up. Slide 3 is a plain text box on the Blank layout, which is not a ' +
      'placeholder and so inherits from otherStyle rather than bodyStyle - the branch of the ' +
      'ten-source cascade that defaultTextStyle applies to.',
    features: {
      shape: 70,
      placeholder: 69,
      field: 25,
      presetGeom: 6,
      gradientFill: 3,
      shadow: 1,
      thumbnail: 1,
    },
  },
  {
    id: 'b04-table',
    slides: 1,
    description:
      'A 4x4 table carrying the built-in style {5C22544A-7EE6-4342-B048-85BDC9FD1C3A} - "Medium Style ' +
      '2 - Accent 1", which is the default a new table gets - with FirstRow, LastRow, FirstCol and ' +
      'HorizBanding on, LastCol and VertBanding off, one horizontal merge and one vertical one. The ' +
      'deck exists to prove a negative: ppt/tableStyles.xml in this package is an EMPTY ' +
      'a:tblStyleLst whose @def is that same GUID, so the style the table names is defined nowhere ' +
      'in the file. A renderer that reads only that part draws this table white and borderless, ' +
      'which is why sub-phase 4.2 has to re-derive all 74 definitions instead of parsing them.',
    features: {
      placeholder: 64,
      shape: 64,
      field: 24,
      presetGeom: 5,
      gradientFill: 3,
      graphicFrame: 1,
      shadow: 1,
      table: 1,
      thumbnail: 1,
    },
  },
  {
    id: 'b05-chart',
    slides: 2,
    description:
      'Two charts, and the second is the one that matters: its c:plotArea holds a c:barChart AND a ' +
      'c:lineChart, so code that reaches for firstElementChild silently drops the line overlay. That ' +
      'is the exact failure sub-phase 9.8 names, in a file Microsoft wrote. Both charts bring the ' +
      'five parts a chart really costs - chart1.xml, its .rels, colors1.xml, style1.xml and an ' +
      'embedded workbook - and the presence of colors1.xml is the point: c:ser/c:spPr is absent ' +
      'here, so a renderer that skips the colour-style part gets every series wrong even with a ' +
      'correct theme. The workbooks are new and empty, created in place; no file on the authoring ' +
      'machine was opened to make them.',
    features: {
      placeholder: 65,
      shape: 65,
      field: 24,
      presetGeom: 5,
      gradientFill: 3,
      alternateContent: 2,
      chart: 2,
      embeddedPackage: 2,
      graphicFrame: 2,
      shadow: 1,
      thumbnail: 1,
    },
  },
  {
    id: 'b06-smartart',
    slides: 3,
    description:
      'Three SmartArt families - Basic Block List, Organization Chart and Basic Process - each with ' +
      'its five parts: data1.xml, layout1.xml, quickStyle1.xml, colors1.xml and the drawing1.xml ' +
      'fallback. The fallback is what sub-phase 4.6 renders, and this deck is where its shape gets ' +
      'settled: 20 preset geometries and 4 custom ones across the three diagrams, all inside dsp: ' +
      'elements that alias onto the p: shape renderer. It is also the fixture for the resolution ' +
      'chain dgm:relIds/@r:dm -> data1.xml -> dsp:dataModelExt/@relId, where the last hop resolves ' +
      "against the SLIDE's rels even though the element lives in the data part.",
    features: {
      placeholder: 66,
      shape: 66,
      field: 24,
      presetGeom: 20,
      customGeom: 4,
      gradientFill: 3,
      graphicFrame: 3,
      smartArt: 3,
      smartArtDrawing: 3,
      shadow: 1,
      thumbnail: 1,
    },
  },
  {
    id: 'b07-ole',
    slides: 1,
    description:
      'An embedded Excel worksheet, and the deck experiment E6 was measured from - committed so the ' +
      'measurement can be diffed rather than trusted. This build still writes the ' +
      'mc:AlternateContent wrapper whose mc:Choice requires the VML prefix, but there is NO ' +
      'vmlDrawing part and no @spid for one to point at: Microsoft stopped emitting the fallback at ' +
      'build 2205 while continuing to read it. p:oleObj carries exactly five attributes - name, ' +
      'r:id, imgW, imgH, progId - and the preview is an EMF that appears in the mc:Fallback only. ' +
      'oleObject counts 2 and alternateContent 1 because both switch branches are scanned, which is ' +
      'the census behaving correctly rather than double-counting. ppt/embeddings/*.bin is never ' +
      'touched by us; here the embedding is a .xlsx, stored rather than deflated.',
    features: {
      placeholder: 64,
      shape: 64,
      field: 24,
      presetGeom: 6,
      gradientFill: 3,
      oleObject: 2,
      alternateContent: 1,
      embeddedPackage: 1,
      graphicFrame: 1,
      picture: 1,
      shadow: 1,
      thumbnail: 1,
    },
  },
  {
    id: 'b08-transitions',
    slides: 5,
    description:
      'Five slide transitions and two animation effects apiece, and it carries the best MCE fixture ' +
      'in the corpus because Microsoft wrote it. PowerPoint wraps every p:transition in ' +
      'mc:AlternateContent; for a post-2010 effect the mc:Choice holds p14:honeycomb while the ' +
      'mc:Fallback holds a plain p:fade. The two branches are NOT the same transition, so a consumer ' +
      'that resolves mc:Choice/@Requires wrongly shows something different rather than something ' +
      'degraded - the failure the whole MCE walker exists to prevent, which a37-mce could only ' +
      'simulate. transition counts 10 for five slides because both branches are scanned. Timing ' +
      'trees use presetID/presetClass, an id space separate from MsoAnimEffect, with clickEffect and ' +
      'withPrevious nodes and one fly-in so a slide has three effects rather than two.',
    features: {
      shape: 78,
      placeholder: 68,
      field: 24,
      presetGeom: 15,
      transition: 10,
      alternateContent: 5,
      animation: 5,
      gradientFill: 3,
      shadow: 1,
      thumbnail: 1,
    },
  },
  {
    id: 'b09-picture',
    slides: 2,
    description:
      'Pictures the two ways PowerPoint places them, with the crop and the effects it writes for ' +
      'each. Slide 1 has a PNG cropped on all four sides - so a:srcRect carries four different ' +
      'non-zero percentages rather than the single one a simpler deck produces - beside a JPEG ' +
      'carrying shadow, reflection, glow and soft edge together. Slide 2 fills the pic placeholder ' +
      'of the Picture with Caption layout instead, which arrives as a blipFill on a placeholder and ' +
      'not as a bare p:pic; the two inherit differently and only one survives a layout change. Both ' +
      "images are this repository's own, written by make-assets.ts from tools/corpus/gen/png.ts " +
      'and jpeg.ts - there is no stock imagery in this corpus and no photograph from the authoring ' +
      'machine.',
    features: {
      placeholder: 67,
      shape: 67,
      field: 24,
      presetGeom: 7,
      gradientFill: 3,
      picture: 2,
      shadow: 2,
      blipFill: 1,
      glow: 1,
      reflection: 1,
      softEdge: 1,
      thumbnail: 1,
    },
  },
];
