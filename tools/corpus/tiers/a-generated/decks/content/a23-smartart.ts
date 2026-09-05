import { DECLARATION, NS_A, NS_R, REL, type ProbePart, type ProbeRel } from '../package.ts';
import { graphicFrame, GRAPHIC_URI } from '../shapes.ts';
import type { ProbeDeck } from '../types.ts';

/**
 * SmartArt: four parts, a fifth that is the only one anybody renders, and two
 * ways for that fifth to give you nothing.
 *
 * Read back from a `Shapes.AddSmartArt` on 2026-08-27, so the part names,
 * content types, relationship types and the extension URI are measured.
 *
 * ## Five parts, four relationships, and a fifth reached sideways
 *
 * The frame's `a:graphicData` holds one element:
 *
 * ```xml
 * <dgm:relIds r:dm="rId2" r:lo="rId3" r:qs="rId4" r:cs="rId5"/>
 * ```
 *
 * data, layout, quickStyle, colours - four rIds in the **slide's** rels. The
 * drawing is not among them. It is reached from inside `data1.xml`:
 *
 * ```xml
 * <dgm:extLst><a:ext uri="http://schemas.microsoft.com/office/drawing/2008/diagram">
 *   <dsp:dataModelExt relId="rId6" minVer="…/drawingml/2006/diagram"/>
 * </a:ext></dgm:extLst>
 * ```
 *
 * and `rId6` **resolves against the slide's rels, not the data part's**, even
 * though the element sits inside `data1.xml`. That is the one genuinely
 * surprising edge in the whole SmartArt story and it is why sub-phase 4.5 is
 * its own sub-phase. Note the `@uri` is a plain URI here rather than the GUID
 * every other Office extension uses.
 *
 * ## The drawing is the only thing that can be rendered
 *
 * `dsp:drawing` is an ordinary shape tree under a different prefix:
 * `dsp:spTree` of `dsp:sp`, each `nvSpPr, spPr, style, txBody, txXfrm`. Alias
 * `dsp:` to `p:` and the shape renderer works unchanged - with one addition,
 * `dsp:txXfrm`, the **text** rectangle, which is a different rectangle from the
 * shape's own. Ignore it and text sits in the wrong place inside every chevron,
 * pentagon and arrow SmartArt ships.
 *
 * Two smaller measured facts: every `dsp:cNvPr` in a real drawing has
 * **`id="0"` and `name=""`**, so identity is `dsp:sp/@modelId` and nothing
 * else; and the fill colours are written as `a:schemeClr` carrying
 * `hueOff`/`satOff`/`lumOff`/`alphaOff` of 0, which is the diagram colour
 * engine's output rather than anything a user chose.
 *
 * ## Which is why this deck has three diagrams and only two drawings
 *
 * A renderer that trusts the fallback exists gets two different kinds of
 * nothing, and both are here:
 *
 * 1. **Slide 2** has a `dsp:dataModelExt`, a drawing part, and a `dsp:spTree`
 *    with **no `dsp:sp` children at all**. The relationship resolves, the part
 *    parses, and there is nothing to draw. Sub-phase 4.5's rule - count the
 *    shapes before trusting the fallback - is about exactly this file.
 * 2. **Slide 3** has no `dsp:dataModelExt` at all, so there is no drawing part
 *    to find. That is what every producer before the 2008 extension wrote, and
 *    what PowerPoint itself writes for a diagram it has never laid out.
 *
 * In both cases the honest answer is a labelled placeholder, not a blank
 * rectangle - and the two cases have to be told apart, because only the second
 * one could be repaired by asking PowerPoint to re-lay-out the diagram.
 *
 * ## The identities are written down three times
 *
 * `dgm:pt type="doc"`'s `prSet` names the layout, quick style and colour set by
 * URN - `urn:microsoft.com/office/officeart/2005/8/layout/default` and friends
 * - and each of those URNs appears again as the `@uniqueId` of the part the
 * matching relationship points at. Three copies of one fact, and nothing checks
 * that they agree. This deck keeps them consistent; sub-phase 1.2 has a rule
 * to say so.
 */

const NS_DGM = 'http://schemas.openxmlformats.org/drawingml/2006/diagram';
const NS_DSP = 'http://schemas.microsoft.com/office/drawing/2008/diagram';
/** The `a:ext/@uri` for `dsp:dataModelExt`: a URI, not a GUID. */
const DSP_EXT_URI = 'http://schemas.microsoft.com/office/drawing/2008/diagram';

const REL_DIAGRAM_DRAWING = 'http://schemas.microsoft.com/office/2007/relationships/diagramDrawing';

const CT = {
  data: 'application/vnd.openxmlformats-officedocument.drawingml.diagramData+xml',
  layout: 'application/vnd.openxmlformats-officedocument.drawingml.diagramLayout+xml',
  quickStyle: 'application/vnd.openxmlformats-officedocument.drawingml.diagramStyle+xml',
  colors: 'application/vnd.openxmlformats-officedocument.drawingml.diagramColors+xml',
  drawing: 'application/vnd.ms-office.drawingml.diagramDrawing+xml',
} as const;

const URN = {
  layout: 'urn:microsoft.com/office/officeart/2005/8/layout/default',
  quickStyle: 'urn:microsoft.com/office/officeart/2005/8/quickstyle/simple1',
  colors: 'urn:microsoft.com/office/officeart/2005/8/colors/accent1_2',
} as const;

// ------------------------------------------------------------- data1..3.xml

const EMPTY_TEXT = '<dgm:t><a:bodyPr/><a:lstStyle/><a:p><a:endParaRPr lang="en-GB"/></a:p></dgm:t>';

const nodeText = (text: string): string =>
  '<dgm:t><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr lang="en-GB"/>' +
  `<a:t>${text}</a:t></a:r></a:p></dgm:t>`;

/** A `dgm:pt`. `CT_Pt` is prSet, spPr, t, extLst. */
const pt = (modelId: string, body: string, attributes = ''): string =>
  `<dgm:pt modelId="${modelId}"${attributes}>${body}</dgm:pt>`;

const cxn = (modelId: string, srcId: string, destId: string, srcOrd: number, extra = ''): string =>
  `<dgm:cxn modelId="${modelId}" srcId="${srcId}" destId="${destId}"` +
  ` srcOrd="${String(srcOrd)}" destOrd="0"${extra}/>`;

/** Ids are fixed literals: a fixture whose bytes move on every build is not one. */
const ID = {
  doc: '{D0AA6F5F-69ED-46F2-9A2E-7EBAD279AA3C}',
  node: [
    '{30CF3CC2-09A7-498A-A84C-3A63B345D038}',
    '{CF493E2B-C32E-45E0-A157-B5DBDA93D192}',
    '{1CAF74B3-2B84-4A08-9F5F-9E8BD0F0E9B1}',
  ],
  parTrans: [
    '{DDD4C78F-253E-4C98-9251-AE19EF14FD0C}',
    '{7D0FDD72-9598-4D47-A4A0-B359B018CEDF}',
    '{9B0A9A0C-4B7C-4B8D-8A45-1B96F0E1C93A}',
  ],
  sibTrans: [
    '{2A3C321A-EC19-40D2-836E-A147765C09D5}',
    '{4F8B7A32-CE47-4E5A-9B3C-7B0E86D6E9E7}',
    '{6D2E38A1-93C7-4A2E-B2D1-8C3A5F0E4B77}',
  ],
  cxn: [
    '{26B7DBC3-D671-48C2-866A-860CE40B2829}',
    '{637615D7-2108-4B58-9F1E-2C0B1F5D9A44}',
    '{B1F2C3D4-5E6F-4A7B-8C9D-0E1F2A3B4C5D}',
  ],
  pres: '{8B6A5C4D-3E2F-4109-8A7B-6C5D4E3F2A1B}',
} as const;

function dataModel(spec: {
  readonly labels: readonly string[];
  readonly drawingRelId?: string;
}): string {
  const points = spec.labels
    .map((label, index) => {
      const node = ID.node[index] ?? '{00000000-0000-0000-0000-000000000000}';
      const par = ID.parTrans[index] ?? '{00000000-0000-0000-0000-000000000001}';
      const sib = ID.sibTrans[index] ?? '{00000000-0000-0000-0000-000000000002}';
      const conn = ID.cxn[index] ?? '{00000000-0000-0000-0000-000000000003}';
      return (
        pt(node, '<dgm:prSet phldrT="[Text]"/><dgm:spPr/>' + nodeText(label)) +
        pt(par, '<dgm:prSet/><dgm:spPr/>' + EMPTY_TEXT, ` type="parTrans" cxnId="${conn}"`) +
        pt(sib, '<dgm:prSet/><dgm:spPr/>' + EMPTY_TEXT, ` type="sibTrans" cxnId="${conn}"`)
      );
    })
    .join('');

  const connections = spec.labels
    .map((_label, index) => {
      const node = ID.node[index] ?? '';
      const par = ID.parTrans[index] ?? '';
      const sib = ID.sibTrans[index] ?? '';
      const conn = ID.cxn[index] ?? '';
      return cxn(conn, ID.doc, node, index, ` parTransId="${par}" sibTransId="${sib}"`);
    })
    .join('');

  return (
    DECLARATION +
    `<dgm:dataModel xmlns:dgm="${NS_DGM}" xmlns:a="${NS_A}">` +
    '<dgm:ptLst>' +
    // The document node names all three identities, which are also the three
    // parts' `@uniqueId` values. Three copies of one fact.
    pt(
      ID.doc,
      `<dgm:prSet loTypeId="${URN.layout}" loCatId="list" qsTypeId="${URN.quickStyle}"` +
        ` qsCatId="simple" csTypeId="${URN.colors}" csCatId="accent1" phldr="0"/>` +
        '<dgm:spPr/>' +
        EMPTY_TEXT,
      ' type="doc"',
    ) +
    points +
    '</dgm:ptLst>' +
    `<dgm:cxnLst>${connections}</dgm:cxnLst>` +
    '<dgm:bg/><dgm:whole/>' +
    (spec.drawingRelId === undefined
      ? ''
      : '<dgm:extLst>' +
        `<a:ext uri="${DSP_EXT_URI}">` +
        `<dsp:dataModelExt xmlns:dsp="${NS_DSP}" relId="${spec.drawingRelId}"` +
        ` minVer="${NS_DGM}"/>` +
        '</a:ext></dgm:extLst>') +
    '</dgm:dataModel>'
  );
}

// ---------------------------------------------------------- drawing1..2.xml

/**
 * One `dsp:sp`.
 *
 * `@id="0"` and `@name=""` on every `dsp:cNvPr` is measured, not a shortcut:
 * PowerPoint numbers nothing here, because `dsp:sp/@modelId` is the identity
 * and it is the key back into `data1.xml`'s `dgm:ptLst`.
 */
function dspShape(spec: {
  readonly modelId: string;
  readonly x: number;
  readonly y: number;
  readonly cx: number;
  readonly cy: number;
  readonly accent: string;
  readonly text: string;
}): string {
  const off = `<a:off x="${String(spec.x)}" y="${String(spec.y)}"/>`;
  const ext = `<a:ext cx="${String(spec.cx)}" cy="${String(spec.cy)}"/>`;
  const offsets = '<a:hueOff val="0"/><a:satOff val="0"/><a:lumOff val="0"/><a:alphaOff val="0"/>';
  return (
    `<dsp:sp modelId="${spec.modelId}">` +
    '<dsp:nvSpPr><dsp:cNvPr id="0" name=""/><dsp:cNvSpPr/></dsp:nvSpPr>' +
    '<dsp:spPr>' +
    `<a:xfrm>${off}${ext}</a:xfrm>` +
    '<a:prstGeom prst="roundRect"><a:avLst/></a:prstGeom>' +
    `<a:solidFill><a:schemeClr val="${spec.accent}">${offsets}</a:schemeClr></a:solidFill>` +
    '<a:ln w="19050" cap="flat" cmpd="sng" algn="ctr">' +
    `<a:solidFill><a:schemeClr val="lt1">${offsets}</a:schemeClr></a:solidFill>` +
    '<a:prstDash val="solid"/><a:miter lim="800000"/></a:ln>' +
    '<a:effectLst/>' +
    '</dsp:spPr>' +
    '<dsp:style>' +
    '<a:lnRef idx="2"><a:scrgbClr r="0" g="0" b="0"/></a:lnRef>' +
    '<a:fillRef idx="1"><a:scrgbClr r="0" g="0" b="0"/></a:fillRef>' +
    '<a:effectRef idx="0"><a:scrgbClr r="0" g="0" b="0"/></a:effectRef>' +
    '<a:fontRef idx="minor"><a:schemeClr val="lt1"/></a:fontRef>' +
    '</dsp:style>' +
    '<dsp:txBody>' +
    '<a:bodyPr spcFirstLastPara="0" vert="horz" wrap="square" lIns="106680" tIns="106680"' +
    ' rIns="106680" bIns="106680" numCol="1" spcCol="1270" anchor="ctr" anchorCtr="0">' +
    '<a:noAutofit/></a:bodyPr><a:lstStyle/>' +
    '<a:p><a:pPr marL="0" lvl="0" indent="0" algn="ctr" defTabSz="1111250">' +
    '<a:lnSpc><a:spcPct val="90000"/></a:lnSpc><a:spcBef><a:spcPct val="0"/></a:spcBef>' +
    '<a:spcAft><a:spcPct val="35000"/></a:spcAft><a:buNone/></a:pPr>' +
    `<a:r><a:rPr lang="en-GB" sz="2500" kern="1200"/><a:t>${spec.text}</a:t></a:r></a:p>` +
    '</dsp:txBody>' +
    // The text rectangle, and it is inset from the shape's own on purpose.
    `<dsp:txXfrm><a:off x="${String(spec.x + 68580)}" y="${String(spec.y + 68580)}"/>` +
    `<a:ext cx="${String(spec.cx - 137160)}" cy="${String(spec.cy - 137160)}"/></dsp:txXfrm>` +
    '</dsp:sp>'
  );
}

const drawing = (shapes: string): string =>
  DECLARATION +
  `<dsp:drawing xmlns:dgm="${NS_DGM}" xmlns:dsp="${NS_DSP}" xmlns:a="${NS_A}">` +
  '<dsp:spTree>' +
  '<dsp:nvGrpSpPr><dsp:cNvPr id="0" name=""/><dsp:cNvGrpSpPr/></dsp:nvGrpSpPr>' +
  '<dsp:grpSpPr/>' +
  shapes +
  '</dsp:spTree></dsp:drawing>';

const LABELS = ['Parse', 'Resolve', 'Render'];

const DRAWING_1 = drawing(
  LABELS.map((text, index) =>
    dspShape({
      modelId: ID.node[index] ?? '',
      x: 76200 + index * 1676400,
      y: 762000,
      cx: 1524000,
      cy: 914400,
      accent: 'accent1',
      text,
    }),
  ).join(''),
);

/** The second kind of nothing: a drawing part whose tree is empty. */
const DRAWING_2 = drawing('');

// ------------------------------------------- layout, quickStyle and colours

/**
 * A minimal `dgm:layoutDef`.
 *
 * PowerPoint re-runs the layout engine only when a user edits the diagram, and
 * this project never runs it at all - sub-phase 4.6 renders the drawing and
 * says so. So this part exists to make the `r:lo` relationship resolve and to
 * carry the `@uniqueId` the data model names, and it is deliberately a fraction
 * of the 4 KB PowerPoint writes rather than a copy of it.
 */
const LAYOUT =
  DECLARATION +
  `<dgm:layoutDef xmlns:dgm="${NS_DGM}" xmlns:a="${NS_A}" xmlns:r="${NS_R}"` +
  ` uniqueId="${URN.layout}">` +
  '<dgm:title val=""/><dgm:desc val=""/>' +
  '<dgm:catLst><dgm:cat type="list" pri="400"/></dgm:catLst>' +
  '<dgm:layoutNode name="diagram">' +
  '<dgm:varLst><dgm:dir val="norm"/></dgm:varLst>' +
  '<dgm:alg type="lin"><dgm:param type="linDir" val="fromL"/></dgm:alg>' +
  '<dgm:shape xmlns:r2="' +
  NS_R +
  '" type="none" r2:blip=""><dgm:adjLst/></dgm:shape>' +
  '<dgm:presOf/><dgm:constrLst/><dgm:ruleLst/>' +
  '<dgm:forEach name="nodes" axis="ch" ptType="node">' +
  '<dgm:layoutNode name="node">' +
  '<dgm:alg type="tx"/>' +
  '<dgm:shape xmlns:r2="' +
  NS_R +
  '" type="roundRect" r2:blip=""><dgm:adjLst/></dgm:shape>' +
  '<dgm:presOf axis="desOrSelf" ptType="node"/><dgm:constrLst/><dgm:ruleLst/>' +
  '</dgm:layoutNode></dgm:forEach>' +
  '</dgm:layoutNode></dgm:layoutDef>';

const QUICK_STYLE =
  DECLARATION +
  `<dgm:styleDef xmlns:dgm="${NS_DGM}" xmlns:a="${NS_A}" uniqueId="${URN.quickStyle}">` +
  '<dgm:title val=""/><dgm:desc val=""/>' +
  '<dgm:catLst><dgm:cat type="simple" pri="10100"/></dgm:catLst>' +
  '<dgm:scene3d><a:camera prst="orthographicFront"/><a:lightRig rig="threePt" dir="t"/></dgm:scene3d>' +
  '<dgm:styleLbl name="node0">' +
  '<dgm:scene3d><a:camera prst="orthographicFront"/><a:lightRig rig="threePt" dir="t"/></dgm:scene3d>' +
  '<dgm:sp3d/><dgm:txPr/>' +
  '<dgm:style>' +
  '<a:lnRef idx="2"><a:scrgbClr r="0" g="0" b="0"/></a:lnRef>' +
  '<a:fillRef idx="1"><a:scrgbClr r="0" g="0" b="0"/></a:fillRef>' +
  '<a:effectRef idx="0"><a:scrgbClr r="0" g="0" b="0"/></a:effectRef>' +
  '<a:fontRef idx="minor"><a:schemeClr val="lt1"/></a:fontRef>' +
  '</dgm:style></dgm:styleLbl></dgm:styleDef>';

const COLORS =
  DECLARATION +
  `<dgm:colorsDef xmlns:dgm="${NS_DGM}" xmlns:a="${NS_A}" uniqueId="${URN.colors}">` +
  '<dgm:title val=""/><dgm:desc val=""/>' +
  '<dgm:catLst><dgm:cat type="accent1" pri="11200"/></dgm:catLst>' +
  '<dgm:styleLbl name="node0">' +
  '<dgm:fillClrLst meth="repeat"><a:schemeClr val="accent1"/></dgm:fillClrLst>' +
  '<dgm:linClrLst meth="repeat"><a:schemeClr val="lt1"/></dgm:linClrLst>' +
  '<dgm:effectClrLst/><dgm:txLinClrLst/><dgm:txFillClrLst/><dgm:txEffectClrLst/>' +
  '</dgm:styleLbl></dgm:colorsDef>';

// ------------------------------------------------------------------ package

const PARTS: readonly ProbePart[] = [
  {
    name: 'ppt/diagrams/data1.xml',
    bytes: dataModel({ labels: LABELS, drawingRelId: 'rId6' }),
    contentType: { kind: 'override', type: CT.data },
  },
  {
    name: 'ppt/diagrams/data2.xml',
    bytes: dataModel({ labels: ['Empty', 'Fallback'], drawingRelId: 'rId6' }),
    contentType: { kind: 'override', type: CT.data },
  },
  {
    // No `dgm:extLst`, so nothing points at a drawing and none exists.
    name: 'ppt/diagrams/data3.xml',
    bytes: dataModel({ labels: ['No', 'Drawing'] }),
    contentType: { kind: 'override', type: CT.data },
  },
  {
    name: 'ppt/diagrams/layout1.xml',
    bytes: LAYOUT,
    contentType: { kind: 'override', type: CT.layout },
  },
  {
    name: 'ppt/diagrams/quickStyle1.xml',
    bytes: QUICK_STYLE,
    contentType: { kind: 'override', type: CT.quickStyle },
  },
  {
    name: 'ppt/diagrams/colors1.xml',
    bytes: COLORS,
    contentType: { kind: 'override', type: CT.colors },
  },
  {
    name: 'ppt/diagrams/drawing1.xml',
    bytes: DRAWING_1,
    contentType: { kind: 'override', type: CT.drawing },
  },
  {
    name: 'ppt/diagrams/drawing2.xml',
    bytes: DRAWING_2,
    contentType: { kind: 'override', type: CT.drawing },
  },
];

/** The four rIds the frame names, plus the fifth only `data1.xml` knows about. */
function diagramRels(dataNumber: number, drawingNumber: number | null): readonly ProbeRel[] {
  const rels: ProbeRel[] = [
    { id: 'rId2', type: REL + 'diagramData', target: `../diagrams/data${String(dataNumber)}.xml` },
    { id: 'rId3', type: REL + 'diagramLayout', target: '../diagrams/layout1.xml' },
    { id: 'rId4', type: REL + 'diagramQuickStyle', target: '../diagrams/quickStyle1.xml' },
    { id: 'rId5', type: REL + 'diagramColors', target: '../diagrams/colors1.xml' },
  ];
  if (drawingNumber !== null) {
    rels.push({
      id: 'rId6',
      type: REL_DIAGRAM_DRAWING,
      target: `../diagrams/drawing${String(drawingNumber)}.xml`,
    });
  }
  return rels;
}

const diagramFrame = (id: number, name: string): string =>
  graphicFrame({
    id,
    name,
    x: 914400,
    y: 1600200,
    cx: 10363200,
    cy: 4114800,
    uri: GRAPHIC_URI.diagram,
    content:
      `<dgm:relIds xmlns:dgm="${NS_DGM}" xmlns:r="${NS_R}"` +
      ' r:dm="rId2" r:lo="rId3" r:qs="rId4" r:cs="rId5"/>',
  });

export const a23SmartArt: ProbeDeck = {
  id: 'a23-smartart',
  title: 'PPTX Studio corpus: a23 smartart',
  description:
    'Three SmartArt diagrams, five diagram parts and two drawing fallbacks. dgm:relIds names four ' +
    'rIds; the fifth, the dsp: drawing, is reached from dsp:dataModelExt inside data1.xml and ' +
    'resolves against the slide rels rather than the data part rels. Slide 1 has a three-shape ' +
    'drawing with dsp:txXfrm; slide 2 has a drawing part whose dsp:spTree is empty; slide 3 has no ' +
    'dsp:dataModelExt at all, so there is no drawing part to find. Those are two different kinds ' +
    'of nothing and only one of them could be repaired.',
  features: {
    shape: 6,
    placeholder: 6,
    gradientFill: 2,
    graphicFrame: 3,
    smartArt: 3,
    smartArtDrawing: 2,
    presetGeom: 3,
  },
  build: () => ({
    title: 'PPTX Studio corpus: a23 smartart',
    parts: PARTS,
    slides: [
      {
        title: 'a23 — a diagram, its four parts and the drawing that renders it',
        rels: diagramRels(1, 1),
        body: diagramFrame(10, 'Diagram 1'),
      },
      {
        title: 'a23 — a drawing fallback with no shapes in it',
        rels: diagramRels(2, 2),
        body: diagramFrame(10, 'Diagram 2'),
      },
      {
        title: 'a23 — no dsp:dataModelExt, so no drawing at all',
        rels: diagramRels(3, null),
        body: diagramFrame(10, 'Diagram 3'),
      },
    ],
  }),
};
