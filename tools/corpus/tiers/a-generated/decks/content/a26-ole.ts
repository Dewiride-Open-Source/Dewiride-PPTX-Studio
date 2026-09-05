import { relsXml, REL, type ProbePart, type ProbeRel } from '../../markup/chassis.ts';
import { probeEmf } from '../../assets/emf.ts';
import { graphicFrame, GRAPHIC_URI, picture, shape } from '../../markup/shapes.ts';
import { textLine, txBody } from '../../markup/text.ts';
import { probeXlsx } from '../../binary/xlsx.ts';
import type { ProbeDeck } from '../../markup/types.ts';

/**
 * OLE objects, in the shape experiment E6 measured rather than the one the
 * research predicted.
 *
 * Build 2205 stopped writing the VML fallback for OLE. It did **not** stop
 * writing the wrapper, and what 16.0.20326 writes for an embedded workbook is
 *
 * ```xml
 * <mc:AlternateContent>
 *   <mc:Choice xmlns:v="urn:schemas-microsoft-com:vml" Requires="v">
 *     <p:oleObj name="Worksheet" r:id="rId2" imgW="…" imgH="…"
 *               progId="Excel.Sheet.12"><p:embed/></p:oleObj>
 *   </mc:Choice>
 *   <mc:Fallback>
 *     <p:oleObj …><p:embed/><p:pic>… <a:blip r:embed="rId3"/> …</p:pic></p:oleObj>
 *   </mc:Fallback>
 * </mc:AlternateContent>
 * ```
 *
 * Three things in that are worth saying out loud.
 *
 * **The Choice requires `v` and there is no VML part.** The prefix is declared
 * on the `mc:Choice` itself and never used for an element. A reader that
 * resolves `Requires="v"`, decides the package must contain a `vmlDrawing`, and
 * treats its absence as corruption is wrong about a file PowerPoint wrote
 * itself - the requirement is on the *consumer's* vocabulary, not on the
 * package's contents.
 *
 * **The preview lives only in the `mc:Fallback`.** Resolve MCE correctly, take
 * the Choice, and there is nothing to draw at all: the Choice branch is a
 * `p:oleObj` with no `p:pic` in it. So the rendering path for an OLE object is
 * to read the branch you would otherwise discard, which is the one case where
 * "prefer the Choice" is exactly the wrong instinct.
 *
 * **The preview is an EMF.** Not a PNG. Sub-phase 10.6's rasterizer is on the
 * critical path for OLE, not an extra.
 *
 * ## Two frames the census counts as four objects
 *
 * The census reads raw, pre-MCE markup, so one logical OLE object inside an
 * `mc:AlternateContent` reports `oleObject` **twice** - once in the Choice and
 * once in the Fallback. That is the right policy for a census, which is meant
 * to say what is in the file rather than what a resolver would keep, and the
 * `features` map here has to say so too.
 *
 * ## The payload is a package, not a compound file
 *
 * `ppt/embeddings/…xlsx` through a `…/relationships/package` relationship, with
 * a `Default Extension="xlsx"`. That is what made this deck writable at all:
 * an OPC package is something this repository can author byte for byte, and a
 * CFB compound file is not. `tools/corpus/tiers/a-generated/binary/xlsx.ts` writes it - five parts,
 * inline strings, no shared string table.
 *
 * The must-not-break rule stands unchanged and is about the other direction:
 * **never rewrite `ppt/embeddings/*`**. Authoring one for a fixture is fine;
 * touching one on export is a guaranteed repair prompt.
 *
 * ## Slide 2 is the transitional shape, and slide 3 is a link
 *
 * Slide 2 writes the pre-2205 form deliberately: a bare `p:oleObj` with
 * `@spid`, no `mc:AlternateContent`, and a real `vmlDrawing` part whose
 * `v:shape/@id` is the same `_x0000_s1026`. That edge - PresentationML
 * attribute to VML element id, through a part that has its own relationships -
 * is the one sub-phase 10.9 has to keep in sync when `p:cNvPr` ids are
 * renumbered, and it exists in a very large number of real decks.
 *
 * Slide 3 takes the other branch of `p:oleObj`'s choice: `p:link` instead of
 * `p:embed`, with an external relationship and no local bytes at all.
 */

const NS_V = 'urn:schemas-microsoft-com:vml';
const NS_O = 'urn:schemas-microsoft-com:office:office';
const NS_MC = 'http://schemas.openxmlformats.org/markup-compatibility/2006';

const CT_VML = 'application/vnd.openxmlformats-officedocument.vmlDrawing';
const CT_XLSX = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';

/** The shape id PresentationML and VML have to agree on. */
const SPID = '_x0000_s1026';

/** Measured on the E6 deck: the preview's natural size in EMU. */
const IMG_W = 5186562;
const IMG_H = 3805404;

const FRAME = { x: 914400, y: 1600200, cx: 5186562, cy: 3805404 } as const;

// ------------------------------------------------------------------ the OLE

/** The `p:pic` a fallback carries. Everything about it is ordinary. */
const preview = (id: number, relId: string): string =>
  picture({
    id,
    name: 'Object preview',
    relId,
    x: FRAME.x,
    y: FRAME.y,
    cx: FRAME.cx,
    cy: FRAME.cy,
    description: '',
  });

function oleObj(spec: {
  readonly name: string;
  readonly relId: string;
  readonly progId: string;
  readonly spid?: string;
  /** `p:embed` or `p:link`. The choice, and the only child before `p:pic`. */
  readonly linked?: boolean;
  readonly pic?: string;
}): string {
  return (
    '<p:oleObj' +
    (spec.spid === undefined ? '' : ` spid="${spec.spid}"`) +
    ` name="${spec.name}" r:id="${spec.relId}"` +
    ` imgW="${String(IMG_W)}" imgH="${String(IMG_H)}" progId="${spec.progId}">` +
    (spec.linked === true ? '<p:link updateAutomatic="1"/>' : '<p:embed/>') +
    (spec.pic ?? '') +
    '</p:oleObj>'
  );
}

const oleFrame = (id: number, name: string, content: string): string =>
  graphicFrame({ id, name, ...FRAME, uri: GRAPHIC_URI.ole, content });

// ------------------------------------------------------------------ the VML

/**
 * `ppt/drawings/vmlDrawing1.vml`.
 *
 * The root element is literally `<xml>`, which is legal and startling, and the
 * part's content type - `application/vnd.openxmlformats-officedocument.vmlDrawing`
 * - does not end in `+xml` even though the part is XML. A reader that decides
 * what to parse from the media type alone skips it.
 *
 * `v:shapetype` is the shared definition and `v:shape/@type` points at it with
 * a fragment reference, `#_x0000_t75`, which is a VML convention and not an
 * OPC one. `o:ole=""` is the empty-string attribute that marks the shape as an
 * OLE site.
 */
const VML_DRAWING =
  `<xml xmlns:v="${NS_V}" xmlns:o="${NS_O}"` +
  ' xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">\r\n' +
  ' <o:shapelayout v:ext="edit">\r\n' +
  '  <o:idmap v:ext="edit" data="1"/>\r\n' +
  ' </o:shapelayout>\r\n' +
  ' <v:shapetype id="_x0000_t75" coordsize="21600,21600" o:spt="75" o:preferrelative="t"\r\n' +
  '  path="m@4@5l@4@11@9@11@9@5xe" filled="f" stroked="f">\r\n' +
  '  <v:stroke joinstyle="miter"/>\r\n' +
  '  <v:formulas>\r\n' +
  '   <v:f eqn="if lineDrawn pixelLineWidth 0"/>\r\n' +
  '   <v:f eqn="sum @0 1 0"/>\r\n' +
  '   <v:f eqn="sum 0 0 @1"/>\r\n' +
  '   <v:f eqn="prod @2 1 2"/>\r\n' +
  '   <v:f eqn="prod @3 21600 pixelWidth"/>\r\n' +
  '   <v:f eqn="prod @3 21600 pixelHeight"/>\r\n' +
  '   <v:f eqn="sum @0 0 1"/>\r\n' +
  '   <v:f eqn="prod @6 1 2"/>\r\n' +
  '   <v:f eqn="prod @7 21600 pixelWidth"/>\r\n' +
  '   <v:f eqn="sum @8 21600 0"/>\r\n' +
  '   <v:f eqn="prod @7 21600 pixelHeight"/>\r\n' +
  '   <v:f eqn="sum @10 21600 0"/>\r\n' +
  '  </v:formulas>\r\n' +
  '  <v:path o:extrusionok="f" gradientshapeok="t" o:connecttype="rect"/>\r\n' +
  '  <o:lock v:ext="edit" aspectratio="t"/>\r\n' +
  ' </v:shapetype>\r\n' +
  ` <v:shape id="${SPID}" type="#_x0000_t75"` +
  ' style=\'position:absolute;left:0;top:0;width:408pt;height:300pt\' o:ole="">\r\n' +
  '  <v:imagedata r:id="rId1" o:title=""/>\r\n' +
  ' </v:shape>\r\n' +
  '</xml>\r\n';

// ------------------------------------------------------------------ package

const PARTS: readonly ProbePart[] = [
  {
    name: 'ppt/embeddings/oleObject1.xlsx',
    bytes: probeXlsx([
      ['PPTX Studio corpus', 'a26-ole'],
      ['This workbook is the OLE payload.', 'Five parts, all authored here.'],
    ]),
    contentType: { kind: 'default', extension: 'xlsx', type: CT_XLSX },
  },
  {
    name: 'ppt/media/image1.emf',
    bytes: probeEmf(),
    contentType: { kind: 'default', extension: 'emf', type: 'image/x-emf' },
  },
  {
    name: 'ppt/drawings/vmlDrawing1.vml',
    bytes: VML_DRAWING,
    contentType: { kind: 'default', extension: 'vml', type: CT_VML },
  },
  {
    // The VML part has relationships of its own, and `rId1` inside it is a
    // different rId1 from every other part's. rIds are scoped per `.rels`.
    name: 'ppt/drawings/_rels/vmlDrawing1.vml.rels',
    bytes: relsXml([{ id: 'rId1', type: REL + 'image', target: '../media/image1.emf' }]),
  },
];

const PACKAGE_REL: ProbeRel = {
  id: 'rId2',
  type: REL + 'package',
  target: '../embeddings/oleObject1.xlsx',
};
const EMF_REL: ProbeRel = { id: 'rId3', type: REL + 'image', target: '../media/image1.emf' };
const VML_REL: ProbeRel = {
  id: 'rId4',
  type: REL + 'vmlDrawing',
  target: '../drawings/vmlDrawing1.vml',
};

const caption = (id: number, lines: readonly string[]): string =>
  shape({
    id,
    name: 'Caption',
    x: 6477000,
    y: 1600200,
    cx: 5029200,
    cy: 3805404,
    textBody: txBody({
      bodyPr: '<a:bodyPr wrap="square" anchor="t"/>',
      paras: lines.map((line) => textLine(line, { sz: 1300 })).join(''),
    }),
  });

export const a26Ole: ProbeDeck = {
  id: 'a26-ole',
  title: 'PPTX Studio corpus: a26 ole',
  description:
    'OLE objects in the shape experiment E6 measured: an mc:AlternateContent whose mc:Choice ' +
    'requires the VML prefix and whose package contains no VML part, whose Choice branch has no ' +
    'preview at all, and whose mc:Fallback carries the only p:pic - with an EMF, not a PNG. The ' +
    'payload is an OPC package under ppt/embeddings, authored here, not a compound file. Slide 2 ' +
    'writes the pre-2205 transitional form with @spid and a real vmlDrawing part whose v:shape ' +
    'carries the same id; slide 3 takes p:oleObj’s other branch, p:link, with an external target ' +
    'and no local bytes. One logical object inside a switch censuses as two, because the census ' +
    'reads raw pre-MCE markup.',
  features: {
    // 6 chassis + one caption a slide.
    shape: 9,
    placeholder: 6,
    gradientFill: 2,
    graphicFrame: 3,
    // Two in the switch on slide 1, one transitional, one linked.
    oleObject: 4,
    alternateContent: 1,
    picture: 3,
    presetGeom: 6,
    vml: 1,
    embeddedPackage: 1,
  },
  build: () => ({
    title: 'PPTX Studio corpus: a26 ole',
    parts: PARTS,
    slides: [
      {
        title: 'a26 — the switch PowerPoint writes, and the preview it hides',
        rels: [PACKAGE_REL, EMF_REL],
        body:
          oleFrame(
            10,
            'Worksheet',
            `<mc:AlternateContent xmlns:mc="${NS_MC}">` +
              `<mc:Choice xmlns:v="${NS_V}" Requires="v">` +
              oleObj({ name: 'Worksheet', relId: 'rId2', progId: 'Excel.Sheet.12' }) +
              '</mc:Choice>' +
              '<mc:Fallback>' +
              oleObj({
                name: 'Worksheet',
                relId: 'rId2',
                progId: 'Excel.Sheet.12',
                pic: preview(11, 'rId3'),
              }) +
              '</mc:Fallback>' +
              '</mc:AlternateContent>',
          ) +
          caption(12, [
            'The Choice requires v and the package has no VML part.',
            'The Choice branch has no p:pic, so resolving MCE the usual way',
            'leaves nothing to draw: the preview is in the branch you discard.',
            'The preview is an EMF. Sub-phase 10.6 is on this path, not beside it.',
          ]),
      },
      {
        title: 'a26 — the transitional form, with @spid and a VML part',
        rels: [PACKAGE_REL, EMF_REL, VML_REL],
        body:
          oleFrame(
            10,
            'Worksheet',
            oleObj({
              spid: SPID,
              name: 'Worksheet',
              relId: 'rId2',
              progId: 'Excel.Sheet.12',
              pic: preview(11, 'rId3'),
            }),
          ) +
          caption(12, [
            'p:oleObj/@spid is _x0000_s1026 and so is v:shape/@id in',
            'ppt/drawings/vmlDrawing1.vml. Nothing but that string joins them.',
            'The VML part has its own .rels, so its rId1 is a different rId1',
            'from the slide’s - rIds are scoped to one relationship part.',
            'Its content type does not end in +xml, though the part is XML.',
          ]),
      },
      {
        title: 'a26 — a linked object, with no bytes in the package',
        rels: [
          {
            id: 'rId2',
            type: REL + 'oleObject',
            target: 'https://example.invalid/corpus/linked-workbook.xlsx',
            external: true,
          },
          EMF_REL,
        ],
        body:
          oleFrame(
            10,
            'Linked worksheet',
            oleObj({
              name: 'Linked worksheet',
              relId: 'rId2',
              progId: 'Excel.Sheet.12',
              linked: true,
              pic: preview(11, 'rId3'),
            }),
          ) +
          caption(12, [
            'p:oleObj holds a choice: p:embed or p:link, never both.',
            'p:link carries @updateAutomatic and the relationship carries',
            'TargetMode="External", so the only local bytes are the preview.',
            'The target is under example.invalid, which RFC 2606 reserves.',
          ]),
      },
    ],
  }),
};
