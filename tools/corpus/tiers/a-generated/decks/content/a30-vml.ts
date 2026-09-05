import { relsXml, REL, type ProbePart, type ProbeRel } from '../../markup/chassis.ts';
import { probeEmf } from '../../assets/emf.ts';
import { graphicFrame, GRAPHIC_URI, picture, shape } from '../../markup/shapes.ts';
import { textLine, txBody } from '../../markup/text.ts';
import type { ProbeDeck } from '../../markup/types.ts';

/**
 * VML: the 1998 drawing language that is still load-bearing in 2026.
 *
 * `a26-ole` is about the OLE object - the switch PowerPoint writes, the preview
 * hidden in the branch a resolver discards, the package payload. This deck
 * borrows its one working entry point and is about the **VML itself**: what the
 * language can express, where its geometry lives, and why the answer is to
 * rasterise it rather than implement it.
 *
 * ## `p:control` is a whole-package refusal, which is why this deck is shaped
 * this way
 *
 * VML has exactly two hosts in PresentationML: `p:oleObj/@spid` and
 * `p:control/@spid`. The second is the more interesting one, because
 * **`p:control` is not in `p:spTree`** - it lives in `p:cSld/p:controls`,
 * after the shape tree - so a slide can carry visible, clickable content that a
 * shape walk never reports.
 *
 * PowerPoint 16.0.20326 refuses every package containing one. Eight variants
 * were built and opened, changing one thing at a time:
 *
 * | variant                                                    | result   |
 * | ---------------------------------------------------------- | -------- |
 * | `<p:controls/>`, empty                                     | opens    |
 * | `<p:control/>`, no attributes at all                       | refused  |
 * | `@name` only                                               | refused  |
 * | `@spid` + `@name`, no `r:id`                               | refused  |
 * | `@r:id` + `@spid`, minimal                                 | refused  |
 * | with `@imgW`/`@imgH` and a `p:pic` preview                 | refused  |
 * | with the ActiveX part `persistStreamInit` and its `.bin`   | refused  |
 * | the same package retyped as a macro-enabled presentation   | refused  |
 *
 * So it is `p:control` itself, not its attributes, not its target, not the
 * ActiveX part's persistence, and not the absence of a VBA project. An empty
 * `p:controls` is fine, which places the refusal precisely. The corpus records
 * it in `ROSTER.md` under **What PowerPoint refuses** and it becomes a
 * `corpus/reject/` fixture; what real control-bearing decks do differently is
 * an open question, and one a Tier B deck cannot answer either, because
 * PowerPoint has no COM entry point that inserts a control.
 *
 * ## The part is XML and its content type does not say so
 *
 * `application/vnd.openxmlformats-officedocument.vmlDrawing`. No `+xml`
 * suffix, though the part is XML - `packages/census` keeps an explicit list of
 * media types that are XML anyway, and this is the first entry in it. A reader
 * that decides what to parse from the suffix skips every VML part in every
 * package ever written. The root element is literally `<xml>`, in no namespace.
 *
 * ## The join is a string and nothing else
 *
 * `p:oleObj/@spid` is `_x0000_s1026`; `v:shape/@id` in the VML part is
 * `_x0000_s1026`. Nothing types that relationship, nothing validates it, and
 * the two live in different parts reached by different relationships. Sub-phase
 * 10.9 keeps them in sync when `p:cNvPr` ids are renumbered, and the reason
 * that is a rule rather than an afterthought is that renumbering is otherwise
 * completely safe.
 *
 * ## What the VML actually says
 *
 * `v:shapetype` is a shared definition and `v:shape/@type` references it with a
 * **fragment**, `#_x0000_t75` - a VML convention, not an OPC one, and not a
 * relationship.
 *
 * `v:group` nests, and inside a group the children are positioned in the
 * group's `@coordsize` space rather than in points. That is the same trap
 * `a:chOff`/`a:chExt` sets in DrawingML, spelled entirely differently, and
 * solved by neither reading the other.
 *
 * Geometry is a **CSS declaration list in an attribute**:
 * `style='position:absolute;left:0;top:0;width:408pt;height:300pt'`. Not
 * attributes, not EMU - CSS 2, in points, with `mso-` extensions mixed in.
 * Everything a DrawingML renderer knows about geometry has to be re-learned,
 * which is the strongest argument for treating VML as a picture of something
 * else rather than as a drawing language to implement.
 *
 * Only `v:shape` counts as `vml` in the census. `v:group`, `v:rect`, `v:oval`
 * and `v:line` have local names of their own, so four of the eight drawing
 * elements below are `v:shape` and four are not.
 */

const NS_V = 'urn:schemas-microsoft-com:vml';
const NS_O = 'urn:schemas-microsoft-com:office:office';
const NS_R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';

const CT_VML = 'application/vnd.openxmlformats-officedocument.vmlDrawing';

const REL_VML = REL + 'vmlDrawing';

const SPID = ['_x0000_s1026', '_x0000_s1027'] as const;

const FRAME = { x: 685800, y: 1600200, cx: 4572000, cy: 2286000 } as const;

// -------------------------------------------------------------------- VML

const SHAPETYPE_75 =
  ' <v:shapetype id="_x0000_t75" coordsize="21600,21600" o:spt="75" o:preferrelative="t"\r\n' +
  '  path="m@4@5l@4@11@9@11@9@5xe" filled="f" stroked="f">\r\n' +
  '  <v:stroke joinstyle="miter"/>\r\n' +
  '  <v:path o:extrusionok="f" gradientshapeok="t" o:connecttype="rect"/>\r\n' +
  '  <o:lock v:ext="edit" aspectratio="t"/>\r\n' +
  ' </v:shapetype>\r\n';

const vmlDocument = (body: string, withImage = false): string =>
  `<xml xmlns:v="${NS_V}" xmlns:o="${NS_O}"` +
  (withImage ? ` xmlns:r="${NS_R}"` : '') +
  '>\r\n' +
  ' <o:shapelayout v:ext="edit">\r\n' +
  '  <o:idmap v:ext="edit" data="1"/>\r\n' +
  ' </o:shapelayout>\r\n' +
  body +
  '</xml>\r\n';

/** An OLE site: the shape a `p:oleObj/@spid` points at, holding the preview. */
const VML_1 = vmlDocument(
  SHAPETYPE_75 +
    ` <v:shape id="${SPID[0]}" type="#_x0000_t75"\r\n` +
    '  style=\'position:absolute;left:0;top:0;width:360pt;height:180pt\' o:preferrelative="t"\r\n' +
    '  filled="f" stroked="f" o:ole="">\r\n' +
    '  <v:fill o:detectmouseclick="t"/>\r\n' +
    '  <v:imagedata r:id="rId1" o:title=""/>\r\n' +
    '  <o:lock v:ext="edit" rotation="t"/>\r\n' +
    ' </v:shape>\r\n',
  true,
);

/**
 * A group, and everything about VML that is not a shape.
 *
 * `v:group/@coordsize` and `@coordorigin` are the child coordinate space, and
 * the children's `style` is in **that** space rather than in points.
 */
const VML_2 = vmlDocument(
  SHAPETYPE_75 +
    ` <v:shape id="${SPID[1]}" type="#_x0000_t75"\r\n` +
    '  style=\'position:absolute;left:0;top:0;width:360pt;height:180pt\' filled="f" stroked="f"\r\n' +
    '  o:ole=""/>\r\n' +
    ' <v:group id="probe_group"\r\n' +
    "  style='position:absolute;left:0;top:0;width:288pt;height:144pt'\r\n" +
    '  coordsize="21600,10800" coordorigin="0,0">\r\n' +
    '  <v:rect id="probe_rect" style=\'position:absolute;left:0;top:0;width:7200;height:10800\'\r\n' +
    '   fillcolor="#0070c0" strokecolor="#1f3864" strokeweight="2pt">\r\n' +
    '   <v:shadow on="t" color="#7f7f7f" offset="3pt,3pt"/>\r\n' +
    '  </v:rect>\r\n' +
    '  <v:oval id="probe_oval" style=\'position:absolute;left:7920;top:1080;width:5760;height:8640\'\r\n' +
    '   fillcolor="#c00000" stroked="f"/>\r\n' +
    '  <v:line id="probe_line" from="14400,1080" to="21600,9720"\r\n' +
    '   strokecolor="#7030a0" strokeweight="3pt">\r\n' +
    '   <v:stroke dashstyle="dash" endarrow="block"/>\r\n' +
    '  </v:line>\r\n' +
    '  <v:shape id="probe_gradient" type="#_x0000_t75"\r\n' +
    '   style=\'position:absolute;left:0;top:0;width:21600;height:1080\' stroked="f">\r\n' +
    '   <v:fill type="gradient" color="#ffc000" color2="#ed7d31" angle="90"/>\r\n' +
    '  </v:shape>\r\n' +
    ' </v:group>\r\n' +
    ' <v:shape id="probe_textbox" type="#_x0000_t75"\r\n' +
    "  style='position:absolute;left:0;top:150pt;width:288pt;height:36pt'\r\n" +
    '  fillcolor="#f2f2f2" strokecolor="#a6a6a6">\r\n' +
    '  <v:textbox inset="4pt,2pt,4pt,2pt">\r\n' +
    '   <div>Geometry is a CSS string in an attribute, in points.</div>\r\n' +
    '  </v:textbox>\r\n' +
    ' </v:shape>\r\n',
);

// ------------------------------------------------------------------ package

const PARTS: readonly ProbePart[] = [
  {
    name: 'ppt/drawings/vmlDrawing1.vml',
    bytes: VML_1,
    contentType: { kind: 'default', extension: 'vml', type: CT_VML },
  },
  { name: 'ppt/drawings/vmlDrawing2.vml', bytes: VML_2 },
  {
    // The VML part has relationships of its own, and its `rId1` is a different
    // `rId1` from the slide's: rIds are scoped to one relationship part.
    name: 'ppt/drawings/_rels/vmlDrawing1.vml.rels',
    bytes: relsXml([{ id: 'rId1', type: REL + 'image', target: '../media/image1.emf' }]),
  },
  {
    name: 'ppt/media/image1.emf',
    bytes: probeEmf(),
    contentType: { kind: 'default', extension: 'emf', type: 'image/x-emf' },
  },
];

const EMF_REL: ProbeRel = { id: 'rId3', type: REL + 'image', target: '../media/image1.emf' };

/**
 * A linked OLE object whose only purpose is to carry an `@spid`.
 *
 * `p:link` rather than `p:embed`, so nothing is embedded and the deck's subject
 * stays the VML. `a26-ole` is the deck about what is inside an OLE object.
 */
const oleSite = (id: number, name: string, spid: string): string =>
  graphicFrame({
    id,
    name,
    ...FRAME,
    uri: GRAPHIC_URI.ole,
    content:
      `<p:oleObj spid="${spid}" name="${name}" r:id="rId2"` +
      ` imgW="${String(FRAME.cx)}" imgH="${String(FRAME.cy)}" progId="Package">` +
      '<p:link updateAutomatic="0"/>' +
      picture({
        id: id + 1,
        name: `${name} preview`,
        relId: 'rId3',
        ...FRAME,
        description: '',
      }) +
      '</p:oleObj>',
  });

const caption = (id: number, y: number, lines: readonly string[]): string =>
  shape({
    id,
    name: 'Caption',
    x: 5638800,
    y,
    cx: 5867400,
    cy: 3200400,
    textBody: txBody({
      bodyPr: '<a:bodyPr wrap="square" anchor="t"/>',
      paras: lines.map((line) => textLine(line, { sz: 1300 })).join(''),
    }),
  });

/** The OLE target is external, so no bytes and no embedding part. */
const LINK_REL: ProbeRel = {
  id: 'rId2',
  type: REL + 'oleObject',
  target: 'https://example.invalid/corpus/vml-host.bin',
  external: true,
};

export const a30Vml: ProbeDeck = {
  id: 'a30-vml',
  title: 'PPTX Studio corpus: a30 vml',
  description:
    'Two vmlDrawing parts and what VML can express: an OLE site with a v:imagedata pointing at an ' +
    'EMF through relationships of its own, and a v:group whose children are laid out in the ' +
    'group coordinate space rather than in points, with a gradient v:fill, a dashed v:stroke ' +
    'carrying an arrowhead, a v:shadow and a v:textbox holding HTML. Reached through p:oleObj/@spid ' +
    'rather than p:control/@spid, because p:control is a whole-package refusal in all eight forms ' +
    'tested - which this deck documents and ROSTER.md records.',
  features: {
    // 6 chassis + one caption a slide.
    shape: 9,
    placeholder: 6,
    gradientFill: 2,
    graphicFrame: 2,
    oleObject: 2,
    picture: 2,
    presetGeom: 5,
    // Four of the eight VML drawing elements are v:shape: the two OLE sites,
    // the gradient bar and the text box.
    vml: 4,
  },
  build: () => ({
    title: 'PPTX Studio corpus: a30 vml',
    parts: PARTS,
    slides: [
      {
        title: 'a30 — an OLE site, and the VML part it names by string',
        rels: [
          LINK_REL,
          EMF_REL,
          { id: 'rId4', type: REL_VML, target: '../drawings/vmlDrawing1.vml' },
        ],
        body:
          oleSite(10, 'Site', SPID[0]) +
          caption(12, FRAME.y, [
            'p:oleObj/@spid is _x0000_s1026 and so is v:shape/@id in',
            'ppt/drawings/vmlDrawing1.vml. Nothing types that join and nothing',
            'validates it; renumbering p:cNvPr ids must not disturb it.',
            'The VML part has .rels of its own, so its rId1 - the EMF - is a',
            'different rId1 from this slide’s. rIds are scoped per .rels.',
            'Its content type does not end in +xml, though the part is XML.',
          ]),
      },
      {
        title: 'a30 — a group, a gradient, an arrowhead and a text box',
        rels: [
          LINK_REL,
          EMF_REL,
          { id: 'rId4', type: REL_VML, target: '../drawings/vmlDrawing2.vml' },
        ],
        body:
          oleSite(10, 'Group site', SPID[1]) +
          caption(12, FRAME.y, [
            'vmlDrawing2.vml holds v:group, v:rect, v:oval, v:line and three',
            'v:shape of its own. The group declares coordsize="21600,10800" and its',
            'children are positioned in that space rather than in points - the',
            'same problem a:chOff and a:chExt solve in DrawingML, solved another',
            'way, and neither reading helps with the other.',
            'v:textbox holds a div. HTML, inside VML, inside OPC.',
          ]),
      },
      {
        title: 'a30 — why this is a preview to rasterise, not a language to implement',
        body: caption(10, 1600200, [
          'VML predates OOXML by eight years and shares nothing with DrawingML:',
          'a different geometry model, a different colour syntax, a coordinate',
          'space per group, CSS in attributes and HTML inside v:textbox.',
          'It survives in three places - an OLE site, a control site, and Word’s',
          'comment and header anchors - and in all three it holds a picture of',
          'something else.',
          'Two rules follow, both sub-phase 1.2’s: never rewrite a VML part, and',
          'keep @spid stable across every id renumbering.',
          'A third thing this deck found: p:control, the other host, is refused',
          'outright by PowerPoint 16.0.20326 in every form tried.',
        ]),
      },
    ],
  }),
};
