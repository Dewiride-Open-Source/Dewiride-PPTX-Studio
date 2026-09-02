import { grid, scheme, shape, solidFill } from '../shapes.ts';
import { textLine, txBody } from '../text.ts';
import type { ProbeDeck } from '../types.ts';

/**
 * `p:timing`: the deepest tree in PresentationML, and the one this project
 * promises to preserve and never to play.
 *
 * The census's `animation` rule counts `p:timing` elements and the feature
 * table marks the phase `preserve`. That promise is only worth something if
 * there is a fixture that would notice it being broken, and animation markup is
 * exactly the shape that gets broken by accident: a nine-level nest of
 * `p:par`/`p:cTn`/`p:childTnLst` inside a slide part that Phase 5 edits every
 * time somebody drags a shape. Sub-phase 0.5's per-node byte fidelity exists
 * for this; here is the tree it has to survive.
 *
 * ## The shape of the canonical nest
 *
 * One click-triggered entrance effect is nine elements deep before any
 * behaviour appears:
 *
 * ```
 * timing > tnLst > par > cTn(tmRoot) > childTnLst > seq(mainSeq)
 *        > cTn > childTnLst > par > cTn > childTnLst > par > cTn
 *        > childTnLst > par > cTn(clickEffect) > childTnLst > set
 * ```
 *
 * The three unlabelled `par` levels are not redundancy. They are, from the
 * outside in, the click group, the paragraph group and the effect group, and
 * PowerPoint's UI reads them back as one row in the Animation Pane. Collapsing
 * them on a round trip loses the grouping without losing the effect, which is
 * the kind of damage that shows up two edits later.
 *
 * ## `@spid` is a `p:cNvPr/@id`, and that is a constraint on renumbering
 *
 * `p:spTgt/@spid` and `p:bldP/@spid` name shapes by their `p:cNvPr/@id`.
 * `p:cNvPr/@id` is unique within a part and may repeat across parts, so
 * duplicating a slide has to renumber them - and every `@spid` in that slide's
 * `p:timing`, plus every `@spid` in a VML drawing, has to move in step. Nothing
 * validates the link: a stale `@spid` is an animation that silently never runs.
 *
 * ## What each slide is for
 *
 * 1. The canonical click-to-fade nest PowerPoint writes, with `p:bldLst`.
 * 2. The five behaviour elements that are not `p:set` - `p:animClr`,
 *    `p:animMotion`, `p:animRot`, `p:animScale` and generic `p:anim` with a
 *    `p:tavLst` of key times.
 * 3. The control constructs: `p:excl`, `p:cmd`, `p:iterate` walking a body by
 *    percentage, `p:subTnLst`, and a `p:bldP` carrying a `p:tmplLst`.
 *
 * Between them that is every child of `p:tnLst` except `p:audio` and
 * `p:video`, which need a media part and belong to `a24-media`.
 */

/** `p:cond`, whose content is a choice of `rtn`, `tgtEl` or `tn` - or nothing. */
const cond = (attributes: string, child = ''): string =>
  child === '' ? `<p:cond ${attributes}/>` : `<p:cond ${attributes}>${child}</p:cond>`;

const spTgt = (spid: number): string => `<p:tgtEl><p:spTgt spid="${String(spid)}"/></p:tgtEl>`;

const attrName = (name: string): string =>
  `<p:attrNameLst><p:attrName>${name}</p:attrName></p:attrNameLst>`;

/**
 * `p:cBhvr`. Sequence: `cTn, tgtEl, attrNameLst` - the target is not first.
 *
 * This and the three builders below it are exported for `a43-kitchen-sink`,
 * which needs a timing tree over a chart, a diagram and an OLE frame. Reusing
 * these rather than writing a second entrance keeps one spelling of the nest in
 * the corpus: markup that has been read, censused and opened in PowerPoint
 * once is worth more than markup that has been written twice.
 */
export function cBhvr(spec: {
  readonly id: number;
  readonly cTnAttributes?: string;
  readonly stCondLst?: string;
  readonly spid: number;
  readonly attribute?: string;
  readonly bhvrAttributes?: string;
}): string {
  const inner = spec.stCondLst ?? '';
  return (
    `<p:cBhvr${spec.bhvrAttributes === undefined ? '' : ' ' + spec.bhvrAttributes}>` +
    `<p:cTn id="${String(spec.id)}"${spec.cTnAttributes === undefined ? '' : ' ' + spec.cTnAttributes}` +
    (inner === '' ? '/>' : `>${inner}</p:cTn>`) +
    spTgt(spec.spid) +
    (spec.attribute === undefined ? '' : attrName(spec.attribute)) +
    '</p:cBhvr>'
  );
}

/**
 * The four wrapper levels every PowerPoint timing tree opens with, ending in
 * the click-effect `p:par` whose `childTnLst` holds `effects`.
 */
export function mainSequence(effects: string, firstId = 3): string {
  const id = (offset: number): string => String(firstId + offset);
  return (
    '<p:timing><p:tnLst>' +
    '<p:par><p:cTn id="1" dur="indefinite" restart="never" nodeType="tmRoot"><p:childTnLst>' +
    '<p:seq concurrent="1" nextAc="seek">' +
    '<p:cTn id="2" dur="indefinite" nodeType="mainSeq"><p:childTnLst>' +
    `<p:par><p:cTn id="${id(0)}" fill="hold">` +
    `<p:stCondLst>${cond('delay="indefinite"')}</p:stCondLst><p:childTnLst>` +
    `<p:par><p:cTn id="${id(1)}" fill="hold">` +
    `<p:stCondLst>${cond('delay="0"')}</p:stCondLst><p:childTnLst>` +
    effects +
    '</p:childTnLst></p:cTn></p:par>' +
    '</p:childTnLst></p:cTn></p:par>' +
    '</p:childTnLst></p:cTn>' +
    // `CT_TLTimeNodeSequence` is `cTn, prevCondLst, nextCondLst`, so these two
    // are siblings of the node above rather than children of it.
    `<p:prevCondLst>${cond('evt="onPrev" delay="0"', '<p:tgtEl><p:sldTgt/></p:tgtEl>')}</p:prevCondLst>` +
    `<p:nextCondLst>${cond('evt="onNext" delay="0"', '<p:tgtEl><p:sldTgt/></p:tgtEl>')}</p:nextCondLst>` +
    '</p:seq>' +
    '</p:childTnLst></p:cTn></p:par>' +
    '</p:tnLst>'
  );
}

/** One click-effect group: a `p:par` whose `cTn` carries the preset triple. */
export function clickEffect(spec: {
  readonly id: number;
  readonly presetID: number;
  readonly presetClass: string;
  readonly presetSubtype: number;
  readonly body: string;
}): string {
  return (
    `<p:par><p:cTn id="${String(spec.id)}" presetID="${String(spec.presetID)}"` +
    ` presetClass="${spec.presetClass}" presetSubtype="${String(spec.presetSubtype)}"` +
    ' fill="hold" nodeType="clickEffect">' +
    `<p:stCondLst>${cond('delay="0"')}</p:stCondLst>` +
    `<p:childTnLst>${spec.body}</p:childTnLst>` +
    '</p:cTn></p:par>'
  );
}

/** `p:set` making a shape visible - the first behaviour of every entrance. */
export function makeVisible(id: number, spid: number): string {
  return (
    '<p:set>' +
    cBhvr({
      id,
      cTnAttributes: 'dur="1" fill="hold"',
      stCondLst: `<p:stCondLst>${cond('delay="0"')}</p:stCondLst>`,
      spid,
      attribute: 'style.visibility',
    }) +
    '<p:to><p:strVal val="visible"/></p:to>' +
    '</p:set>'
  );
}

const SLIDE_ONE_TIMING =
  mainSequence(
    clickEffect({
      id: 5,
      presetID: 10,
      presetClass: 'entr',
      presetSubtype: 0,
      body:
        makeVisible(6, 10) +
        '<p:animEffect transition="in" filter="fade">' +
        cBhvr({ id: 7, cTnAttributes: 'dur="500"', spid: 10 }) +
        '</p:animEffect>',
    }),
  ) +
  // `p:bldLst` is the second child of `p:timing`, and it is what the UI reads
  // to decide whether a text body animates as one object or by paragraph.
  '<p:bldLst><p:bldP spid="10" grpId="0"/></p:bldLst>' +
  '</p:timing>';

const SLIDE_TWO_TIMING =
  mainSequence(
    clickEffect({
      id: 5,
      presetID: 26,
      presetClass: 'emph',
      presetSubtype: 0,
      body:
        // Colour. `p:to` under `p:animClr` holds a DrawingML colour, which is a
        // different content model from `p:to` under `p:set`.
        '<p:animClr clrSpc="rgb">' +
        cBhvr({
          id: 6,
          cTnAttributes: 'dur="1000" fill="hold"',
          spid: 10,
          attribute: 'fillcolor',
        }) +
        '<p:to><a:srgbClr val="C0392B"/></p:to>' +
        '</p:animClr>' +
        // Motion. `@path` is SVG-ish and in fractions of the slide, not EMU.
        '<p:animMotion origin="layout" path="M 0 0 L 0.1 0.05 E" pathEditMode="relative">' +
        cBhvr({ id: 7, cTnAttributes: 'dur="2000" fill="hold"', spid: 11, attribute: 'ppt_x' }) +
        '<p:rCtr x="5000" y="2500"/>' +
        '</p:animMotion>' +
        // Rotation, in 60000ths of a degree: 21600000 is one full turn.
        '<p:animRot by="21600000">' +
        cBhvr({ id: 8, cTnAttributes: 'dur="2000" fill="hold"', spid: 11, attribute: 'r' }) +
        '</p:animRot>' +
        // Scale. `p:by` is a `CT_TLPoint`, percentages in thousandths.
        '<p:animScale>' +
        cBhvr({ id: 9, cTnAttributes: 'dur="1000" fill="hold"', spid: 12, attribute: 'ppt_w' }) +
        '<p:by x="150000" y="150000"/>' +
        '</p:animScale>' +
        // The generic one, with explicit key times. `@tm` is a percentage of
        // the duration in thousandths, so 100000 is the end.
        '<p:anim calcmode="lin" valueType="num">' +
        cBhvr({
          id: 10,
          cTnAttributes: 'dur="1000" fill="hold"',
          spid: 12,
          attribute: 'ppt_y',
          bhvrAttributes: 'additive="base"',
        }) +
        '<p:tavLst>' +
        '<p:tav tm="0"><p:val><p:fltVal val="0.5"/></p:val></p:tav>' +
        '<p:tav tm="100000"><p:val><p:fltVal val="0.75"/></p:val></p:tav>' +
        '</p:tavLst>' +
        '</p:anim>',
    }),
  ) + '</p:timing>';

const SLIDE_THREE_TIMING =
  mainSequence(
    clickEffect({
      id: 5,
      presetID: 1,
      presetClass: 'entr',
      presetSubtype: 0,
      body:
        // `p:excl` - only one child of this node may run at a time.
        '<p:excl>' +
        '<p:cTn id="6" fill="hold">' +
        // `CT_TLCommonTimeNodeData` is stCondLst, endCondLst, endSync,
        // iterate, childTnLst, subTnLst - so `p:iterate` comes before the
        // children it iterates over, and `p:subTnLst` after them.
        `<p:stCondLst>${cond('delay="0"')}</p:stCondLst>` +
        '<p:iterate type="lt"><p:tmPct val="10000"/></p:iterate>' +
        '<p:childTnLst>' +
        makeVisible(7, 10) +
        '</p:childTnLst>' +
        '<p:subTnLst>' +
        '<p:animRot by="5400000">' +
        cBhvr({ id: 8, cTnAttributes: 'dur="500" fill="hold"', spid: 11, attribute: 'r' }) +
        '</p:animRot>' +
        '</p:subTnLst>' +
        '</p:cTn>' +
        '</p:excl>' +
        // `p:cmd` - a command sent to the target rather than a property to
        // interpolate. This is how media playback is driven.
        '<p:cmd type="evt" cmd="onstopaudio">' +
        cBhvr({ id: 9, cTnAttributes: 'dur="1"', spid: 12 }) +
        '</p:cmd>',
    }),
  ) +
  // A build with a template: paragraph-level animation, each level animating
  // with the timeline the template names.
  '<p:bldLst>' +
  '<p:bldP spid="10" grpId="0" build="p" bldLvl="2" animBg="1">' +
  '<p:tmplLst><p:tmpl lvl="1"><p:tnLst>' +
  '<p:par><p:cTn id="20" presetID="1" presetClass="entr" fill="hold" nodeType="withEffect">' +
  `<p:stCondLst>${cond('delay="0"')}</p:stCondLst>` +
  `<p:childTnLst>${makeVisible(21, 10)}</p:childTnLst>` +
  '</p:cTn></p:par>' +
  '</p:tnLst></p:tmpl></p:tmplLst>' +
  '</p:bldP>' +
  '</p:bldLst>' +
  '</p:timing>';

/** Three shapes with the ids every `@spid` above names: 10, 11 and 12. */
function targets(labels: readonly [string, string, string]): string {
  const cell = grid(3, 1);
  return labels
    .map((label, index) =>
      shape({
        id: 10 + index,
        name: label,
        ...cell(index),
        fill: solidFill(scheme(`accent${String(index + 1)}`)),
        textBody: txBody({
          bodyPr:
            '<a:bodyPr wrap="square" lIns="91440" tIns="45720" rIns="91440" bIns="45720"' +
            ' anchor="ctr"/>',
          paras: textLine(label, { sz: 1400 }, { algn: 'ctr' }),
        }),
      }),
    )
    .join('');
}

export const a17Animations: ProbeDeck = {
  id: 'a17-animations',
  title: 'PPTX Studio corpus: a17 animations',
  description:
    'Three p:timing trees: the nine-level click-to-fade nest PowerPoint writes with its p:bldLst, ' +
    'the five behaviour elements that are not p:set with their differing content models, and the ' +
    'control constructs - p:excl, p:cmd, p:iterate, p:subTnLst and a p:bldP with a p:tmplLst. ' +
    'Every p:spTgt/@spid names a p:cNvPr/@id that exists on the slide, which is the link nothing ' +
    'validates and slide duplication has to move in step. Preserved across an export, never played.',
  features: {
    shape: 15,
    placeholder: 6,
    presetGeom: 9,
    gradientFill: 2,
    animation: 3,
  },
  build: () => ({
    title: 'PPTX Studio corpus: a17 animations',
    slides: [
      {
        title: 'a17 — the canonical click-to-fade nest',
        layout: 0,
        body: targets(['spid 10 — fades in', 'spid 11', 'spid 12']),
        tail: SLIDE_ONE_TIMING,
      },
      {
        title: 'a17 — every behaviour element except audio and video',
        layout: 0,
        body: targets([
          'spid 10 — animClr',
          'spid 11 — animMotion, animRot',
          'spid 12 — animScale, anim',
        ]),
        tail: SLIDE_TWO_TIMING,
      },
      {
        title: 'a17 — excl, cmd, iterate, subTnLst and a build template',
        layout: 0,
        body: targets([
          'spid 10 — build target',
          'spid 11 — subTnLst target',
          'spid 12 — cmd target',
        ]),
        tail: SLIDE_THREE_TIMING,
      },
    ],
  }),
};
