import { grid, group, prstGeom, scheme, shape, solidFill, type Cell } from '../../markup/shapes.ts';
import { textLine, txBody } from '../../markup/text.ts';
import type { ProbeDeck } from '../../markup/types.ts';

/**
 * Markup Compatibility, in the four shapes a real deck uses it in.
 *
 * ## Why this deck is load-bearing rather than exotic
 *
 * `mc:Choice/@Requires` and `mc:Ignorable` hold **prefixes, not URIs**. That
 * one sentence is architectural bet 2's whole justification for a hand-rolled
 * XML layer: the browser's `DOMParser`/`XMLSerializer` pair is *permitted by
 * spec* to rewrite namespace prefixes, and `@xmldom/xmldom` loses them outright
 * under a prefixed parent. Rewrite `p14` to `ns0` anywhere in a part and a
 * `Requires="p14"` two hundred lines away stops naming a declared prefix -
 * which converts ignorable extension markup into a hard error, silently, in a
 * part that still looks well-formed.
 *
 * Every probe below is therefore also a probe of prefix survival. The
 * declarations are deliberately scattered: some on the `mc:Choice` itself,
 * which is where PowerPoint puts them, and some on the part root, which is
 * where a `mc:Ignorable` has to be.
 *
 * ## The four shapes
 *
 * **1. The switch.** `mc:AlternateContent` holding `mc:Choice+` and an optional
 * `mc:Fallback`. Slide 1 covers the cases that are easy to get wrong: two
 * `mc:Choice` children where the first one that is satisfied wins, a
 * `@Requires` naming two prefixes at once where **both** must be understood, a
 * switch with no `mc:Fallback` at all, and a `@Requires` whose prefix is
 * declared on an ancestor rather than on the `mc:Choice`.
 *
 * **2. Nesting.** A switch inside a `mc:Choice`, and another inside a
 * `mc:Fallback`. Legal, and the reason a consumer cannot resolve MCE with one
 * pass over the top-level children. What is *not* legal, and is not here, is
 * an `mc:AlternateContent` as the direct child of another: it has to be inside
 * a branch.
 *
 * **3. Ignoring.** `mc:Ignorable` on the part root naming a prefix, so elements
 * and attributes in that namespace are dropped rather than refused. This is
 * the mechanism `p14:` attributes ride on in every PowerPoint deck written
 * since 2010, and it is completely separate from the switch.
 *
 * **4. Ignoring, but keeping the children.** `mc:ProcessContent` names an
 * element that is to be removed while its *content* is processed in its place,
 * so a wrapper in an unknown namespace can hold shapes that survive it. And
 * `mc:PreserveElements`/`mc:PreserveAttributes`, which are the editor-facing
 * half: they ask a consumer that is rewriting the file to keep ignorable
 * markup rather than drop it. This project is exactly the consumer those two
 * are addressed to.
 *
 * ## The namespaces
 *
 * `p14` is real and PowerPoint understands it. `zz` and `yy` are
 * `example`-domain URIs nothing implements, which is what makes a `Requires`
 * naming them fall through to the fallback in PowerPoint and lets the deck
 * state which branch should render.
 *
 * `mc:MustUnderstand` is the one member of the family this deck names but does
 * not stress: it is a demand that the consumer refuse the part when it does not
 * understand a namespace, so a probe naming `zz` would be a deck nothing can
 * open, including us. It names `p14` instead, which PowerPoint does understand,
 * so the assertion is real and the file still opens.
 *
 * A fifth case is deliberately absent and is recorded in `ROSTER.md` instead: a
 * `@Requires` naming a prefix that is declared **nowhere**. MCE says that is an
 * error; what PowerPoint does about it was measured separately rather than by
 * making every other probe in this deck depend on the answer.
 *
 * ## What PowerPoint does, measured
 *
 * The shape counts on open are 5, 5, 5 - a title and four shapes on every
 * slide - and slide 3 is the interesting one, because the four are the
 * ignorable-attribute shape, the shape inside `zz:wrapper`, the
 * `mc:PreserveAttributes` shape and the `mc:MustUnderstand` shape. The shape
 * inside `zz:dropped` is **not** there and the one inside `zz:wrapper`
 * **is**, which is `mc:ProcessContent` working: the wrapper is removed and its
 * content processed in its place, while the unnamed wrapper takes its content
 * with it. Four would mean the wrapper was not honoured and six would mean the
 * dropped element leaked; five is the answer, and only five is.
 *
 * Re-saving on 2026-08-28 shows the rest. **`mc:Ignorable` is gone and so is
 * everything it named**, which is correct - once the ignorable markup has been
 * dropped there is nothing left to declare. The nine `mc:AlternateContent`
 * elements come back as six, because PowerPoint resolves the switches it
 * understands and keeps the ones it does not. So a consumer that means to
 * preserve a document has to resolve MCE for **rendering** and leave the
 * markup alone for **writing**, which are two different passes over the same
 * tree - and is exactly why sub-phase 0.6's MCE walker resolves rather than
 * rewrites.
 */

const NS_MC = 'http://schemas.openxmlformats.org/markup-compatibility/2006';
const NS_P14 = 'http://schemas.microsoft.com/office/powerpoint/2010/main';
const NS_ZZ = 'http://pptx-studio.example/ns/a37';
const NS_YY = 'http://pptx-studio.example/ns/a37/second';

/** The part root, with every prefix the ignoring probes need in scope. */
const IGNORING_ROOT =
  ` xmlns:mc="${NS_MC}"` +
  ` xmlns:zz="${NS_ZZ}"` +
  ` xmlns:yy="${NS_YY}"` +
  ` xmlns:p14="${NS_P14}"` +
  ' mc:Ignorable="zz yy"' +
  ' mc:MustUnderstand="p14"' +
  ' mc:ProcessContent="zz:wrapper"' +
  ' mc:PreserveElements="zz:keepMe"' +
  ' mc:PreserveAttributes="zz:keepThis"';

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

/** `mc:AlternateContent`, declaring `mc` on itself the way PowerPoint does. */
function alternate(branches: string): string {
  return `<mc:AlternateContent xmlns:mc="${NS_MC}">${branches}</mc:AlternateContent>`;
}

/** One `mc:Choice`. `declare` is the namespace declaration `@Requires` needs. */
function choice(requires: string, declare: string, content: string): string {
  return `<mc:Choice${declare} Requires="${requires}">${content}</mc:Choice>`;
}

const fallback = (content: string): string => `<mc:Fallback>${content}</mc:Fallback>`;

const declareZz = ` xmlns:zz="${NS_ZZ}"`;
const declareYy = ` xmlns:yy="${NS_YY}"`;
const declareP14 = ` xmlns:p14="${NS_P14}"`;

export const a37Mce: ProbeDeck = {
  id: 'a37-mce',
  title: 'PPTX Studio corpus: a37 MCE',
  description:
    'Markup Compatibility in its four shapes: the mc:AlternateContent switch with several ' +
    'mc:Choice children, a @Requires naming two prefixes at once, a switch with no mc:Fallback and ' +
    'one whose prefix is declared on an ancestor; switches nested inside both a Choice and a ' +
    'Fallback; mc:Ignorable on the part root with ignorable elements and attributes inline; and ' +
    'mc:ProcessContent, mc:PreserveElements, mc:PreserveAttributes and mc:MustUnderstand. Every ' +
    'one of these holds a prefix rather than a URI, which is why a serializer permitted to rename ' +
    'prefixes turns ignorable markup into a hard error somewhere else in the part.',
  features: {
    shape: 28,
    placeholder: 6,
    presetGeom: 22,
    gradientFill: 2,
    group: 1,
    alternateContent: 9,
  },
  build: () => ({
    title: 'PPTX Studio corpus: a37 MCE',
    slides: [
      {
        title: 'a37 — the switch, four ways',
        // `zz` is declared here so the fourth probe's `mc:Choice` can require a
        // prefix it does not declare itself. That is legal, is what a real
        // document does when several switches share a namespace, and is the
        // case that breaks first when prefixes are rewritten.
        rootAttributes: declareZz,
        body:
          // 1. Two choices. `zz` is understood by nobody, `p14` by PowerPoint,
          //    so the second one renders and the fallback does not.
          alternate(
            choice(
              'zz',
              declareZz,
              box(10, 'Choice 1 — requires zz', cell(0), 'accent1', [
                'choice 1: Requires="zz"',
                'nothing understands zz',
                'you should not see this',
              ]),
            ) +
              choice(
                'p14',
                declareP14,
                box(11, 'Choice 2 — requires p14', cell(0), 'accent2', [
                  'choice 2: Requires="p14"',
                  'PowerPoint understands p14',
                  'this is the one that renders',
                ]),
              ) +
              fallback(
                box(12, 'Fallback', cell(0), 'accent3', [
                  'the fallback',
                  'reached only if no choice is',
                ]),
              ),
          ) +
          // 2. `@Requires` naming two prefixes. Both must be understood.
          alternate(
            choice(
              'p14 zz',
              declareP14 + declareZz,
              box(13, 'Requires both', cell(1), 'accent4', [
                'Requires="p14 zz"',
                'both, not either',
                'zz is unknown, so this loses',
              ]),
            ) +
              fallback(
                box(14, 'Fallback for both', cell(1), 'accent5', [
                  'the fallback renders',
                  'because zz fails the pair',
                ]),
              ),
          ) +
          // 3. No fallback at all. `mc:Fallback` is [0..1].
          alternate(
            choice(
              'p14',
              declareP14,
              box(15, 'No fallback', cell(2), 'accent6', [
                'a switch with no mc:Fallback',
                'legal: Fallback is optional',
                'and this choice is satisfied',
              ]),
            ),
          ) +
          // 4. The prefix is declared on the slide root, not on the choice.
          alternate(
            choice(
              'zz',
              '',
              box(16, 'Ancestor-declared prefix', cell(3), 'accent1', [
                'Requires="zz", declared on p:sld',
                'not on the mc:Choice',
                'rename the prefix and this breaks',
              ]),
            ) +
              fallback(
                box(17, 'Ancestor fallback', cell(3), 'accent2', [
                  'this renders',
                  'zz is declared, not understood',
                ]),
              ),
          ),
      },
      {
        title: 'a37 — switches inside switches',
        body:
          // A switch inside a Choice. The outer choice is satisfied, so the
          // inner switch is what gets resolved.
          alternate(
            choice(
              'p14',
              declareP14,
              alternate(
                choice(
                  'yy',
                  declareYy,
                  box(10, 'Inner choice', cell(0), 'accent3', [
                    'outer Choice satisfied',
                    'inner Requires="yy" is not',
                  ]),
                ) +
                  fallback(
                    box(11, 'Inner fallback', cell(0), 'accent4', [
                      'nested two deep',
                      'inside a Choice',
                      'this is what renders',
                    ]),
                  ),
              ),
            ) +
              fallback(
                box(12, 'Outer fallback', cell(0), 'accent5', ['unreached', 'p14 is understood']),
              ),
          ) +
          // A switch inside a Fallback. The outer choice fails, so the fallback
          // is taken, and it is itself a switch.
          alternate(
            choice(
              'zz',
              declareZz,
              box(13, 'Outer choice', cell(1), 'accent6', ['unreached', 'zz is not understood']),
            ) +
              fallback(
                alternate(
                  choice(
                    'p14',
                    declareP14,
                    box(14, 'Choice inside a fallback', cell(1), 'accent1', [
                      'nested inside mc:Fallback',
                      'and this inner choice',
                      'is the one that renders',
                    ]),
                  ) + fallback(box(15, 'Innermost fallback', cell(1), 'accent2', ['unreached'])),
                ),
              ),
          ) +
          // A switch inside a group's own spTree, which is a different position
          // in the schema from the slide's.
          group({
            id: 16,
            name: 'Group holding a switch',
            x: cell(2).x,
            y: cell(2).y,
            cx: cell(2).cx,
            cy: cell(2).cy,
            childOffsetX: 0,
            childOffsetY: 0,
            childWidth: cell(2).cx,
            childHeight: cell(2).cy,
            children: alternate(
              choice(
                'zz',
                declareZz,
                box(
                  17,
                  'In-group choice',
                  { x: 0, y: 0, cx: cell(2).cx, cy: cell(2).cy },
                  'accent3',
                  ['unreached'],
                ),
              ) +
                fallback(
                  box(
                    18,
                    'In-group fallback',
                    { x: 0, y: 0, cx: cell(2).cx, cy: cell(2).cy },
                    'accent4',
                    ['a switch inside p:grpSp', 'a different schema position', 'same resolution'],
                  ),
                ),
            ),
          }) +
          box(19, 'Control', cell(3), 'accent5', [
            'no switch at all',
            'the shape a reader compares to',
          ]),
      },
      {
        title: 'a37 — ignoring, and keeping what is ignored',
        rootAttributes: IGNORING_ROOT,
        body:
          // An ignorable attribute on an element we do understand. `mc:Ignorable`
          // names the prefix; the attribute goes away and the element stays.
          '<p:sp><p:nvSpPr>' +
          '<p:cNvPr id="10" name="Ignorable attribute" zz:note="dropped" yy:note="also dropped"/>' +
          '<p:cNvSpPr/><p:nvPr/></p:nvSpPr>' +
          '<p:spPr>' +
          `<a:xfrm><a:off x="${String(cell(0).x)}" y="${String(cell(0).y)}"/>` +
          `<a:ext cx="${String(cell(0).cx)}" cy="${String(cell(0).cy)}"/></a:xfrm>` +
          prstGeom('roundRect') +
          solidFill(scheme('accent6', '<a:lumMod val="60000"/><a:lumOff val="40000"/>')) +
          '</p:spPr>' +
          txBody({
            bodyPr: '<a:bodyPr wrap="square" anchor="ctr"/>',
            paras:
              textLine('ignorable attributes', { sz: 1100, b: true }, { algn: 'ctr' }) +
              textLine('zz:note and yy:note on p:cNvPr', { sz: 1100 }, { algn: 'ctr' }) +
              textLine('the shape survives, the attributes do not', { sz: 1100 }, { algn: 'ctr' }),
          }) +
          '</p:sp>' +
          // An ignorable element, whole. Removed with its content.
          '<zz:dropped>' +
          box(11, 'Never seen', cell(1), 'accent1', [
            'inside zz:dropped',
            'removed with its parent',
          ]) +
          '</zz:dropped>' +
          // The same, but named by `mc:ProcessContent`: the wrapper goes and its
          // children are processed in its place.
          '<zz:wrapper>' +
          box(12, 'Survives its wrapper', cell(1), 'accent2', [
            'inside zz:wrapper',
            'named by mc:ProcessContent',
            'so the wrapper goes and this stays',
          ]) +
          '</zz:wrapper>' +
          // `mc:PreserveElements` and `mc:PreserveAttributes` are addressed at
          // an editor rather than a viewer: keep this rather than dropping it.
          '<zz:keepMe/>' +
          '<p:sp><p:nvSpPr>' +
          '<p:cNvPr id="13" name="Preserve me" zz:keepThis="an attribute an editor must not drop"/>' +
          '<p:cNvSpPr/><p:nvPr/></p:nvSpPr>' +
          '<p:spPr>' +
          `<a:xfrm><a:off x="${String(cell(2).x)}" y="${String(cell(2).y)}"/>` +
          `<a:ext cx="${String(cell(2).cx)}" cy="${String(cell(2).cy)}"/></a:xfrm>` +
          prstGeom('roundRect') +
          solidFill(scheme('accent3', '<a:lumMod val="60000"/><a:lumOff val="40000"/>')) +
          '</p:spPr>' +
          txBody({
            bodyPr: '<a:bodyPr wrap="square" anchor="ctr"/>',
            paras:
              textLine('mc:PreserveAttributes', { sz: 1100, b: true }, { algn: 'ctr' }) +
              textLine('ignorable, and an editor', { sz: 1100 }, { algn: 'ctr' }) +
              textLine('is asked to keep it anyway', { sz: 1100 }, { algn: 'ctr' }),
          }) +
          '</p:sp>' +
          box(14, 'MustUnderstand', cell(3), 'accent4', [
            'mc:MustUnderstand="p14"',
            'refuse the part if p14 is unknown',
            'PowerPoint knows it, so this opens',
          ]),
      },
    ],
  }),
};
