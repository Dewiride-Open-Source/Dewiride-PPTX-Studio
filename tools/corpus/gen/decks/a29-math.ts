import { cNvPrXml } from '../package.ts';
import { prstGeom, scheme, solidFill } from '../shapes.ts';
import { textLine, txBody } from '../text.ts';
import type { ProbeDeck } from '../types.ts';

/**
 * Office Math, which a slide has no schema-legal place to put.
 *
 * `CT_TextParagraph` is `pPr, (br|fld|r)*, endParaRPr`. There is no branch for
 * an equation anywhere in DrawingML text, so **every** `m:oMath` in a
 * PresentationML package is extension markup reached through
 * `mc:AlternateContent` - which is why the census gives `math` a phase of
 * `preserve` rather than a sub-phase number. It is carried across an export
 * byte for byte and never interpreted.
 *
 * Carrying it correctly is not nothing, though. An equation is a whole tree
 * with its own vocabulary, and the one thing an editor must not do is
 * re-serialize the paragraph that holds it from a text model that has no
 * concept of a fraction.
 *
 * ## Two placements, because the wrapper is where this deck is least certain
 *
 * Slide 1 puts the switch at the `p:txBody` position inside a `p:sp`: the
 * Choice holds a whole `p:txBody` whose paragraph contains `a14:m`, and the
 * Fallback holds a whole `p:txBody` of plain runs. Slide 2 puts it one level
 * out, at the `p:spTree` child position, so the Choice and the Fallback are
 * two entire shapes.
 *
 * Both are legal MCE and both occur. This deck writes both **because it could
 * not measure either**: PowerPoint has no COM entry point that inserts an
 * equation - there is no `Shapes.AddEquation`, and `TextRange.Text` cannot
 * hold one - so unlike the rest of the Hard content group this markup is
 * written from ECMA-376 Part 1 §22.1 and from the MCE rules, not read back
 * from a file PowerPoint wrote. `ROSTER.md` records that as a declared gap.
 * What is verified is that PowerPoint opens the package and keeps both shapes.
 *
 * The content of `m:oMath` is a different matter: §22.1 is a complete and
 * precise specification, and the eleven constructs below are transcribed from
 * it rather than guessed.
 *
 * ## `m:oMathPara` versus `m:oMath`
 *
 * `m:oMath` is an inline equation - it sits in a line of text. `m:oMathPara`
 * is a display equation and holds one or more `m:oMath` children plus an
 * `m:oMathParaPr` saying how the group is justified. A reader that treats them
 * as synonyms loses the justification and, in a multi-equation group, the
 * alignment between the lines.
 *
 * ## The attribute namespace is `m` too
 *
 * `<m:jc m:val="centerGroup"/>` - the value attribute is **prefixed**, unlike
 * every `@val` in DrawingML. Office Math inherits WordprocessingML's convention
 * where attributes carry the namespace, and a reader that looks for an
 * unprefixed `val` finds nothing on every property element in the tree.
 */

const NS_M = 'http://schemas.openxmlformats.org/officeDocument/2006/math';
const NS_A14 = 'http://schemas.microsoft.com/office/drawing/2010/main';
const NS_MC = 'http://schemas.openxmlformats.org/markup-compatibility/2006';

// ------------------------------------------------------------------ builders

/** A run of mathematical text. `m:t`, not `a:t`. */
const mr = (text: string): string => `<m:r><m:t>${text}</m:t></m:r>`;

/** An element slot: `m:e`, `m:num`, `m:den`, `m:sub`, `m:sup`, `m:deg`. */
const slot = (tag: string, body: string): string => `<m:${tag}>${body}</m:${tag}>`;

/** `m:f` - a fraction. `m:fPr/m:type` picks bar, skewed, linear or no bar. */
const frac = (numerator: string, denominator: string, type = 'bar'): string =>
  `<m:f><m:fPr><m:type m:val="${type}"/></m:fPr>` +
  slot('num', numerator) +
  slot('den', denominator) +
  '</m:f>';

/** `m:sSup`, `m:sSub`, `m:sSubSup` - the three script forms. */
const sup = (base: string, exponent: string): string =>
  '<m:sSup><m:sSupPr/>' + slot('e', base) + slot('sup', exponent) + '</m:sSup>';

const sub = (base: string, index: string): string =>
  '<m:sSub><m:sSubPr/>' + slot('e', base) + slot('sub', index) + '</m:sSub>';

const subSup = (base: string, index: string, exponent: string): string =>
  '<m:sSubSup><m:sSubSupPr/>' +
  slot('e', base) +
  slot('sub', index) +
  slot('sup', exponent) +
  '</m:sSubSup>';

/** `m:rad`. `m:degHide` is what makes a square root show no index. */
const radical = (body: string, degree?: string): string =>
  '<m:rad><m:radPr>' +
  `<m:degHide m:val="${degree === undefined ? '1' : '0'}"/>` +
  '</m:radPr>' +
  slot('deg', degree ?? '') +
  slot('e', body) +
  '</m:rad>';

/** `m:nary` - a sum, product or integral, with its limits. */
const nary = (character: string, lower: string, upper: string, body: string): string =>
  '<m:nary><m:naryPr>' +
  `<m:chr m:val="${character}"/>` +
  '<m:limLoc m:val="undOvr"/>' +
  '<m:subHide m:val="0"/><m:supHide m:val="0"/>' +
  '</m:naryPr>' +
  slot('sub', lower) +
  slot('sup', upper) +
  slot('e', body) +
  '</m:nary>';

/** `m:d` - a delimiter pair. `m:begChr`/`m:endChr` choose the brackets. */
const delim = (open: string, close: string, body: string): string =>
  `<m:d><m:dPr><m:begChr m:val="${open}"/><m:endChr m:val="${close}"/>` +
  '<m:grow m:val="1"/></m:dPr>' +
  slot('e', body) +
  '</m:d>';

/** `m:m` - a matrix. `m:mcs` declares the columns, `m:mr` is one row. */
const matrix = (rows: ReadonlyArray<readonly string[]>): string => {
  const columns = rows[0]?.length ?? 0;
  return (
    '<m:m><m:mPr><m:mcs>' +
    `<m:mc><m:mcPr><m:count m:val="${String(columns)}"/>` +
    '<m:mcJc m:val="center"/></m:mcPr></m:mc>' +
    '</m:mcs></m:mPr>' +
    rows
      .map((cells) => '<m:mr>' + cells.map((cell) => slot('e', cell)).join('') + '</m:mr>')
      .join('') +
    '</m:m>'
  );
};

/** `m:func` - a named function with an argument, so `sin` is not italicised. */
const func = (name: string, argument: string): string =>
  '<m:func><m:funcPr/>' + slot('fName', mr(name)) + slot('e', argument) + '</m:func>';

/** `m:acc` - an accent above. `m:limLow` - a limit below an operator. */
const accent = (character: string, body: string): string =>
  `<m:acc><m:accPr><m:chr m:val="${character}"/></m:accPr>` + slot('e', body) + '</m:acc>';

const limLow = (base: string, limit: string): string =>
  '<m:limLow><m:limLowPr/>' + slot('e', base) + slot('lim', limit) + '</m:limLow>';

/** `m:bar` - a bar over or under. `m:box` - a grouping with no visible effect. */
const bar = (body: string, position = 'top'): string =>
  `<m:bar><m:barPr><m:pos m:val="${position}"/></m:barPr>` + slot('e', body) + '</m:bar>';

const oMath = (body: string): string => `<m:oMath xmlns:m="${NS_M}">${body}</m:oMath>`;

/** A display equation group, with the justification `m:oMath` alone cannot carry. */
const oMathPara = (equations: readonly string[], justification = 'centerGroup'): string =>
  `<m:oMathPara xmlns:m="${NS_M}">` +
  `<m:oMathParaPr><m:jc m:val="${justification}"/></m:oMathParaPr>` +
  equations.map((body) => `<m:oMath>${body}</m:oMath>`).join('') +
  '</m:oMathPara>';

// ---------------------------------------------------------------- equations

/** Mass-energy, the shortest thing that needs a superscript. */
const EINSTEIN = sup(mr('E = mc'), mr('2'));

/** The quadratic formula: a fraction with a radical in its numerator. */
const QUADRATIC =
  mr('x = ') +
  frac(mr('&#8722;b &#177; ') + radical(sup(mr('b'), mr('2')) + mr(' &#8722; 4ac')), mr('2a'));

/** A sum with both limits, a subscripted term and a delimiter pair. */
const SUM = nary(
  '&#8721;',
  mr('i = 1'),
  mr('n'),
  delim('(', ')', sub(mr('x'), mr('i')) + mr(' &#8722; ') + accent('&#772;', mr('x'))),
);

/** A matrix inside brackets, a named function, and a limit. */
const MIXED =
  delim(
    '[',
    ']',
    matrix([
      [mr('a'), mr('b')],
      [mr('c'), mr('d')],
    ]),
  ) +
  mr(' , ') +
  func('sin', delim('(', ')', mr('&#952;'))) +
  mr(' , ') +
  limLow(mr('lim'), mr('n &#8594; &#8734;')) +
  frac(mr('1'), mr('n'), 'skw');

/** Everything else: a sub-superscript, an under-bar, and a linear fraction. */
const REMAINDER =
  subSup(mr('C'), mr('n'), mr('k')) +
  mr(' , ') +
  bar(mr('AB'), 'bot') +
  mr(' , ') +
  frac(mr('p'), mr('q'), 'lin') +
  mr(' , ') +
  radical(mr('8'), mr('3'));

// ------------------------------------------------------------- the two hosts

const BOX = { x: 685800, y: 1600200, cx: 10820400, cy: 1600200 } as const;

const mathBody = (paragraphs: readonly string[]): string =>
  '<p:txBody><a:bodyPr wrap="square" anchor="ctr"/><a:lstStyle/>' +
  paragraphs
    .map(
      (equation) =>
        `<a:p><a14:m xmlns:a14="${NS_A14}">${equation}</a14:m>` +
        '<a:endParaRPr lang="en-GB" sz="2000" dirty="0"/></a:p>',
    )
    .join('') +
  '</p:txBody>';

const plainBody = (lines: readonly string[]): string =>
  txBody({
    bodyPr: '<a:bodyPr wrap="square" anchor="ctr"/>',
    paras: lines.map((line) => textLine(line, { sz: 2000 })).join(''),
  });

/** The shape around either body. Everything but `p:txBody` is shared. */
const shell = (id: number, name: string, y: number, body: string): string =>
  '<p:sp><p:nvSpPr>' +
  cNvPrXml({ id, name }) +
  '<p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr>' +
  '<p:spPr>' +
  `<a:xfrm><a:off x="${String(BOX.x)}" y="${String(y)}"/>` +
  `<a:ext cx="${String(BOX.cx)}" cy="${String(BOX.cy)}"/></a:xfrm>` +
  prstGeom('rect') +
  solidFill(scheme('bg1')) +
  '</p:spPr>' +
  body +
  '</p:sp>';

/** Slide 1's shape: the switch is the `p:txBody`, and the shell is shared. */
const switchedBody = (
  id: number,
  name: string,
  y: number,
  equations: readonly string[],
  fallback: readonly string[],
): string =>
  '<p:sp><p:nvSpPr>' +
  cNvPrXml({ id, name }) +
  '<p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr>' +
  '<p:spPr>' +
  `<a:xfrm><a:off x="${String(BOX.x)}" y="${String(y)}"/>` +
  `<a:ext cx="${String(BOX.cx)}" cy="${String(BOX.cy)}"/></a:xfrm>` +
  prstGeom('rect') +
  solidFill(scheme('bg1')) +
  '</p:spPr>' +
  `<mc:AlternateContent xmlns:mc="${NS_MC}">` +
  `<mc:Choice xmlns:a14="${NS_A14}" Requires="a14">` +
  mathBody(equations) +
  '</mc:Choice>' +
  '<mc:Fallback>' +
  plainBody(fallback) +
  '</mc:Fallback>' +
  '</mc:AlternateContent>' +
  '</p:sp>';

export const a29Math: ProbeDeck = {
  id: 'a29-math',
  title: 'PPTX Studio corpus: a29 math',
  description:
    'Office Math inside DrawingML text, which has no schema-legal place for it, so every equation ' +
    'here arrives through mc:AlternateContent - at the p:txBody position on slide 1 and at the ' +
    'p:spTree child position on slide 2, because both are legal and neither could be measured. ' +
    'Eleven constructs from ECMA-376 §22.1: fractions in three styles, all three script forms, ' +
    'radicals with and without an index, an n-ary sum with both limits, delimiters, a matrix, a ' +
    'named function, an accent, a limit and a bar. m:oMathPara carries the justification that ' +
    'm:oMath alone cannot, and every property attribute is m:val, prefixed.',
  features: {
    // 6 chassis + two switched shapes on slide 1 and their two fallbacks,
    // two on slide 2 and its one fallback, and one caption a slide.
    shape: 14,
    placeholder: 6,
    gradientFill: 2,
    presetGeom: 8,
    // Two in the first m:oMathPara, one inline, and one in each of slide 2s.
    math: 5,
    alternateContent: 3,
  },
  build: () => ({
    title: 'PPTX Studio corpus: a29 math',
    slides: [
      {
        title: 'a29 — the switch at the p:txBody position',
        body:
          switchedBody(
            10,
            'Display equation',
            1600200,
            [oMathPara([EINSTEIN, QUADRATIC])],
            ['E = mc^2', 'x = (-b +/- sqrt(b^2 - 4ac)) / 2a'],
          ) +
          switchedBody(
            11,
            'Inline equation',
            3352800,
            [oMath(SUM)],
            ['sum from i=1 to n of (x_i - xbar)'],
          ) +
          shell(
            12,
            'Caption',
            4953000,
            plainBody([
              'The Choice and the Fallback are two whole p:txBody elements.',
              'm:oMathPara holds the justification; m:oMath alone cannot.',
            ]),
          ),
      },
      {
        title: 'a29 — the switch at the p:spTree position',
        body:
          `<mc:AlternateContent xmlns:mc="${NS_MC}">` +
          `<mc:Choice xmlns:a14="${NS_A14}" Requires="a14">` +
          shell(10, 'Matrices, functions and limits', 1600200, mathBody([oMathPara([MIXED])])) +
          shell(11, 'Scripts, bars and radicals', 3352800, mathBody([oMathPara([REMAINDER])])) +
          '</mc:Choice>' +
          '<mc:Fallback>' +
          shell(
            10,
            'Equations, as text',
            1600200,
            plainBody([
              '[a b; c d] , sin(theta) , lim n->inf 1/n',
              'C_n^k , AB-under-bar , p/q , cuberoot 8',
            ]),
          ) +
          '</mc:Fallback>' +
          '</mc:AlternateContent>' +
          shell(
            12,
            'Caption',
            4953000,
            plainBody([
              'Here the switch replaces two whole shapes with one.',
              'A Fallback need not have the same shape count as its Choice.',
            ]),
          ),
      },
      {
        title: 'a29 — what a text model must not flatten',
        body: shell(
          10,
          'Caption',
          1600200,
          plainBody([
            'An equation is a tree, and sub-phase 6.1 round-trips paragraphs',
            'through a ProseMirror model that has no node for a fraction.',
            'So a paragraph holding a14:m is never re-serialized from the model:',
            'it is carried across as the bytes it arrived as. The census gives',
            'math a phase of "preserve" for exactly this reason.',
          ]),
        ),
      },
    ],
  }),
};
