import { grid, type Cell } from '../shapes.ts';
import { textLine, txBody } from '../text.ts';
import type { ProbeDeck } from '../types.ts';

/**
 * The lexical forms every producer in this repository declines to emit.
 *
 * ## What this deck is for
 *
 * Sub-phase 0.5's gate is "parse, serialize, byte-identical for 100% of parts",
 * and every fixture it runs against comes out of one of this repository's three
 * PresentationML builders. A convention none of them emits is a convention the
 * gate never tests. `C-LEX` is the rule that says so out loud, and this deck is
 * the fixture that gives it something to check.
 *
 * `conventions.test.ts` asserts the opposite of this deck for our own output -
 * no single-quoted attributes, one declaration form, no stray whitespace - and
 * it scans one probe deck rather than all of them. This is the deck that would
 * fail it, and the two are not in conflict: that assertion is about what we
 * **write**, and this is a fixture about what we must **read** and give back
 * unchanged. If anyone ever widens that test to the whole roster, this deck is
 * the exemption it needs.
 *
 * ## The correction that made it a real probe rather than a hypothetical
 *
 * ADR 0005 counted 70,822 of 98,777 self-closing tags in the real-world corpus
 * written spaced, `<a:off ... />`, against 0 of 1,305 from PowerPoint
 * 16.0.20326, and concluded that no producer available to this project emits
 * the dominant form.
 *
 * That conclusion was wrong, and `a22-chartex` is what found it. In the deck
 * PowerPoint wrote on 2026-08-27, `ppt/charts/chartEx1.xml` writes **12 of its
 * 12** self-closing tags spaced while the other 3,046 across the same sixty
 * parts write none. PowerPoint has more than one XML serializer and the ChartEx
 * one is the odd member. The form is reachable from a producer, in a part, on
 * demand - so `C-LEX` has a second producer for it rather than none.
 *
 * ## Mixed within one part, on purpose
 *
 * The obvious way to build this deck would be one part written entirely
 * spaced, which is what `chartEx1.xml` is. This one mixes the forms inside a
 * single `ppt/slides/slideN.xml`, sibling by sibling, and that is the stronger
 * fixture for the architecture we actually chose.
 *
 * Architectural bet 2 is that byte fidelity is **per-node, not per-part**: a
 * clean `XNode` re-emits by slicing the original buffer, and mutating one node
 * clears `raw` on that node and its ancestors only. A part that is uniformly
 * spaced cannot tell a per-node serializer from a per-part one - both re-emit
 * it correctly. A part where `<a:off .../>` and `<a:off ... />` are siblings
 * can, because any implementation that decided a spacing policy for the part
 * and re-serialized from the model gets exactly half of them wrong. Slide 3 is
 * nothing but that: four shapes alternating between the two forms.
 *
 * Every shape is also internally mixed, because its caption goes through the
 * chassis's `txBody` and comes out tight while its geometry is hand-written.
 *
 * ## The forms, and why each is legal
 *
 * From the XML 1.0 productions, not from habit:
 *
 * | form              | production                                          |
 * | ----------------- | --------------------------------------------------- |
 * | `<a:off … />`     | `EmptyElemTag ::= '<' Name (S Attribute)* S? '/>'`  |
 * | `<a:off …⇥/>`     | the same `S?`, spelled with a tab                   |
 * | `<a:off …␣␣/>`    | the same `S?`; `S` is one **or more**               |
 * | `<p:spPr >`       | `STag ::= '<' Name (S Attribute)* S? '>'`           |
 * | `</a:xfrm >`      | `ETag ::= '</' Name S? '>'`                         |
 * | `val = "1"`       | `Eq ::= S? '=' S?`                                  |
 * | `val='1'`         | `AttValue ::= '"' … '"' \| "'" … "'"`               |
 *
 * Every one is well-formed XML that a conformant parser must accept, every one
 * changes the bytes of the part, and not one is emitted anywhere else in this
 * repository. The single-quoted attribute is the sharpest: it is the only form
 * here that a serializer cannot reproduce by remembering one boolean, because
 * the quote character has to be recorded per attribute rather than per part.
 *
 * ## And PowerPoint normalises all of it away
 *
 * Measured by re-saving on 2026-08-28: the thirty-five spaced self-closing tags
 * come back as **zero**. PowerPoint re-serializes any part it rewrites in its
 * own lexical form, which is the behaviour to expect and the reason this deck
 * has to be generated rather than authored.
 *
 * It also means a deck that has been through PowerPoint is no longer a fixture
 * for `C-LEX`: the round trip that proves a package still opens is the same
 * round trip that erases the lexical conventions the rule exists to check. The
 * two verifications cannot be run on the same file, and this is the deck where
 * that stops being a technicality.
 *
 * ## What is deliberately not here
 *
 * No pretty-printing. Whitespace **between** elements is character data whose
 * significance depends on `xml:space` and on the schema; everything above is
 * whitespace **inside markup**, which has no meaning at all and must
 * nevertheless survive byte for byte. Conflating the two is how a serializer
 * ends up "tidying" a `p:timing` tree, and the two belong in different decks.
 */

const TIGHT = '';
const SPACE = ' ';
const TWO_SPACES = '  ';
const TAB = '\t';

/** `<name attrs GAP/>`. `GAP` is the whitespace this deck is about. */
function empty(name: string, attrs: string, gap: string): string {
  return '<' + name + (attrs === '' ? '' : ' ' + attrs) + gap + '/>';
}

interface Lexis {
  /** Whitespace before `/>` on every self-closing tag in this shape. */
  readonly gap: string;
  /** The quote character for every attribute in this shape's geometry. */
  readonly quote: string;
  /** Whitespace either side of `=` in every attribute. */
  readonly eq: string;
  /** Whitespace before `>` on a start tag and on an end tag. */
  readonly tagGap: string;
}

const attr = (name: string, value: string, lexis: Lexis): string =>
  name + lexis.eq + '=' + lexis.eq + lexis.quote + value + lexis.quote;

/**
 * One probe shape, whole, hand-written.
 *
 * The chassis's `shape()` emits tight markup and is right to; a deck about the
 * bytes cannot go through a builder that decides them.
 */
function box(
  id: number,
  name: string,
  cell: Cell,
  lexis: Lexis,
  lines: readonly string[],
  accent: string,
): string {
  const a = (n: string, v: string): string => attr(n, v, lexis);
  const gap = lexis.tagGap;
  return (
    '<p:sp><p:nvSpPr>' +
    empty('p:cNvPr', a('id', String(id)) + ' ' + a('name', name), lexis.gap) +
    empty('p:cNvSpPr', '', lexis.gap) +
    empty('p:nvPr', '', lexis.gap) +
    '</p:nvSpPr>' +
    '<p:spPr' +
    gap +
    '>' +
    '<a:xfrm' +
    gap +
    '>' +
    empty('a:off', a('x', String(cell.x)) + ' ' + a('y', String(cell.y)), lexis.gap) +
    empty('a:ext', a('cx', String(cell.cx)) + ' ' + a('cy', String(cell.cy)), lexis.gap) +
    '</a:xfrm' +
    gap +
    '>' +
    '<a:prstGeom ' +
    a('prst', 'roundRect') +
    gap +
    '>' +
    empty('a:avLst', '', lexis.gap) +
    '</a:prstGeom' +
    gap +
    '>' +
    '<a:solidFill><a:schemeClr ' +
    a('val', accent) +
    gap +
    '>' +
    empty('a:lumMod', a('val', '75000'), lexis.gap) +
    '</a:schemeClr></a:solidFill>' +
    '</p:spPr>' +
    txBody({
      bodyPr: '<a:bodyPr wrap="square" anchor="ctr"/>',
      paras: lines
        .map((line, index) => textLine(line, { sz: 1200, b: index === 0 }, { algn: 'ctr' }))
        .join(''),
    }) +
    '</p:sp>'
  );
}

const cell = grid(2, 2);

/** The self-closing forms: identical shapes differing in one whitespace run. */
const TIGHT_LEXIS: Lexis = { gap: TIGHT, quote: '"', eq: '', tagGap: '' };
const SPACED: Lexis = { ...TIGHT_LEXIS, gap: SPACE };
const TWO: Lexis = { ...TIGHT_LEXIS, gap: TWO_SPACES };
const TABBED: Lexis = { ...TIGHT_LEXIS, gap: TAB };

/** The attribute-level and tag-level forms. */
const SINGLE_QUOTED: Lexis = { gap: TIGHT, quote: "'", eq: '', tagGap: '' };
const SPACED_EQUALS: Lexis = { gap: TIGHT, quote: '"', eq: ' ', tagGap: '' };
const TAG_GAP: Lexis = { gap: TIGHT, quote: '"', eq: '', tagGap: ' ' };
const EVERYTHING: Lexis = { gap: SPACE, quote: "'", eq: ' ', tagGap: ' ' };

export const a36SpacedTags: ProbeDeck = {
  id: 'a36-spaced-tags',
  title: 'PPTX Studio corpus: a36 spaced tags',
  description:
    'The seven lexical forms no producer in this repository emits, mixed sibling by sibling inside ' +
    'single slide parts rather than uniformly across one: whitespace before /> spelled with one ' +
    'space, two spaces and a tab; whitespace before > on a start tag and on an end tag; whitespace ' +
    'either side of =; and single-quoted attribute values. Every form is a plain XML 1.0 ' +
    'production. Mixing them within a part is what distinguishes a per-node serializer from a ' +
    'per-part one, which is architectural bet 2 and cannot be tested by a part that is uniformly ' +
    'anything. ADR 0005 said no available producer wrote the spaced form; PowerPoint&apos;s ChartEx ' +
    'serializer writes it 12 times out of 12, which is what makes this a measurement.',
  features: {
    shape: 18,
    placeholder: 6,
    presetGeom: 12,
    gradientFill: 2,
  },
  build: () => ({
    title: 'PPTX Studio corpus: a36 spaced tags',
    slides: [
      {
        title: 'a36 — four spellings of the same self-closing tag',
        body:
          box(
            10,
            'Tight',
            cell(0),
            TIGHT_LEXIS,
            ['tight', '<a:off .../>', 'what we emit'],
            'accent1',
          ) +
          box(
            11,
            'One space',
            cell(1),
            SPACED,
            ['one space', '<a:off ... />', 'ChartEx writes this'],
            'accent2',
          ) +
          box(
            12,
            'Two spaces',
            cell(2),
            TWO,
            ['two spaces', '<a:off ...  />', 'S is one or more'],
            'accent3',
          ) +
          box(
            13,
            'A tab',
            cell(3),
            TABBED,
            ['a tab', '<a:off ...\\t/>', 'a tab is S too'],
            'accent4',
          ),
      },
      {
        title: 'a36 — whitespace and quoting inside markup',
        body:
          box(
            10,
            'Single quoted',
            cell(0),
            SINGLE_QUOTED,
            ['single quotes', "val='75000'", 'recorded per attribute'],
            'accent5',
          ) +
          box(
            11,
            'Spaced equals',
            cell(1),
            SPACED_EQUALS,
            ['spaced equals', 'val = "75000"', 'Eq ::= S? = S?'],
            'accent6',
          ) +
          box(
            12,
            'Tag gap',
            cell(2),
            TAG_GAP,
            ['gap before >', '<p:spPr > and </a:xfrm >', 'STag and ETag both'],
            'accent1',
          ) +
          box(
            13,
            'All of them',
            cell(3),
            EVERYTHING,
            ['all four at once', 'spaced, quoted, gapped', 'still well-formed'],
            'accent2',
          ),
      },
      {
        title: 'a36 — siblings that disagree',
        body:
          box(10, 'Sibling tight', cell(0), TIGHT_LEXIS, ['sibling 1', 'tight'], 'accent3') +
          box(11, 'Sibling spaced', cell(1), SPACED, ['sibling 2', 'spaced'], 'accent4') +
          box(12, 'Sibling tight again', cell(2), TIGHT_LEXIS, ['sibling 3', 'tight'], 'accent5') +
          box(
            13,
            'Sibling spaced again',
            cell(3),
            SPACED,
            ['sibling 4', 'spaced', 'a part-level policy gets half of these wrong'],
            'accent6',
          ),
      },
    ],
  }),
};
