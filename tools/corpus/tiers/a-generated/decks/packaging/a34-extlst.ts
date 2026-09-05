import { grid, prstGeom, scheme, shape, solidFill } from '../../markup/shapes.ts';
import { textLine, txBody } from '../../markup/text.ts';
import type { ProbeDeck } from '../../markup/types.ts';

/**
 * `a:extLst` and `p:extLst` as what they are: an ordered, opaque list.
 *
 * ## The rule this deck exists to make falsifiable
 *
 * From the must-not-break rules: *never rebuild an `extLst` from a typed
 * model*. It is easy to agree with and easy to violate, because the natural
 * shape for an extension list in a typed model is a map from `@uri` to a
 * payload - and that one decision loses four things at once:
 *
 * 1. **Order.** `CT_OfficeArtExtensionList` is a sequence. A map does not have
 *    one, and re-emitting in insertion or sorted order is a different file.
 * 2. **Duplicates.** Nothing in the schema says `@uri` is unique. A map keyed
 *    on it silently drops the second.
 * 3. **Case.** GUID URIs appear upper- and lower-cased in the wild. A
 *    case-insensitive key collapses two extensions into one; a case-sensitive
 *    key that then *normalises* the string on write changes the bytes.
 * 4. **Content.** An extension is `xsd:any` with lax processing. Its payload
 *    may be an element we have never heard of, several of them, a text node, a
 *    CDATA section, a comment, or nothing at all.
 *
 * Slide 2 is one `a:extLst` containing all four hazards at once, so a
 * round-trip that survives it cannot be storing extensions in a map.
 *
 * ## Where an extLst can be, which is nearly everywhere
 *
 * Every host in this deck is a different complex type and a different position
 * in its sequence, and in all of them `extLst` is **last**:
 *
 * | host          | sequence, ending in                                   |
 * | ------------- | ------------------------------------------------------ |
 * | `p:cNvPr`     | `hlinkClick, hlinkHover, extLst`                       |
 * | `p:spPr`      | `…, effectLst, scene3d, sp3d, extLst`                  |
 * | `a:bodyPr`    | `prstTxWarp, <autofit>, scene3d, <3d>, extLst`         |
 * | `p:cSld`      | `bg, spTree, custDataLst, controls, extLst`            |
 * | `p:sld`       | `cSld, clrMapOvr, transition, timing, extLst`          |
 * | `p:presentation` | `…, defaultTextStyle, modifyVerifier, extLst`       |
 *
 * The `p:cNvPr` one and the `p:sld` one are the two a real deck always has -
 * `adec:decorative` lives in the first and `p14:creationId` in `p:cSld`'s - and
 * the pair on one shape is the trap `a19-decorative` documents: a shape has
 * two extension lists and they are not interchangeable.
 *
 * ## The URIs are unknown on purpose
 *
 * Not one of them is an extension any application implements. That is the
 * point: the contract is that an unknown `@uri` is carried through untouched,
 * and a fixture built from URIs we *do* understand would test the opposite
 * property. They are `example`-domain and `urn:` forms so they cannot ever
 * collide with a real one, and the GUID-shaped pair is valid hexadecimal so
 * that anything parsing it as a GUID succeeds and still has to keep the case.
 *
 * ## What PowerPoint does with them, measured
 *
 * It opens without complaint, which is the specified behaviour for lax
 * processing and worth having as a measurement rather than an assumption. What
 * the re-save pass then found on 2026-08-28 is better than that: **the contract
 * holds.**
 *
 * All thirteen `a:ext` children of slide 2's list come back **in the order they
 * were written**, with `{C0A9F0B1-3333-…}` still appearing **twice** and
 * `{C0A9F0B1-2222-…}` and `{c0a9f0b1-2222-…}` still two entries differing only
 * in case. The `a:ext` with no payload survives as an empty element. So all
 * four hazards above are hazards for *us*, and none of them is something
 * PowerPoint gets wrong — which makes this deck a fixture with a known right
 * answer rather than a question.
 *
 * Two things do not survive, and both are worth more than the ones that do:
 *
 * - **Only the last child element of an `a:ext` is kept.** The extension
 *   holding `a`, `b`, `c` comes back holding `c`; the one holding four
 *   collide-named children comes back holding the fourth. PowerPoint treats
 *   `CT_OfficeArtExtension` as carrying one element however many its `xsd:any`
 *   permits. A writer that appends a second child to an existing extension
 *   loses the first, silently, on the next PowerPoint save.
 * - **A comment inside an extension is dropped. A CDATA section is not.** Both
 *   are things a typed model would lose; only one is something PowerPoint
 *   loses. Preserving the CDATA is the stronger of the two results, because it
 *   means the payload is being carried as text rather than re-serialized from
 *   a parse.
 *
 * Every `p:extLst` in this deck survives too, at all three levels, and
 * PowerPoint adds two of its own on every deck in the corpus - `p14:creationId`
 * and its neighbours - which is the baseline any diff has to subtract first.
 */

const NS = 'http://pptx-studio.example/ns/a34';

/** Two URIs differing only in case. A case-folding key collapses them. */
const UPPER = '{C0A9F0B1-2222-4A22-8222-222222222222}';
const LOWER = '{c0a9f0b1-2222-4a22-8222-222222222222}';

/** One URI used twice in the same list. Nothing in the schema forbids it. */
const TWICE = '{C0A9F0B1-3333-4A33-8333-333333333333}';

/** `<a:ext uri="…">payload</a:ext>`, or an empty one when there is no payload. */
function ext(uri: string, payload = ''): string {
  return payload === '' ? `<a:ext uri="${uri}"/>` : `<a:ext uri="${uri}">${payload}</a:ext>`;
}

/** The same, in the PresentationML namespace. A different element, same idea. */
function pExt(uri: string, payload = ''): string {
  return payload === '' ? `<p:ext uri="${uri}"/>` : `<p:ext uri="${uri}">${payload}</p:ext>`;
}

const extLst = (...children: readonly string[]): string =>
  '<a:extLst>' + children.join('') + '</a:extLst>';
const pExtLst = (...children: readonly string[]): string =>
  '<p:extLst>' + children.join('') + '</p:extLst>';

/**
 * A payload element in a namespace nothing here implements.
 *
 * Declared on the payload itself rather than on the part root, which is what
 * every real extension does - `a16:creationId`, `p14:media`, `asvg:svgBlip` -
 * and is the reason a serializer that rewrites prefixes breaks extensions
 * before it breaks anything else.
 */
function payload(local: string, attrs = '', inner = ''): string {
  const open = `<zz:${local} xmlns:zz="${NS}"${attrs === '' ? '' : ' ' + attrs}`;
  return inner === '' ? open + '/>' : open + '>' + inner + `</zz:${local}>`;
}

const cell = grid(2, 2);

/** A labelled probe shape. The label is the only thing a reader sees. */
function box(
  id: number,
  name: string,
  index: number,
  accent: string,
  lines: readonly string[],
  extras: {
    readonly cNvPrExt?: string;
    readonly spPrExt?: string;
    readonly bodyPrExt?: string;
  } = {},
): string {
  const c = cell(index);
  return shape({
    id,
    name,
    x: c.x,
    y: c.y,
    cx: c.cx,
    cy: c.cy,
    geometry: prstGeom('roundRect'),
    fill: solidFill(scheme(accent, '<a:lumMod val="60000"/><a:lumOff val="40000"/>')),
    ...(extras.cNvPrExt === undefined ? {} : { extLst: extras.cNvPrExt }),
    ...(extras.spPrExt === undefined ? {} : { spPrExtLst: extras.spPrExt }),
    textBody: txBody({
      bodyPr:
        extras.bodyPrExt === undefined
          ? '<a:bodyPr wrap="square" anchor="ctr"/>'
          : '<a:bodyPr wrap="square" anchor="ctr">' + extras.bodyPrExt + '</a:bodyPr>',
      paras: lines
        .map((line, i) => textLine(line, { sz: 1200, b: i === 0 }, { algn: 'ctr' }))
        .join(''),
    }),
  });
}

export const a34ExtLst: ProbeDeck = {
  id: 'a34-extlst',
  title: 'PPTX Studio corpus: a34 extLst',
  description:
    'a:extLst and p:extLst at six different hosts - p:cNvPr, p:spPr, a:bodyPr, p:cSld, p:sld and ' +
    'p:presentation - carrying nothing but URIs no application implements, so the only property ' +
    'under test is that an unknown extension is carried through untouched. One list holds all four ' +
    'hazards a map-keyed-on-uri model loses: a deliberately unsorted order, the same uri twice, two ' +
    'GUIDs differing only in case, and payloads that are variously absent, several elements, text, ' +
    'CDATA and a comment. Every payload declares its own namespace on itself, the way every real ' +
    'extension does, which is what a prefix-rewriting serializer breaks first.',
  features: {
    shape: 18,
    placeholder: 6,
    presetGeom: 12,
    gradientFill: 2,
  },
  build: () => ({
    title: 'PPTX Studio corpus: a34 extLst',
    slides: [
      {
        title: 'a34 — an extLst at every host a shape has',
        body:
          box(
            10,
            'cNvPr extension',
            0,
            'accent1',
            ['p:cNvPr/a:extLst', 'where adec:decorative lives'],
            {
              cNvPrExt: extLst(
                ext('{C0A9F0B1-1111-4A11-8111-111111111111}', payload('nonVisual', 'note="cNvPr"')),
              ),
            },
          ) +
          box(11, 'spPr extension', 1, 'accent2', ['p:spPr/a:extLst', 'after scene3d and sp3d'], {
            spPrExt: extLst(
              ext('{C0A9F0B1-1111-4A11-8111-111111111112}', payload('shapeProps', 'note="spPr"')),
            ),
          }) +
          box(
            12,
            'bodyPr extension',
            2,
            'accent3',
            ['a:bodyPr/a:extLst', 'after the 3-D children'],
            {
              bodyPrExt: extLst(
                ext(
                  '{C0A9F0B1-1111-4A11-8111-111111111113}',
                  payload('bodyProps', 'note="bodyPr"'),
                ),
              ),
            },
          ) +
          box(
            13,
            'All three at once',
            3,
            'accent4',
            ['all three on one shape', 'and they are not the same list'],
            {
              cNvPrExt: extLst(ext('{C0A9F0B1-1111-4A11-8111-111111111114}', payload('one'))),
              spPrExt: extLst(ext('{C0A9F0B1-1111-4A11-8111-111111111115}', payload('two'))),
              bodyPrExt: extLst(ext('{C0A9F0B1-1111-4A11-8111-111111111116}', payload('three'))),
            },
          ),
        cSldTail: pExtLst(
          pExt(
            '{C0A9F0B1-4444-4A44-8444-444444444441}',
            payload('cSldMarker', 'note="p:cSld, where p14:creationId lives"'),
          ),
        ),
      },
      {
        title: 'a34 — one list, four ways to lose it',
        body:
          box(
            10,
            'The hazardous list',
            0,
            'accent5',
            ['seven a:ext children', 'unsorted, duplicated, case-paired'],
            {
              spPrExt: extLst(
                // Deliberately not sorted, by any ordering: a serializer that
                // emits these in sorted or insertion-into-a-map order is
                // detectable from the bytes alone.
                ext('urn:pptx-studio:corpus:a34:third', payload('third')),
                ext('http://pptx-studio.example/ext/first', payload('first')),
                ext(TWICE, payload('duplicate', 'occurrence="1"')),
                ext(UPPER, payload('cased', 'spelling="upper"')),
                ext(TWICE, payload('duplicate', 'occurrence="2"')),
                ext(LOWER, payload('cased', 'spelling="lower"')),
                // No payload at all. `xsd:any minOccurs="0"`.
                ext('urn:pptx-studio:corpus:a34:empty'),
              ),
            },
          ) +
          box(
            11,
            'Payloads that are not one element',
            1,
            'accent6',
            ['several children, text,', 'CDATA and a comment'],
            {
              spPrExt: extLst(
                ext(
                  'urn:pptx-studio:corpus:a34:several',
                  payload('a') + payload('b') + payload('c'),
                ),
                ext(
                  'urn:pptx-studio:corpus:a34:nested',
                  payload('outer', 'depth="1"', payload('inner', 'depth="2"')),
                ),
                ext(
                  'urn:pptx-studio:corpus:a34:text',
                  payload('withText', '', 'character data inside an extension'),
                ),
                ext(
                  'urn:pptx-studio:corpus:a34:cdata',
                  payload('withCdata', '', '<![CDATA[<not markup> & not an entity]]>'),
                ),
                ext(
                  'urn:pptx-studio:corpus:a34:comment',
                  '<!-- a comment inside an extension -->' + payload('afterComment'),
                ),
              ),
            },
          ) +
          box(
            12,
            'A local name that collides',
            2,
            'accent1',
            ['zz:transition, zz:table,', 'zz:oMath — none of them are'],
            {
              spPrExt: extLst(
                ext(
                  'urn:pptx-studio:corpus:a34:collide',
                  payload('transition') + payload('table') + payload('oMath') + payload('audio'),
                ),
              ),
            },
          ) +
          box(13, 'No extension at all', 3, 'accent2', [
            'the control shape',
            'nothing to preserve',
          ]),
      },
      {
        title: 'a34 — the two extLst hosts outside a shape',
        body:
          box(10, 'Slide level', 0, 'accent3', ['p:sld/p:extLst', 'after p:timing']) +
          box(11, 'cSld level', 1, 'accent4', [
            'p:cSld/p:extLst',
            'one level in, a different list',
          ]) +
          box(12, 'Presentation level', 2, 'accent5', [
            'p:presentation/p:extLst',
            'the whole package',
          ]) +
          box(13, 'All three on this slide', 3, 'accent6', [
            'and the two on the slide',
            'are siblings of different parents',
          ]),
        cSldTail: pExtLst(
          pExt('urn:pptx-studio:corpus:a34:csld', payload('cSldLevel')),
          pExt('urn:pptx-studio:corpus:a34:csld-2', payload('cSldLevelAgain')),
        ),
        // `CT_Slide` is cSld, clrMapOvr, transition, timing, extLst, so this is
        // last and there is nowhere else for it to go.
        tail: pExtLst(pExt('urn:pptx-studio:corpus:a34:sld', payload('slideLevel'))),
      },
    ],
    // `CT_Presentation` ends `…, defaultTextStyle, modifyVerifier, extLst`.
    presentationTail: pExtLst(
      pExt('urn:pptx-studio:corpus:a34:presentation', payload('presentationLevel')),
      pExt('urn:pptx-studio:corpus:a34:presentation-2'),
    ),
  }),
};
