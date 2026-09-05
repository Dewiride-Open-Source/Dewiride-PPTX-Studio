import {
  DECLARATION,
  NS_A,
  NS_P,
  NS_R,
  type ProbePart,
  type ProbeRel,
} from '../../markup/chassis.ts';
import { grid, shape } from '../../markup/shapes.ts';
import { textLine, txBody } from '../../markup/text.ts';
import type { ProbeDeck } from '../../markup/types.ts';

/**
 * Comments, in both of the two formats a deck can carry - because PowerPoint
 * stopped writing the standard one.
 *
 * ## The census undercounts a deck commented today, and this deck is the proof
 *
 * ECMA-376 puts comments in `ppt/comments/commentN.xml` as a `p:cmLst` of
 * `p:cm`, with the author list in `ppt/commentAuthors.xml`. That is what the
 * census's `comment` rule matches, keyed on the content type ending
 * `presentationml.comments+xml`.
 *
 * PowerPoint 16.0.20326 does not write it. A comment added on 2026-08-27
 * through `Slide.Comments.Add` produced instead:
 *
 * | part                                       | content type                            |
 * | ------------------------------------------ | --------------------------------------- |
 * | `ppt/comments/modernComment_100_4A8498BC.xml` | `application/vnd.ms-powerpoint.comments+xml` |
 * | `ppt/authors.xml`                          | `application/vnd.ms-powerpoint.authors+xml`  |
 *
 * with the root elements `p188:cmLst` and `p188:authorLst` in
 * `http://schemas.microsoft.com/office/powerpoint/2018/8/main`, a slide
 * relationship of type `.../office/2018/10/relationships/comments`, and a
 * `p188:commentRel` in the slide's own `p:extLst` pointing at it. The comment
 * identifies its slide through `pc:sldMkLst` - a document marker carrying
 * `@sldId` and the slide's `p14:creationId` - rather than through the
 * relationship direction.
 *
 * So a census of a deck commented in PowerPoint this year reports **zero
 * comments**, and the number is not wrong about the markup: it is right about
 * `p:cmLst` and blind to what replaced it. This deck carries both forms so the
 * gap is a thing the corpus states rather than a thing nobody noticed. Its
 * `features` map says `comment: 2` and those two are the classic parts.
 *
 * ## Anchoring
 *
 * A classic `p:cm` has a `p:pos` and nothing else - it is pinned to a point on
 * the slide, not to a shape, and moving the shape leaves the marker behind.
 * That is the whole of the standard's model, and slide 1 shows it with two
 * authors commenting at two points.
 *
 * A modern comment **can** anchor to a shape, through a `pc:spMkLst` beside the
 * `pc:sldMkLst`. This deck does not write one. Nothing in any specification
 * describes that structure and the measurement above produced only the slide
 * marker, so writing a shape marker would be inventing markup - which is the
 * one thing the architecture forbids. It is a gap, it is declared here, and it
 * closes with a Tier B deck that comments on a shape.
 */

const NS_P188 = 'http://schemas.microsoft.com/office/powerpoint/2018/8/main';
const NS_PC = 'http://schemas.microsoft.com/office/powerpoint/2013/main/command';

const CT_COMMENTS = 'application/vnd.openxmlformats-officedocument.presentationml.comments+xml';
const CT_COMMENT_AUTHORS =
  'application/vnd.openxmlformats-officedocument.presentationml.commentAuthors+xml';
const CT_MODERN_COMMENTS = 'application/vnd.ms-powerpoint.comments+xml';
const CT_MODERN_AUTHORS = 'application/vnd.ms-powerpoint.authors+xml';

const REL_COMMENTS = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments';
const REL_COMMENT_AUTHORS =
  'http://schemas.openxmlformats.org/officeDocument/2006/relationships/commentAuthors';
const REL_MODERN_COMMENTS = 'http://schemas.microsoft.com/office/2018/10/relationships/comments';
const REL_MODERN_AUTHORS = 'http://schemas.microsoft.com/office/2018/10/relationships/authors';

const NS_DECLS = `xmlns:a="${NS_A}" xmlns:r="${NS_R}" xmlns:p="${NS_P}"`;

/**
 * `p:cmAuthorLst`.
 *
 * `@clrIdx` is what colours the marker, and it is an index rather than a
 * colour: two authors with the same `clrIdx` are indistinguishable on screen.
 * `@lastIdx` is the highest `p:cm/@idx` this author has used, and it is the
 * only thing that stops two comments by one author colliding.
 */
const COMMENT_AUTHORS =
  DECLARATION +
  `<p:cmAuthorLst ${NS_DECLS}>` +
  '<p:cmAuthor id="1" name="Corpus Reviewer" initials="CR" lastIdx="2" clrIdx="0"/>' +
  '<p:cmAuthor id="2" name="Corpus Second Reader" initials="CS" lastIdx="1" clrIdx="1"/>' +
  '</p:cmAuthorLst>';

/** `p:cmLst`. `CT_Comment` is `pos, text, extLst`, and `pos` is not optional. */
function classicComments(
  comments: readonly {
    readonly authorId: number;
    readonly idx: number;
    readonly x: number;
    readonly y: number;
    readonly text: string;
  }[],
): string {
  return (
    DECLARATION +
    `<p:cmLst ${NS_DECLS}>` +
    comments
      .map(
        (comment) =>
          `<p:cm authorId="${String(comment.authorId)}" dt="2026-08-27T09:00:00.000"` +
          ` idx="${String(comment.idx)}">` +
          `<p:pos x="${String(comment.x)}" y="${String(comment.y)}"/>` +
          `<p:text>${comment.text}</p:text>` +
          '</p:cm>',
      )
      .join('') +
    '</p:cmLst>'
  );
}

/** The 2018 form, written exactly as PowerPoint 16.0.20326 was measured writing it. */
const MODERN_AUTHORS =
  DECLARATION +
  `<p188:authorLst ${NS_DECLS} xmlns:p188="${NS_P188}">` +
  '<p188:author id="{0D4E7A21-93C5-4F86-B107-52E8D3A9146B}" name="Corpus Reviewer"' +
  ' initials="CR" userId="Corpus" providerId="None"/>' +
  '</p188:authorLst>';

const SLIDE_TWO_CREATION_ID = '1500000002';

const MODERN_COMMENT =
  DECLARATION +
  `<p188:cmLst ${NS_DECLS} xmlns:p188="${NS_P188}">` +
  '<p188:cm id="{5C0A93B7-6E41-4D28-8B95-7F320A61D4E8}"' +
  ' authorId="{0D4E7A21-93C5-4F86-B107-52E8D3A9146B}" created="2026-08-27T09:05:00.000">' +
  `<pc:sldMkLst xmlns:pc="${NS_PC}"><pc:docMk/>` +
  `<pc:sldMk cId="${SLIDE_TWO_CREATION_ID}" sldId="257"/></pc:sldMkLst>` +
  '<p188:pos x="4572000" y="2286000"/>' +
  '<p188:txBody><a:bodyPr/><a:lstStyle/>' +
  '<a:p><a:r><a:rPr lang="en-GB"/><a:t>The 2018 comment format, on slide 2.</a:t></a:r></a:p>' +
  '</p188:txBody></p188:cm></p188:cmLst>';

const PARTS: readonly ProbePart[] = [
  {
    name: 'ppt/commentAuthors.xml',
    bytes: COMMENT_AUTHORS,
    contentType: { kind: 'override', type: CT_COMMENT_AUTHORS },
  },
  {
    name: 'ppt/comments/comment1.xml',
    bytes: classicComments([
      {
        authorId: 1,
        idx: 1,
        x: 1704,
        y: 1029,
        text: 'A p:cm from the first author, pinned to a point rather than to a shape.',
      },
      {
        authorId: 2,
        idx: 1,
        x: 4820,
        y: 1029,
        text: 'A second author on the same slide, with its own clrIdx.',
      },
    ]),
    contentType: { kind: 'override', type: CT_COMMENTS },
  },
  {
    name: 'ppt/comments/comment2.xml',
    bytes: classicComments([
      {
        authorId: 1,
        idx: 2,
        x: 1704,
        y: 1029,
        text: 'The classic form and the 2018 form, on one slide.',
      },
    ]),
    contentType: { kind: 'override', type: CT_COMMENTS },
  },
  {
    name: 'ppt/authors.xml',
    bytes: MODERN_AUTHORS,
    contentType: { kind: 'override', type: CT_MODERN_AUTHORS },
  },
  {
    name: 'ppt/comments/modernComment_257_4A8498BC.xml',
    bytes: MODERN_COMMENT,
    contentType: { kind: 'override', type: CT_MODERN_COMMENTS },
  },
];

/**
 * Both author lists hang off the presentation, not off a slide. Left out of
 * these relationships each one is an orphan part, which is what the census
 * reported the first time the benchmark generator wrote a comment.
 *
 * `rId9` and `rId10` because the chassis has already allocated eight: the
 * master, three slides, presProps, viewProps, theme and tableStyles. rIds are
 * `xsd:ID` scoped to one `.rels` part, so a deck that adds its own has to know
 * what the part already holds - which is the same coupling `a13-sections` has
 * to live with, and the reason a global rId allocator is a bug.
 */
const PRESENTATION_RELS: readonly ProbeRel[] = [
  { id: 'rId9', type: REL_COMMENT_AUTHORS, target: 'commentAuthors.xml' },
  { id: 'rId10', type: REL_MODERN_AUTHORS, target: 'authors.xml' },
];

/** A caption saying what a reader should be looking for on this slide. */
function caption(id: number, lines: readonly string[]): string {
  return shape({
    id,
    name: 'Caption ' + String(id),
    ...grid(1, 3)(0),
    textBody: txBody({ paras: lines.map((line) => textLine(line, { sz: 1600 })).join('') }),
  });
}

export const a15Comments: ProbeDeck = {
  id: 'a15-comments',
  title: 'PPTX Studio corpus: a15 comments',
  description:
    'Comments in both formats: the standard p:cmLst with p:cmAuthorLst on two slides, and the ' +
    '2018 p188:cmLst with ppt/authors.xml that PowerPoint 16.0.20326 actually writes today. The ' +
    "census's comment rule matches only the first, so a deck commented in PowerPoint this year " +
    'reports zero - which is the finding this deck exists to record. A classic comment is pinned ' +
    'to a point and never to a shape; the shape-anchored modern form is declared here as an ' +
    'unmeasured gap rather than invented.',
  features: {
    shape: 9,
    placeholder: 6,
    presetGeom: 3,
    gradientFill: 2,
    comment: 2,
  },
  build: () => ({
    title: 'PPTX Studio corpus: a15 comments',
    parts: PARTS,
    presentationRels: PRESENTATION_RELS,
    slides: [
      {
        title: 'a15 — two classic comments, two authors',
        layout: 0,
        body: caption(10, [
          'ppt/comments/comment1.xml holds two p:cm elements,',
          'one per author, each with a p:pos and no shape.',
        ]),
        rels: [{ id: 'rId2', type: REL_COMMENTS, target: '../comments/comment1.xml' }],
      },
      {
        title: 'a15 — the classic form and the 2018 form together',
        layout: 0,
        body: caption(10, [
          'comment2.xml is a p:cmLst; modernComment_257_4A8498BC.xml is a p188:cmLst.',
          'Only the first is what the census counts.',
        ]),
        rels: [
          { id: 'rId2', type: REL_COMMENTS, target: '../comments/comment2.xml' },
          {
            id: 'rId3',
            type: REL_MODERN_COMMENTS,
            target: '../comments/modernComment_257_4A8498BC.xml',
          },
        ],
        // The creation id the modern comment's `pc:sldMk` names. PowerPoint
        // writes one of these on every slide, and it is what a document marker
        // resolves against - so a deck with modern comments and no creation
        // ids is a deck whose comments point at nothing.
        cSldTail:
          '<p:extLst><p:ext uri="{BB962C8B-B14F-4D97-AF65-F5344CB8AC3E}">' +
          '<p14:creationId xmlns:p14="http://schemas.microsoft.com/office/powerpoint/2010/main"' +
          ` val="${SLIDE_TWO_CREATION_ID}"/></p:ext></p:extLst>`,
        tail:
          '<p:extLst><p:ext uri="{6950BFC3-D8DA-4A85-94F7-54DA5524770B}">' +
          `<p188:commentRel xmlns:p188="${NS_P188}" r:id="rId3"/></p:ext></p:extLst>`,
      },
      {
        title: 'a15 — a slide with no comments at all',
        layout: 0,
        body: caption(10, [
          'No comment relationship and no comment part.',
          'An author list with nothing pointing at it is still reachable',
          'from the presentation, which is why it is not an orphan.',
        ]),
      },
    ],
  }),
};
