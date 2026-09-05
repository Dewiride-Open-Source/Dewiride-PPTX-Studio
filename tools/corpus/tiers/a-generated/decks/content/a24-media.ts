import { cNvPrXml, REL, type ProbePart, type ProbeRel } from '../../markup/chassis.ts';
import { probeMp4, probeWav } from '../../assets/media.ts';
import { quadrantPng } from '../../assets/png.ts';
import { shape } from '../../markup/shapes.ts';
import { textLine, txBody } from '../../markup/text.ts';
import type { ProbeDeck } from '../../markup/types.ts';

/**
 * Media: one part, two relationships, and an `r:id` that points at nothing.
 *
 * Read back from a deck PowerPoint 16.0.20326 wrote on 2026-08-27 after
 * `Shapes.AddMediaObject2` on a WAV the corpus generator authored. Every
 * relationship type, content type and attribute below is from that file.
 *
 * ## The double relationship
 *
 * An embedded media part is related to **twice, by two different types, from
 * the same slide**:
 *
 * ```xml
 * <Relationship Id="rId2" Type="…/2007/relationships/media" Target="../media/media1.wav"/>
 * <Relationship Id="rId3" Type="…/officeDocument/2006/relationships/audio" Target="../media/media1.wav"/>
 * ```
 *
 * The 2006 one is named by `a:audioFile`; the 2007 Microsoft one is named by
 * `p14:media` inside `p:nvPr`'s extension list. Two rIds, one part. Sub-phase
 * 10.8 has to de-duplicate that on export or the media GC keeps one edge and
 * drops the other, and 1.3's mark-and-sweep has to treat both as roots.
 *
 * ## `@r:link` on something that is not linked
 *
 * `<a:audioFile r:link="rId3"/>` - and the target is **internal**. The
 * attribute is called `link` for both embedded and linked media, so the only
 * way to tell them apart is to look at whether the relationship carries
 * `TargetMode="External"`. Inferring from the attribute name gets it backwards
 * for every embedded clip in existence. Slide 2 has one of each so the
 * difference is visible in one place.
 *
 * ## An `r:id` whose value is the empty string
 *
 * ```xml
 * <a:hlinkClick r:id="" action="ppaction://media"/>
 * ```
 *
 * on the `p:cNvPr` of every media shape. It is not a dangling relationship, it
 * is a deliberately empty one - the `action` is the whole content and the id
 * is a required attribute with nothing to put in it. Sub-phase 1.2's rule that
 * every `r:id` resolves has to exempt the empty string, and a reader that
 * treats `""` as a lookup miss reports a corrupt package for a file PowerPoint
 * wrote itself.
 *
 * ## One WAV is not enough, and that was measured the hard way
 *
 * The transition sound on slide 3 is a **second part** holding the same bytes.
 * With one shared WAV this deck was refused - and interestingly, not with the
 * usual message: "PowerPoint could not open the file", where every other
 * refusal this corpus has found says "the file or directory is corrupted and
 * unreadable". Slides 1 and 2 opened together, slides 2 and 3 opened together,
 * and slides 1 and 3 - the media object and the transition sound, the two that
 * shared the WAV - did not. Splitting the part fixed it, which is also what
 * PowerPoint itself does: it writes `media1.wav` for a media object and
 * `audio1.wav` for a transition sound even when a user picked the same file.
 *
 * ## Two gaps this deck closes
 *
 * `a16-transitions` could not write `p:transition/p:sndAc/p:stSnd` and
 * `a17-animations` could not write `p:audio` or `p:video`, because all three
 * need a media part. Slide 1 carries the measured `p:audio` node - a
 * `p:cMediaNode vol="80000"` whose `p:cTn` has `display="0"` and an end
 * condition of `onStopAudio` against `p:sldTgt` - slide 2 carries `p:video`,
 * and slide 3 carries
 *
 * ```xml
 * <p:sndAc><p:stSnd><p:snd r:embed="rId2" name="probe.wav"/></p:stSnd></p:sndAc>
 * ```
 *
 * which is where the `@name` attribute turns out to be the original file name,
 * not the part name.
 *
 * ## What is not here, and why
 *
 * The video part is a valid ISO-BMFF container with no tracks - see
 * `tools/corpus/tiers/a-generated/assets/media.ts`. **PowerPoint refuses to insert it**: "PowerPoint
 * cannot insert a video from the selected file. Verify that the necessary codec
 * for this media format is installed." That is an insert-time check on content,
 * not a package rule, and the markup here is the markup it writes for a video
 * it did accept. A clip that decodes needs an encoder, an encoder needs a
 * download, and a Tier B deck closes it. `ROSTER.md` carries the gap.
 */

const NS_P14 = 'http://schemas.microsoft.com/office/powerpoint/2010/main';

/** The 2007 media relationship, which is not under the ECMA base. */
const REL_MEDIA = 'http://schemas.microsoft.com/office/2007/relationships/media';
/** `p:ext/@uri` for `p14:media`. Measured. */
const MEDIA_EXT_URI = '{DAA4B4D4-6D71-4841-9C94-3DE7FCFB9230}';

const POSTER = { cx: 2743200, cy: 2057400 } as const;

/**
 * A media shape.
 *
 * `p:nvPr` is `ph, <media>, custDataLst, extLst`, so the audio or video element
 * comes before the extension list that names the same file again.
 */
function mediaPic(spec: {
  readonly id: number;
  readonly name: string;
  readonly descr: string;
  /** `a:audioFile` or `a:videoFile`. */
  readonly element: 'audioFile' | 'videoFile';
  /** The 2006 relationship, named by `@r:link` whether embedded or not. */
  readonly linkRelId: string;
  /** The 2007 relationship. Absent on linked media, which has no local bytes. */
  readonly mediaRelId?: string;
  readonly posterRelId: string;
  readonly x: number;
  readonly y: number;
}): string {
  return (
    '<p:pic><p:nvPicPr>' +
    cNvPrXml({
      id: spec.id,
      name: spec.name,
      descr: spec.descr,
      // A required r:id with nothing in it. See the file comment.
      hlinkClick: '<a:hlinkClick r:id="" action="ppaction://media"/>',
    }) +
    '<p:cNvPicPr><a:picLocks noChangeAspect="1"/></p:cNvPicPr>' +
    '<p:nvPr>' +
    `<a:${spec.element} r:link="${spec.linkRelId}"/>` +
    (spec.mediaRelId === undefined
      ? ''
      : `<p:extLst><p:ext uri="${MEDIA_EXT_URI}">` +
        `<p14:media xmlns:p14="${NS_P14}" r:embed="${spec.mediaRelId}"/>` +
        '</p:ext></p:extLst>') +
    '</p:nvPr>' +
    '</p:nvPicPr>' +
    `<p:blipFill><a:blip r:embed="${spec.posterRelId}"/>` +
    '<a:stretch><a:fillRect/></a:stretch></p:blipFill>' +
    '<p:spPr>' +
    `<a:xfrm><a:off x="${String(spec.x)}" y="${String(spec.y)}"/>` +
    `<a:ext cx="${String(POSTER.cx)}" cy="${String(POSTER.cy)}"/></a:xfrm>` +
    '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom>' +
    '</p:spPr></p:pic>'
  );
}

// ------------------------------------------------------------------ timing

/** The `p:cMediaNode` PowerPoint writes, wrapped in `p:audio` or `p:video`. */
function mediaNode(tag: 'audio' | 'video', tnId: number, spid: number): string {
  return (
    `<p:${tag}><p:cMediaNode vol="80000">` +
    `<p:cTn id="${String(tnId)}" fill="hold" display="0">` +
    '<p:stCondLst><p:cond delay="indefinite"/></p:stCondLst>' +
    // `onStopAudio` for video too: the enumeration has no video member.
    '<p:endCondLst><p:cond evt="onStopAudio" delay="0">' +
    '<p:tgtEl><p:sldTgt/></p:tgtEl></p:cond></p:endCondLst>' +
    '</p:cTn>' +
    `<p:tgtEl><p:spTgt spid="${String(spid)}"/></p:tgtEl>` +
    '</p:cMediaNode>' +
    `</p:${tag}>`
  );
}

/**
 * The interactive sequence a click on a media shape runs.
 *
 * `nodeType="interactiveSeq"`, `presetClass="mediacall"` and a `p:cmd` whose
 * `@cmd` is the string `playFrom(0.0)` - a tiny expression language inside an
 * attribute, and one more reason animation is a preserve-only feature here.
 */
function mediaTiming(spids: readonly number[], nodes: string): string {
  const sequences = spids
    .map((spid, index) => {
      const base = 2 + index * 5;
      return (
        '<p:seq concurrent="1" nextAc="seek">' +
        `<p:cTn id="${String(base)}" restart="whenNotActive" fill="hold"` +
        ' evtFilter="cancelBubble" nodeType="interactiveSeq">' +
        `<p:stCondLst><p:cond evt="onClick" delay="0"><p:tgtEl><p:spTgt spid="${String(spid)}"/>` +
        '</p:tgtEl></p:cond></p:stCondLst>' +
        '<p:endSync evt="end" delay="0"><p:rtn val="all"/></p:endSync>' +
        '<p:childTnLst>' +
        `<p:par><p:cTn id="${String(base + 1)}" fill="hold">` +
        '<p:stCondLst><p:cond delay="0"/></p:stCondLst><p:childTnLst>' +
        `<p:par><p:cTn id="${String(base + 2)}" fill="hold">` +
        '<p:stCondLst><p:cond delay="0"/></p:stCondLst><p:childTnLst>' +
        `<p:par><p:cTn id="${String(base + 3)}" presetID="1" presetClass="mediacall"` +
        ' presetSubtype="0" fill="hold" nodeType="clickEffect">' +
        '<p:stCondLst><p:cond delay="0"/></p:stCondLst><p:childTnLst>' +
        '<p:cmd type="call" cmd="playFrom(0.0)"><p:cBhvr>' +
        `<p:cTn id="${String(base + 4)}" dur="250" fill="hold"/>` +
        `<p:tgtEl><p:spTgt spid="${String(spid)}"/></p:tgtEl>` +
        '</p:cBhvr></p:cmd>' +
        '</p:childTnLst></p:cTn></p:par>' +
        '</p:childTnLst></p:cTn></p:par>' +
        '</p:childTnLst></p:cTn></p:par>' +
        '</p:childTnLst></p:cTn>' +
        `<p:nextCondLst><p:cond evt="onClick" delay="0"><p:tgtEl><p:spTgt spid="${String(spid)}"/>` +
        '</p:tgtEl></p:cond></p:nextCondLst>' +
        '</p:seq>'
      );
    })
    .join('');
  return (
    '<p:timing><p:tnLst><p:par>' +
    '<p:cTn id="1" dur="indefinite" restart="never" nodeType="tmRoot"><p:childTnLst>' +
    sequences +
    nodes +
    '</p:childTnLst></p:cTn></p:par></p:tnLst></p:timing>'
  );
}

// ------------------------------------------------------------------ package

const PARTS: readonly ProbePart[] = [
  {
    name: 'ppt/media/media1.wav',
    bytes: probeWav(),
    // Measured: `audio/x-wav`, not `audio/wav`, and a Default rather than an
    // Override - the same treatment png and jpeg get.
    contentType: { kind: 'default', extension: 'wav', type: 'audio/x-wav' },
  },
  {
    // The transition sound is a **second copy** of the same tone, and it has to
    // be. A single WAV that is both a media object's target on one slide and a
    // p:snd on another is a whole-package refusal - see the file comment - and
    // PowerPoint's own naming keeps them apart too: media1.wav for the object,
    // audio1.wav for the sound.
    name: 'ppt/media/audio1.wav',
    bytes: probeWav(),
  },
  {
    name: 'ppt/media/media2.mp4',
    bytes: probeMp4(),
    contentType: { kind: 'default', extension: 'mp4', type: 'video/mp4' },
  },
  {
    name: 'ppt/media/image1.png',
    bytes: quadrantPng(),
    contentType: { kind: 'default', extension: 'png', type: 'image/png' },
  },
];

const IMAGE_TYPE = REL + 'image';

const caption = (id: number, name: string, y: number, lines: readonly string[]): string =>
  shape({
    id,
    name,
    x: 685800,
    y,
    cx: 10820400,
    cy: 1371600,
    textBody: txBody({
      bodyPr: '<a:bodyPr wrap="square" anchor="t"/>',
      paras: lines.map((line) => textLine(line, { sz: 1300 })).join(''),
    }),
  });

const SLIDE_1_RELS: readonly ProbeRel[] = [
  { id: 'rId2', type: REL_MEDIA, target: '../media/media1.wav' },
  { id: 'rId3', type: REL + 'audio', target: '../media/media1.wav' },
  { id: 'rId4', type: IMAGE_TYPE, target: '../media/image1.png' },
];

const SLIDE_2_RELS: readonly ProbeRel[] = [
  { id: 'rId2', type: REL_MEDIA, target: '../media/media2.mp4' },
  { id: 'rId3', type: REL + 'video', target: '../media/media2.mp4' },
  { id: 'rId4', type: IMAGE_TYPE, target: '../media/image1.png' },
  // The linked one. Same element, same attribute, and only `TargetMode` says so.
  {
    id: 'rId5',
    type: REL + 'video',
    target: 'https://example.invalid/corpus/linked-clip.mp4',
    external: true,
  },
];

export const a24Media: ProbeDeck = {
  id: 'a24-media',
  title: 'PPTX Studio corpus: a24 media',
  description:
    'Embedded audio, embedded video and a linked video, written the way PowerPoint 16.0.20326 ' +
    'writes them: one media part related to twice from the same slide by two relationship types, ' +
    'a:audioFile and a:videoFile carrying @r:link even when the target is internal, and an ' +
    'a:hlinkClick whose r:id is the empty string. Closes two gaps a16 and a17 declared - ' +
    'p:transition/p:sndAc/p:stSnd with an embedded WAV, and p:audio and p:video inside p:tnLst. ' +
    'The WAV is a real 440 Hz tone; the MP4 is a valid container with no tracks, which PowerPoint ' +
    'will not itself insert.',
  features: {
    // 6 chassis + three captions.
    shape: 9,
    placeholder: 6,
    picture: 3,
    // Three captions and three media pictures, all plain rectangles.
    presetGeom: 6,
    gradientFill: 2,
    audio: 1,
    video: 2,
    // Only the two embedded parts. A linked clip has no local bytes to name.
    media: 2,
    // One per media shape, every one of them empty.
    hyperlink: 3,
    animation: 2,
    transition: 1,
  },
  build: () => ({
    title: 'PPTX Studio corpus: a24 media',
    parts: PARTS,
    slides: [
      {
        title: 'a24 — embedded audio, and the same part related to twice',
        rels: SLIDE_1_RELS,
        body:
          mediaPic({
            id: 10,
            name: 'probe',
            descr: 'A quarter second of 440 Hz, authored by the corpus generator',
            element: 'audioFile',
            linkRelId: 'rId3',
            mediaRelId: 'rId2',
            posterRelId: 'rId4',
            x: 685800,
            y: 1600200,
          }) +
          caption(11, 'Slide 1 caption', 4114800, [
            'rId2 is …/2007/relationships/media, named by p14:media r:embed.',
            'rId3 is …/2006/relationships/audio, named by a:audioFile r:link.',
            'Both target ppt/media/media1.wav. Export must not drop either.',
          ]),
        tail: mediaTiming([10], mediaNode('audio', 7, 10)),
      },
      {
        title: 'a24 — embedded video beside a linked one',
        rels: SLIDE_2_RELS,
        body:
          mediaPic({
            id: 10,
            name: 'Embedded clip',
            descr: 'An ISO-BMFF container with no tracks; the markup is the probe',
            element: 'videoFile',
            linkRelId: 'rId3',
            mediaRelId: 'rId2',
            posterRelId: 'rId4',
            x: 685800,
            y: 1600200,
          }) +
          mediaPic({
            id: 11,
            name: 'Linked clip',
            descr: 'TargetMode="External"; no p14:media, because there are no local bytes',
            element: 'videoFile',
            linkRelId: 'rId5',
            posterRelId: 'rId4',
            x: 4114800,
            y: 1600200,
          }) +
          caption(12, 'Slide 2 caption', 4114800, [
            'Both shapes say r:link. Only the relationship says which is linked.',
            'The linked target is under example.invalid, which RFC 2606 reserves,',
            'so nothing here can resolve to a real host.',
          ]),
        tail: mediaTiming([10, 11], mediaNode('video', 12, 10) + mediaNode('video', 13, 11)),
      },
      {
        title: 'a24 — a transition with a start sound',
        rels: [{ id: 'rId2', type: REL + 'audio', target: '../media/audio1.wav' }],
        body: caption(10, 'Slide 3 caption', 1600200, [
          'p:transition holds p:sndAc, and p:sndAc holds a choice:',
          'p:stSnd, which carries a p:snd with an r:embed and a @name,',
          'or p:endSnd, which is CT_Empty and carries nothing at all.',
          'a16-transitions wrote the second and could not write the first.',
          '@name is the file name the sound came from, not the part name.',
        ]),
        tail:
          '<p:transition spd="slow"><p:fade thruBlk="1"/>' +
          '<p:sndAc><p:stSnd><p:snd r:embed="rId2" name="probe.wav"/></p:stSnd></p:sndAc>' +
          '</p:transition>',
      },
    ],
  }),
};
