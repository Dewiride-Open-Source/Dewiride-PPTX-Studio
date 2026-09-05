import { CONTENT_TYPE, PartStore, REL_TYPE } from '@pptx-studio/opc';
import { NS } from '@pptx-studio/xml';
import { describe, expect, it } from 'vitest';
import { comparePackages, roundTripPackage, summarizeRoundTrip } from './roundtrip.js';
import { fixtureBytes, MAIN_PART } from '../testing/package.js';

/**
 * A comparator is two tests, and only one of them is obvious.
 *
 * The obvious one is that it agrees when the packages agree. The one that
 * decides whether it is worth having is that it *disagrees* when they differ,
 * and that it does not disagree over things that are not differences - because
 * a comparator with either fault ends the same way. One that cries wolf gets
 * turned off; one that never does was never a check.
 *
 * So every block below comes in a pair: what must be absorbed, and what must be
 * caught.
 */

const enc = (text: string): Uint8Array => new TextEncoder().encode(text);
const IMAGE_A = '/media/a.png';
const IMAGE_B = '/media/b.png';

interface Edge {
  readonly id: string;
  readonly type: string;
  readonly target: string;
  readonly external?: boolean;
}

/**
 * A package with one XML part, some media, and relationships whose ids the test
 * chooses.
 *
 * Choosing the ids is the whole point: two packages that differ only in how
 * they numbered their relationships are the same package, and the only way to
 * say so in a test is to number them differently on purpose.
 */
function build(
  body: string,
  edges: readonly Edge[] = [],
  media: readonly string[] = [],
): PartStore {
  const store = PartStore.create();
  store.addPart(MAIN_PART, CONTENT_TYPE.xml, enc('<doc xmlns:r="' + NS.r + '">' + body + '</doc>'));
  store.rootRelationships().addWithId('rId1', REL_TYPE.officeDocument, 'doc.xml');
  for (const part of media) store.addPart(part, CONTENT_TYPE.png, enc('pixels of ' + part));
  const rels = store.relationships(MAIN_PART);
  for (const edge of edges) {
    rels.addWithId(
      edge.id,
      edge.type,
      edge.target,
      edge.external === true ? 'External' : 'Internal',
    );
  }
  // Written and re-opened, rather than handed over as it stands. A package that
  // has been through an archive is what every real caller compares, and it is
  // the only state in which its relationship parts exist as parts.
  return PartStore.open(store.write());
}

describe('what two archives are allowed to disagree about', () => {
  it('calls a package identical to itself identical', () => {
    const report = comparePackages(build('<a x="1"/>'), build('<a x="1"/>'));
    expect(report.differences).toEqual([]);
    expect(report.ok).toBe(true);
  });

  it('sees past attribute order', () => {
    const report = comparePackages(build('<a x="1" y="2"/>'), build('<a y="2" x="1"/>'));
    expect(report.ok).toBe(true);
  });

  it('sees past a relationship renumbering when the markup follows', () => {
    // The claim the plan makes in one clause - "relationship-graph isomorphism
    // with rIds treated as opaque labels" - as an assertion. Both packages say
    // the picture uses image A; they disagree only about what that edge is
    // called, and PowerPoint renumbers on every save.
    const before = build(
      '<pic r:embed="rId2"/>',
      [{ id: 'rId2', type: REL_TYPE.image, target: 'media/a.png' }],
      [IMAGE_A],
    );
    const after = build(
      '<pic r:embed="rId7"/>',
      [{ id: 'rId7', type: REL_TYPE.image, target: 'media/a.png' }],
      [IMAGE_A],
    );
    const report = comparePackages(before, after);
    expect(report.differences).toEqual([]);
    expect(report.relabelled).toHaveLength(1);
    const [moved] = report.relabelled;
    expect(moved?.source).toBe(MAIN_PART);
    expect(moved?.from).toBe('rId2');
    expect(moved?.to).toBe('rId7');
    // The target is carried so a report can say *why* two ids are the same one.
    expect(moved?.target).toContain('/media/a.png');
  });

  it('gives a relationship part the same digest either side of a renumbering', () => {
    const edges = (id: string): Edge[] => [{ id, type: REL_TYPE.image, target: 'media/a.png' }];
    const before = comparePackages(
      build('<pic r:embed="rId2"/>', edges('rId2'), [IMAGE_A]),
      build('<pic r:embed="rId7"/>', edges('rId7'), [IMAGE_A]),
    );
    const rels = before.parts.find(
      (part) => part.how === 'relationships' && part.part.endsWith('doc.xml.rels'),
    );
    expect(rels?.digest).toBe(rels?.writtenDigest);
    expect(rels?.same).toBe(true);
  });

  it('sees past an absolute target written where a relative one was', () => {
    const report = comparePackages(
      build(
        '<pic r:embed="rId2"/>',
        [{ id: 'rId2', type: REL_TYPE.image, target: 'media/a.png' }],
        [IMAGE_A],
      ),
      build(
        '<pic r:embed="rId2"/>',
        [{ id: 'rId2', type: REL_TYPE.image, target: '/media/a.png' }],
        [IMAGE_A],
      ),
    );
    expect(report.ok).toBe(true);
  });

  it('never compares the content-type stream as a document', () => {
    // It is a map with several spellings, and the resolved type of every part
    // is compared instead. A comparator that diffed the stream would report a
    // `Default` that became an `Override` as a change to the deck.
    const before = build('<a/>');
    expect(before.partNames.some((name) => name.includes('Content_Types'))).toBe(false);
    const report = comparePackages(before, build('<a/>'));
    expect(report.parts.map((part) => part.part)).not.toContain('[Content_Types].xml');
  });
});

describe('what has to be caught', () => {
  it('catches a changed attribute value, and says where', () => {
    const report = comparePackages(build('<a x="1"/>'), build('<a x="2"/>'));
    expect(report.ok).toBe(false);
    const difference = report.differences.find((entry) => entry.kind === 'xml');
    expect(difference?.part).toBe(MAIN_PART);
    expect(difference?.detail).toContain('x="1"');
    expect(difference?.detail).toContain('x="2"');
  });

  it('catches a renumbering the markup did not follow', () => {
    // The case that makes the relabelling worth doing carefully. Both packages
    // have two images and both say `rId1`; they disagree about which image
    // rId1 names, so one shows A and the other shows B. Treating the ids as
    // interchangeable *labels* is what lets this be seen - matched by target,
    // rId1 on the left is rId2 on the right, and the markup then does not line
    // up.
    const before = build(
      '<pic r:embed="rId1"/>',
      [
        { id: 'rId1', type: REL_TYPE.image, target: 'media/a.png' },
        { id: 'rId2', type: REL_TYPE.image, target: 'media/b.png' },
      ],
      [IMAGE_A, IMAGE_B],
    );
    const after = build(
      '<pic r:embed="rId1"/>',
      [
        { id: 'rId1', type: REL_TYPE.image, target: 'media/b.png' },
        { id: 'rId2', type: REL_TYPE.image, target: 'media/a.png' },
      ],
      [IMAGE_A, IMAGE_B],
    );
    const report = comparePackages(before, after);
    expect(report.ok).toBe(false);
    expect(report.differences.map((entry) => entry.kind)).toContain('xml');
  });

  it('catches a relationship that is gone', () => {
    const report = comparePackages(
      build('<a/>', [{ id: 'rId2', type: REL_TYPE.image, target: 'media/a.png' }], [IMAGE_A]),
      build('<a/>', [], [IMAGE_A]),
    );
    const difference = report.differences.find((entry) => entry.kind === 'relationship');
    expect(difference?.part).toBe(MAIN_PART);
    expect(difference?.detail).toContain('rId2');
    expect(difference?.detail).toContain('is in the original and not in what was written');
  });

  it('catches a relationship that appeared', () => {
    const report = comparePackages(
      build('<a/>', [], [IMAGE_A]),
      build('<a/>', [{ id: 'rId2', type: REL_TYPE.image, target: 'media/a.png' }], [IMAGE_A]),
    );
    const difference = report.differences.find((entry) => entry.kind === 'relationship');
    expect(difference?.detail).toContain('was written and is not in the original');
  });

  it('catches a relationship whose type changed', () => {
    const report = comparePackages(
      build('<a/>', [{ id: 'rId2', type: REL_TYPE.image, target: 'media/a.png' }], [IMAGE_A]),
      build('<a/>', [{ id: 'rId2', type: REL_TYPE.thumbnail, target: 'media/a.png' }], [IMAGE_A]),
    );
    expect(report.differences.filter((entry) => entry.kind === 'relationship')).toHaveLength(2);
  });

  it('catches an external target that changed, without normalising the URI', () => {
    const external = (target: string): Edge[] => [
      { id: 'rId2', type: REL_TYPE.hyperlink, target, external: true },
    ];
    expect(
      comparePackages(
        build('<a/>', external('http://example.com/')),
        build('<a/>', external('http://example.com/')),
      ).ok,
    ).toBe(true);
    expect(
      comparePackages(
        build('<a/>', external('http://example.com/')),
        build('<a/>', external('http://Example.COM/')),
      ).ok,
    ).toBe(false);
  });

  it('catches one byte of a media part, and reports both digests', () => {
    const before = PartStore.create();
    before.addPart(IMAGE_A, CONTENT_TYPE.png, enc('pixels'));
    const after = PartStore.create();
    after.addPart(IMAGE_A, CONTENT_TYPE.png, enc('pixelt'));

    const report = comparePackages(before, after);
    const difference = report.differences.find((entry) => entry.kind === 'binary');
    expect(difference?.part).toBe(IMAGE_A);
    const media = report.parts.find((part) => part.part === IMAGE_A);
    expect(media?.how).toBe('binary');
    expect(media?.same).toBe(false);
    expect(media?.digest).toMatch(/^[0-9a-f]{64}$/);
    expect(media?.writtenDigest).not.toBe(media?.digest);
  });

  it('catches a part that is gone and a part that appeared', () => {
    const report = comparePackages(build('<a/>', [], [IMAGE_A]), build('<a/>', [], [IMAGE_B]));
    const kinds = report.differences.map((entry) => entry.kind);
    expect(kinds).toContain('part-removed');
    expect(kinds).toContain('part-added');
  });

  it('catches a content type that changed on a part nobody edited', () => {
    const before = PartStore.create();
    before.addPart(IMAGE_A, CONTENT_TYPE.png, enc('pixels'));
    const after = PartStore.create();
    after.addPart(IMAGE_A, CONTENT_TYPE.jpeg, enc('pixels'));

    const report = comparePackages(before, after);
    const difference = report.differences.find((entry) => entry.kind === 'content-type');
    expect(difference?.detail).toContain(CONTENT_TYPE.png);
    expect(difference?.detail).toContain(CONTENT_TYPE.jpeg);
  });

  it('does not mistake a difference in case for a difference in content type', () => {
    const before = PartStore.create();
    before.addPart(IMAGE_A, 'image/PNG', enc('pixels'));
    const after = PartStore.create();
    after.addPart(IMAGE_A, 'image/png', enc('pixels'));
    expect(comparePackages(before, after).ok).toBe(true);
  });

  it('reports a part that will not parse rather than calling it equal', () => {
    const before = build('<a/>');
    const after = build('<a/>');
    after.replacePart(MAIN_PART, enc('<doc><unclosed></doc>'));

    const report = comparePackages(before, after);
    const difference = report.differences.find((entry) => entry.kind === 'unreadable');
    expect(difference?.part).toBe(MAIN_PART);
    expect(difference?.detail).toContain('what was written');
  });
});

describe('reading a package and writing it back', () => {
  const bytes = fixtureBytes({ media: ['/ppt/media/image1.png'] });

  it('round-trips with nothing to report', () => {
    const result = roundTripPackage(bytes, { validate: false });
    expect(result.ok).toBe(true);
    expect(result.comparison.differences).toEqual([]);
    expect(result.exported.rewritten).toEqual([]);
  });

  it('round-trips through a different entry order and a different deflate level', () => {
    // The reason the comparison is not byte equality, made falsifiable. Both
    // knobs change the archive and neither changes the document, so a
    // byte-comparing gate would be red here and this one is green.
    const result = roundTripPackage(bytes, {
      validate: false,
      normalizeEntryOrder: true,
      deflateLevel: 1,
    });
    expect(result.ok).toBe(true);
    expect(result.exported.bytes).not.toEqual(bytes);
  });

  it('counts what it compared', () => {
    const result = roundTripPackage(bytes, { validate: false });
    const { xml, binary, relationships, same } = result.comparison.counts;
    expect(xml + binary + relationships).toBe(result.comparison.parts.length);
    expect(same).toBe(result.comparison.parts.length);
    expect(binary).toBe(1);
    expect(relationships).toBe(2);
  });

  it('summarises in one line', () => {
    const result = roundTripPackage(bytes, { validate: false });
    expect(summarizeRoundTrip(result.comparison)).toBe(
      '4/4 parts identical (1 xml, 2 rels, 1 binary)',
    );
  });

  it('reports the difference when the export changed something', () => {
    const opened = roundTripPackage(bytes, { validate: false });
    expect(opened.ok).toBe(true);

    const changed = roundTripPackage(bytes, {
      validate: false,
      verifyPreservation: false,
      prepare: [
        {
          name: 'edit the main part',
          run: (context) => {
            context.store.replacePart(MAIN_PART, enc('<doc edited="yes"/>'));
          },
        },
      ],
    });
    expect(changed.ok).toBe(false);
    expect(changed.comparison.differences[0]?.kind).toBe('xml');
    expect(summarizeRoundTrip(changed.comparison)).toContain('difference(s)');
  });
});
