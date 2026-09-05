import { describe, expect, it } from 'vitest';
import { CONTENT_TYPE, CONTENT_TYPES_PART, REL_TYPE } from '../constants.js';
import { isOpcError } from '../errors.js';
import { PartStore } from './part-store.js';
import {
  buildPackage,
  contentTypesXml,
  MINIMAL_PRESENTATION,
  relsXml,
  relXml,
} from '../testing/build-package.js';
import { buildZip } from '../testing/build-zip.js';
import { readZip } from '../zip/zip-reader.js';
import { deflatedEntry, passthroughEntry, writeZip } from '../zip/zip-writer.js';

const encoder = new TextEncoder();
const decoder = new TextDecoder();

function codeOf(fn: () => unknown): string {
  try {
    fn();
  } catch (error) {
    if (isOpcError(error)) return error.code;
    return 'NOT_AN_OPC_ERROR: ' + (error instanceof Error ? error.constructor.name : typeof error);
  }
  return 'DID_NOT_THROW';
}

/** A package with a slide, its layout, and the relationships between them. */
function deck(): Uint8Array {
  return buildPackage({
    contentTypesXml: contentTypesXml(
      '<Default Extension="rels" ContentType="' + CONTENT_TYPE.relationships + '"/>',
      '<Default Extension="xml" ContentType="' + CONTENT_TYPE.xml + '"/>',
      '<Default Extension="png" ContentType="' + CONTENT_TYPE.png + '"/>',
      '<Override PartName="/ppt/presentation.xml" ContentType="' +
        CONTENT_TYPE.presentation +
        '"/>',
      '<Override PartName="/ppt/slides/slide1.xml" ContentType="' + CONTENT_TYPE.slide + '"/>',
      '<Override PartName="/ppt/slideLayouts/slideLayout1.xml" ContentType="' +
        CONTENT_TYPE.slideLayout +
        '"/>',
    ),
    parts: [
      {
        name: 'ppt/_rels/presentation.xml.rels',
        content: relsXml(relXml('rId1', REL_TYPE.slide, 'slides/slide1.xml')),
      },
      { name: 'ppt/slides/slide1.xml', content: '<p:sld/>' },
      {
        name: 'ppt/slides/_rels/slide1.xml.rels',
        content: relsXml(
          relXml('rId1', REL_TYPE.slideLayout, '../slideLayouts/slideLayout1.xml'),
          relXml('rId2', REL_TYPE.image, '../media/image1.png'),
        ),
      },
      { name: 'ppt/slideLayouts/slideLayout1.xml', content: '<p:sldLayout/>' },
      { name: 'ppt/media/image1.png', content: new Uint8Array([0x89, 0x50, 0x4e, 0x47]) },
    ],
  });
}

describe('opening a package', () => {
  it('finds every part, in archive order, with its content type', () => {
    const store = PartStore.open(deck());
    expect(store.partNames[0]).toBe('/_rels/.rels');
    expect(store.size).toBe(7);
    expect(store.contentTypeOf('/ppt/slides/slide1.xml')).toBe(CONTENT_TYPE.slide);
    expect(store.contentTypeOf('/ppt/media/image1.png')).toBe(CONTENT_TYPE.png);
    expect(store.contentTypeOf('/ppt/slides/_rels/slide1.xml.rels')).toBe(
      CONTENT_TYPE.relationships,
    );
  });

  it('does not treat the content-type stream as a part', () => {
    const store = PartStore.open(deck());
    expect(store.partNames.some((n) => n.includes('Content_Types'))).toBe(false);
    expect(store.has(CONTENT_TYPES_PART)).toBe(false);
  });

  it('reads a part’s bytes and reports its size without reading it', () => {
    const store = PartStore.open(deck());
    expect(decoder.decode(store.read('/ppt/slides/slide1.xml'))).toBe('<p:sld/>');
    expect(store.info('/ppt/media/image1.png')).toMatchObject({
      contentType: CONTENT_TYPE.png,
      size: 4,
      fromArchive: true,
    });
  });

  it('looks a part up case-insensitively, as OPC requires', () => {
    const store = PartStore.open(deck());
    expect(store.has('/PPT/Slides/Slide1.XML')).toBe(true);
    expect(decoder.decode(store.read('/PPT/Slides/Slide1.XML'))).toBe('<p:sld/>');
  });

  it('refuses an archive with no content-type stream', () => {
    expect(codeOf(() => PartStore.open(buildPackage({ drop: [CONTENT_TYPES_PART] })))).toBe(
      'ERR_MISSING_PACKAGE_PART',
    );
  });

  it('refuses a part name the grammar rejects', () => {
    expect(
      codeOf(() =>
        PartStore.open(
          buildZip([
            { name: '[Content_Types].xml', data: encoder.encode('<Types/>') },
            { name: '../evil.xml', data: encoder.encode('x') },
          ]),
        ),
      ),
    ).toBe('ERR_INVALID_PART_NAME');
  });

  it('drops directory entries, which are not parts and carry no meaning', () => {
    const store = PartStore.open(
      buildPackage({ parts: [{ name: 'ppt/', content: new Uint8Array(0) }] }),
    );
    expect(store.partNames.some((n) => n.endsWith('/'))).toBe(false);
    expect(readZip(store.write()).entries.some((e) => e.isDirectory)).toBe(false);
  });

  it('finds a lowercase content-type stream, and writes it back canonically', () => {
    // PowerPoint opens a package whose stream is named `[content_types].xml`,
    // so refusing one would fail a file that works everywhere else. ZIP entry
    // names are case-sensitive, so this is a real divergence rather than a
    // formality: we read either spelling and write the one the spec fixes.
    const lower = buildPackage().slice();
    const bytes = readZip(lower);
    const rebuilt = writeZip(
      bytes.entries.map((e) => {
        const input = passthroughEntry(bytes, e);
        return e.name === CONTENT_TYPES_PART ? { ...input, name: '[content_types].xml' } : input;
      }),
    );
    const store = PartStore.open(rebuilt);
    expect(store.contentTypeOf('/ppt/presentation.xml')).toBe(CONTENT_TYPE.presentation);
    expect(readZip(store.write()).has(CONTENT_TYPES_PART)).toBe(true);
    expect(readZip(store.write()).has('[content_types].xml')).toBe(false);
  });

  it('refuses two entries that resolve to the same part name', () => {
    const bytes = writeZip([
      deflatedEntry(CONTENT_TYPES_PART, encoder.encode('<Types/>')),
      deflatedEntry('ppt/a.xml', encoder.encode('1')),
      deflatedEntry('ppt/A.xml', encoder.encode('2')),
    ]);
    expect(codeOf(() => PartStore.open(bytes))).toBe('ERR_DUPLICATE_ENTRY');
  });
});

describe('walking the relationship graph', () => {
  it('goes root -> presentation -> slide -> layout', () => {
    const store = PartStore.open(deck());
    const root = store.rootRelationships();
    const presentation = root.resolve(root.firstOfType(REL_TYPE.officeDocument)!);
    expect(presentation).toBe('/ppt/presentation.xml');

    const presRels = store.relationships(presentation);
    const slide = presRels.resolve(presRels.firstOfType(REL_TYPE.slide)!);
    expect(slide).toBe('/ppt/slides/slide1.xml');

    const slideRels = store.relationships(slide);
    expect(slideRels.resolve(slideRels.firstOfType(REL_TYPE.slideLayout)!)).toBe(
      '/ppt/slideLayouts/slideLayout1.xml',
    );
    expect(slideRels.resolve(slideRels.firstOfType(REL_TYPE.image)!)).toBe('/ppt/media/image1.png');
  });

  it('returns the same object twice, so a mutation is not lost', () => {
    const store = PartStore.open(deck());
    const once = store.relationships('/ppt/slides/slide1.xml');
    once.add(REL_TYPE.image, '../media/image2.png');
    expect(store.relationships('/ppt/slides/slide1.xml').size).toBe(3);
    expect(store.relationships('/ppt/slides/slide1.xml')).toBe(once);
  });

  it('gives an empty collection for a part with no .rels', () => {
    const store = PartStore.open(deck());
    expect(store.relationships('/ppt/slideLayouts/slideLayout1.xml').size).toBe(0);
  });
});

describe('the round trip', () => {
  it('comes back with the same entry names in the same order', () => {
    const original = deck();
    const before = readZip(original);
    const after = readZip(PartStore.open(original).write());
    expect(after.entries.map((e) => e.name)).toEqual(before.entries.map((e) => e.name));
  });

  it('comes back byte-for-byte identical when nothing was touched', () => {
    const original = deck();
    expect([...PartStore.open(original).write()]).toEqual([...original]);
  });

  it('preserves an entry order that is not the one Office writes', () => {
    // Verified against the installed PowerPoint: a real deck rebuilt with the
    // content-type stream written last opens with every slide intact. So the
    // ordering is a convention, and preservation beats normalisation.
    const original = buildPackage({
      order: (names) => [...names.filter((n) => n !== CONTENT_TYPES_PART), CONTENT_TYPES_PART],
    });
    const after = readZip(PartStore.open(original).write());
    expect(after.entries[after.entries.length - 1]!.name).toBe(CONTENT_TYPES_PART);
    expect([...PartStore.open(original).write()]).toEqual([...original]);
  });

  it('normalises the order on request', () => {
    const original = buildPackage({
      order: (names) => [...names.filter((n) => n !== CONTENT_TYPES_PART), CONTENT_TYPES_PART],
    });
    const after = readZip(PartStore.open(original).write({ normalizeEntryOrder: true }));
    expect(after.entries.slice(0, 2).map((e) => e.name)).toEqual([
      CONTENT_TYPES_PART,
      '_rels/.rels',
    ]);
  });

  it('rewrites only what changed and passes the rest through untouched', () => {
    const original = deck();
    const before = readZip(original);
    const store = PartStore.open(original);
    store.replacePart('/ppt/slides/slide1.xml', encoder.encode('<p:sld changed="1"/>'));
    const after = readZip(store.write());

    expect(decoder.decode(after.readByName('ppt/slides/slide1.xml'))).toBe('<p:sld changed="1"/>');
    for (const name of ['ppt/presentation.xml', 'ppt/media/image1.png', CONTENT_TYPES_PART]) {
      expect([...after.raw(after.get(name)!)], name).toEqual([...before.raw(before.get(name)!)]);
    }
  });
});

describe('changing the package', () => {
  it('adds a part with its content type and a relationship to it', () => {
    const store = PartStore.open(deck());
    store.addPart('/ppt/slides/slide2.xml', CONTENT_TYPE.slide, encoder.encode('<p:sld/>'));
    const rel = store
      .relationships('/ppt/presentation.xml')
      .addTo(REL_TYPE.slide, '/ppt/slides/slide2.xml');
    expect(rel.id).toBe('rId2');
    expect(rel.target).toBe('slides/slide2.xml');

    const back = PartStore.open(store.write());
    expect(back.contentTypeOf('/ppt/slides/slide2.xml')).toBe(CONTENT_TYPE.slide);
    const rels = back.relationships('/ppt/presentation.xml');
    expect(rels.resolve(rels.byId('rId2')!)).toBe('/ppt/slides/slide2.xml');
  });

  it('creates a .rels part for a part that had none', () => {
    const store = PartStore.open(deck());
    store
      .relationships('/ppt/slideLayouts/slideLayout1.xml')
      .addTo(REL_TYPE.slideMaster, '/ppt/slideMasters/slideMaster1.xml');
    store.addPart(
      '/ppt/slideMasters/slideMaster1.xml',
      CONTENT_TYPE.slideMaster,
      encoder.encode('<x/>'),
    );

    const back = PartStore.open(store.write());
    expect(back.has('/ppt/slideLayouts/_rels/slideLayout1.xml.rels')).toBe(true);
    expect(back.relationships('/ppt/slideLayouts/slideLayout1.xml').size).toBe(1);
  });

  it('removes a part, its .rels and its Override together', () => {
    const store = PartStore.open(deck());
    store.relationships('/ppt/presentation.xml').remove('rId1');
    expect(store.removePart('/ppt/slides/slide1.xml')).toBe(true);
    expect(store.has('/ppt/slides/_rels/slide1.xml.rels')).toBe(false);
    expect(store.contentTypes.for('/ppt/slides/slide1.xml')).toBe(CONTENT_TYPE.xml);

    const back = PartStore.open(store.write());
    expect(back.has('/ppt/slides/slide1.xml')).toBe(false);
    expect(back.has('/ppt/media/image1.png')).toBe(true);
  });

  it('drops a .rels part once its last relationship goes', () => {
    const store = PartStore.open(deck());
    const rels = store.relationships('/ppt/slides/slide1.xml');
    rels.remove('rId1');
    rels.remove('rId2');
    const back = readZip(store.write());
    expect(back.has('ppt/slides/_rels/slide1.xml.rels')).toBe(false);
  });

  it('refuses to add a part that is already there, or one that collides', () => {
    const store = PartStore.open(deck());
    expect(
      codeOf(() => store.addPart('/ppt/slides/slide1.xml', CONTENT_TYPE.slide, new Uint8Array())),
    ).toBe('ERR_PART_EXISTS');
    expect(
      codeOf(() => store.addPart('/PPT/Slides/Slide1.xml', CONTENT_TYPE.slide, new Uint8Array())),
    ).toBe('ERR_PART_EXISTS');
    // M1.11: one name cannot be both a part and a folder.
    expect(
      codeOf(() =>
        store.addPart('/ppt/slides/slide1.xml/extra.xml', CONTENT_TYPE.slide, new Uint8Array()),
      ),
    ).toBe('ERR_INVALID_PART_NAME');
  });

  it('refuses to read or replace a part that is not there', () => {
    const store = PartStore.open(deck());
    expect(codeOf(() => store.read('/ppt/slides/slide9.xml'))).toBe('ERR_PART_NOT_FOUND');
    expect(codeOf(() => store.replacePart('/nope.xml', new Uint8Array()))).toBe(
      'ERR_PART_NOT_FOUND',
    );
    expect(store.removePart('/nope.xml')).toBe(false);
  });

  it('builds a package from nothing', () => {
    const store = PartStore.create();
    store.contentTypes.setDefault('rels', CONTENT_TYPE.relationships);
    store.addPart(
      '/ppt/presentation.xml',
      CONTENT_TYPE.presentation,
      encoder.encode(MINIMAL_PRESENTATION),
    );
    store.rootRelationships().addTo(REL_TYPE.officeDocument, '/ppt/presentation.xml');

    const back = PartStore.open(store.write());
    const root = back.rootRelationships();
    expect(root.resolve(root.firstOfType(REL_TYPE.officeDocument)!)).toBe('/ppt/presentation.xml');
    expect(readZip(store.write()).entries[0]!.name).toBe(CONTENT_TYPES_PART);
  });
});

describe('what write refuses to emit', () => {
  it('refuses a part with no content type', () => {
    // PowerPoint refuses this package with its own error code, 0x80CB8002 -
    // a different one from the generic corruption it reports for a broken
    // relationship graph. Verified by deleting the Default from a real deck.
    const store = PartStore.open(
      buildPackage({ parts: [{ name: 'ppt/media/image1.png', content: new Uint8Array([1]) }] }),
    );
    expect(codeOf(() => store.write())).toBe('ERR_MISSING_CONTENT_TYPE');
  });

  it('refuses a dangling relationship when we broke it by deleting the target', () => {
    const store = PartStore.open(deck());
    store.removePart('/ppt/media/image1.png');
    expect(codeOf(() => store.write())).toBe('ERR_DANGLING_RELATIONSHIP');
  });

  it('preserves a relationship that was already dangling when we opened the file', () => {
    // Verified against the installed PowerPoint: an extra image relationship
    // pointing at a part that is not in the package opens with every slide and
    // picture intact. Refusing to export such a deck would make a file that
    // works everywhere else impossible to save, which is the opposite of what
    // this layer is for.
    const stale = buildPackage({
      parts: [
        {
          name: 'ppt/_rels/presentation.xml.rels',
          content: relsXml(
            relXml('rId1', REL_TYPE.image, 'media/gone.png'),
            relXml('rId2', REL_TYPE.theme, 'theme/alsogone.xml'),
          ),
        },
      ],
    });
    const store = PartStore.open(stale);
    expect(() => store.write()).not.toThrow();
    expect([...PartStore.open(stale).write()]).toEqual([...stale]);

    // Preserved, but not silent.
    const reported = store.danglingRelationships();
    expect(reported.map((d) => d.relationship.id).sort()).toEqual(['rId1', 'rId2']);
    expect(reported[0]!.source).toBe('/ppt/_rels/presentation.xml.rels');
    expect(reported.every((d) => d.relationship.origin === 'archive')).toBe(true);
  });

  it('still refuses a relationship we added ourselves that points at nothing', () => {
    const store = PartStore.open(deck());
    const rel = store
      .relationships('/ppt/presentation.xml')
      .addTo(REL_TYPE.slide, '/ppt/slides/slide9.xml');
    expect(rel.origin).toBe('added');
    expect(codeOf(() => store.write())).toBe('ERR_DANGLING_RELATIONSHIP');
  });

  it('refuses a package with no root relationships', () => {
    const store = PartStore.open(deck());
    store.removePart('/_rels/.rels');
    expect(codeOf(() => store.write())).toBe('ERR_MISSING_PACKAGE_PART');
  });

  it('refuses a relationship added to a part that does not exist', () => {
    const store = PartStore.open(deck());
    store.relationships('/ppt/presentation.xml').addTo(REL_TYPE.slide, '/ppt/slides/slide9.xml');
    expect(codeOf(() => store.write())).toBe('ERR_DANGLING_RELATIONSHIP');
  });

  it('lets an external target through without looking for a part', () => {
    const store = PartStore.open(deck());
    store
      .relationships('/ppt/slides/slide1.xml')
      .add(REL_TYPE.hyperlink, 'https://example.com/', 'External');
    const back = PartStore.open(store.write());
    const rels = back.relationships('/ppt/slides/slide1.xml');
    expect(rels.firstOfType(REL_TYPE.hyperlink)?.target).toBe('https://example.com/');
  });

  it('reads a part name with a literal space but refuses to write one', () => {
    // ADR 0002 rated a non-`pchar` character a warning on the reasoning that
    // `my image.png` is harmless. It is not: PowerPoint refuses a package
    // containing that name (0x808D1001) even when the part is an orphan and
    // fully content-typed — while accepting `%20`, the same name spelled
    // properly. So the grammar stays lenient and the writer does not.
    const withSpace = buildPackage({
      contentTypesXml: contentTypesXml(
        '<Default Extension="rels" ContentType="' + CONTENT_TYPE.relationships + '"/>',
        '<Default Extension="xml" ContentType="' + CONTENT_TYPE.xml + '"/>',
        '<Default Extension="png" ContentType="' + CONTENT_TYPE.png + '"/>',
        '<Override PartName="/ppt/presentation.xml" ContentType="' +
          CONTENT_TYPE.presentation +
          '"/>',
      ),
      parts: [{ name: 'ppt/media/my image.png', content: new Uint8Array([1]) }],
    });
    const store = PartStore.open(withSpace);
    expect(store.has('/ppt/media/my image.png')).toBe(true);
    expect(codeOf(() => store.write())).toBe('ERR_INVALID_PART_NAME');

    // The percent-encoded spelling writes fine.
    const encoded = PartStore.open(withSpace);
    encoded.removePart('/ppt/media/my image.png');
    encoded.addPart('/ppt/media/my%20image.png', CONTENT_TYPE.png, new Uint8Array([1]));
    expect(PartStore.open(encoded.write()).has('/ppt/media/my%20image.png')).toBe(true);
  });

  it('refuses a relationship that targets a relationship part', () => {
    // The spec says implementers "shall treat any such relationship as
    // invalid", and PowerPoint agrees by refusing the package.
    const store = PartStore.open(deck());
    store
      .relationships('/ppt/presentation.xml')
      .addTo(REL_TYPE.slide, '/ppt/slides/_rels/slide1.xml.rels');
    expect(codeOf(() => store.write())).toBe('ERR_INVALID_RELATIONSHIP_TARGET');
  });

  it('tolerates an orphan part, because orphan is harmless and dangling is not', () => {
    const store = PartStore.open(deck());
    store.addPart('/ppt/unused.xml', CONTENT_TYPE.xml, encoder.encode('<x/>'));
    expect(PartStore.open(store.write()).has('/ppt/unused.xml')).toBe(true);
  });
});

describe('what the store says it is about to rewrite', () => {
  it('names nothing when nothing was edited', () => {
    expect(PartStore.open(deck()).rewrittenParts()).toEqual([]);
  });

  it('names a replaced part and an added one', () => {
    const store = PartStore.open(deck());
    store.replacePart('/ppt/slides/slide1.xml', encoder.encode('<p:sld edited="1"/>'));
    store.addPart('/ppt/extra.xml', CONTENT_TYPE.xml, encoder.encode('<x/>'));

    expect(store.rewrittenParts()).toEqual(['/ppt/slides/slide1.xml', '/ppt/extra.xml']);
  });

  it('names a relationship collection edited in place', () => {
    // `PartInfo.fromArchive` cannot see this one - the parsed collection is the
    // edit and `replacePart` was never called - and `write` re-emits it anyway.
    const store = PartStore.open(deck());
    store.relationships('/ppt/slides/slide1.xml').remove('rId2');

    expect(store.info('/ppt/slides/_rels/slide1.xml.rels')?.fromArchive).toBe(true);
    expect(store.rewrittenParts()).toEqual(['/ppt/slides/_rels/slide1.xml.rels']);
  });
});

describe('materializing relationships', () => {
  it('makes an in-place edit visible to everything that reads bytes', () => {
    // The bug this method exists for: until it runs, `read` returns the `.rels`
    // as it arrived, so anything checking the markup - the validator does,
    // deliberately, because its own parser refuses the malformed ids three of
    // the rules report - is looking at a package that is not the one about to
    // be written.
    const store = PartStore.open(deck());
    store.relationships('/ppt/slides/slide1.xml').remove('rId2');
    expect(decoder.decode(store.read('/ppt/slides/_rels/slide1.xml.rels'))).toContain('rId2');

    store.materializeRelationships();
    expect(decoder.decode(store.read('/ppt/slides/_rels/slide1.xml.rels'))).not.toContain('rId2');
    expect(store.info('/ppt/slides/_rels/slide1.xml.rels')?.fromArchive).toBe(false);
  });

  it('gives a collection that grew from nothing a part of its own', () => {
    // Before this, such a `.rels` was not in `partNames` at all, so nothing
    // downstream could see it - a relationship part created this session was
    // the one part of an export that went unchecked.
    const store = PartStore.open(deck());
    expect(store.has('/ppt/slideLayouts/_rels/slideLayout1.xml.rels')).toBe(false);
    store
      .relationships('/ppt/slideLayouts/slideLayout1.xml')
      .addTo(REL_TYPE.image, '/ppt/media/image1.png');

    store.materializeRelationships();
    expect(store.partNames).toContain('/ppt/slideLayouts/_rels/slideLayout1.xml.rels');
    expect(store.contentTypes.for('/ppt/slideLayouts/_rels/slideLayout1.xml.rels')).toBe(
      CONTENT_TYPE.relationships,
    );
  });

  it('drops a collection emptied of everything rather than writing an empty one', () => {
    const store = PartStore.open(deck());
    const rels = store.relationships('/ppt/slides/slide1.xml');
    for (const rel of [...rels.all]) rels.remove(rel.id);

    store.materializeRelationships();
    expect(store.has('/ppt/slides/_rels/slide1.xml.rels')).toBe(false);
  });

  it('is idempotent, and leaves write with nothing left to do', () => {
    const store = PartStore.open(deck());
    store.relationships('/ppt/slides/slide1.xml').remove('rId2');

    store.materializeRelationships();
    const once = store.write();
    store.materializeRelationships();
    expect(store.write()).toEqual(once);
  });
});
