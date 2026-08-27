import { CONTENT_TYPE, PartStore, REL_TYPE, storedEntry, writeZip } from '@pptx-studio/opc';
import { describe, expect, it } from 'vitest';
import { censusPackage } from './census.js';
import { formatCensus } from './format.js';
import { EXTENSION_NS, OOXML_NS } from './namespaces.js';
import type { PackageCensus } from './types.js';

const encoder = new TextEncoder();

const DECLARATION = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n';
const NS_DECLS =
  'xmlns:a="' + OOXML_NS.a + '" xmlns:r="' + OOXML_NS.r + '" xmlns:p="' + OOXML_NS.p + '"';

interface PartSpec {
  readonly name: string;
  readonly contentType: string;
  readonly content: string | Uint8Array;
  /** Relationships this part declares: `[type, target]`. */
  readonly rels?: readonly (readonly [string, string])[];
  /** Add an `Override` even when a `Default` would already cover it. */
  readonly override?: boolean;
}

/**
 * A package built through the public `opc` API and nothing else.
 *
 * Deliberately not a shared fixture builder. Every test below states the deck
 * it needs in full, because a census is a claim about what is in a file and a
 * shared fixture would make it a claim about what is in a fixture.
 */
function pack(
  parts: readonly PartSpec[],
  rootRels: readonly (readonly [string, string])[],
): Uint8Array {
  const store = PartStore.create();
  store.contentTypes.setDefault('rels', CONTENT_TYPE.relationships);
  store.contentTypes.setDefault('xml', CONTENT_TYPE.xml);
  store.contentTypes.setDefault('png', CONTENT_TYPE.png);

  for (const part of parts) {
    const bytes = typeof part.content === 'string' ? encoder.encode(part.content) : part.content;
    store.addPart(part.name, part.contentType, bytes);
    for (const [type, target] of part.rels ?? []) {
      store.relationships(part.name).addTo(type, target);
    }
  }
  for (const [type, target] of rootRels) {
    store.rootRelationships().addTo(type, target);
  }
  return store.write();
}

/** The smallest package a census will call a presentation. */
function presentationXml(body = '', attributes = ''): string {
  return (
    DECLARATION +
    '<p:presentation ' +
    NS_DECLS +
    (attributes === '' ? '' : ' ' + attributes) +
    '>' +
    '<p:sldSz cx="12192000" cy="6858000" type="screen16x9"/>' +
    '<p:notesSz cx="6858000" cy="9144000"/>' +
    body +
    '</p:presentation>'
  );
}

function slideXml(body: string, extraDecls = ''): string {
  return (
    DECLARATION +
    '<p:sld ' +
    NS_DECLS +
    (extraDecls === '' ? '' : ' ' + extraDecls) +
    '><p:cSld><p:spTree>' +
    '<p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>' +
    '<p:grpSpPr/>' +
    body +
    '</p:spTree></p:cSld></p:sld>'
  );
}

/** A one-slide deck whose slide holds exactly `body`. */
function oneSlide(body: string, extraDecls = '', presentationBody = ''): Uint8Array {
  return pack(
    [
      {
        name: '/ppt/presentation.xml',
        contentType: CONTENT_TYPE.presentation,
        content: presentationXml(presentationBody),
        rels: [[REL_TYPE.slide, '/ppt/slides/slide1.xml']],
      },
      {
        name: '/ppt/slides/slide1.xml',
        contentType: CONTENT_TYPE.slide,
        content: slideXml(body, extraDecls),
      },
    ],
    [[REL_TYPE.officeDocument, '/ppt/presentation.xml']],
  );
}

/** An archive assembled without `PartStore`, for packages it would refuse. */
function rawPack(entries: readonly (readonly [string, string])[]): Uint8Array {
  return writeZip(entries.map(([name, xml]) => storedEntry(name, encoder.encode(xml))));
}

function contentTypes(...extra: readonly string[]): string {
  return (
    DECLARATION +
    '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
    '<Default Extension="rels" ContentType="' +
    CONTENT_TYPE.relationships +
    '"/>' +
    '<Default Extension="xml" ContentType="' +
    CONTENT_TYPE.xml +
    '"/>' +
    '<Override PartName="/ppt/presentation.xml" ContentType="' +
    CONTENT_TYPE.presentation +
    '"/>' +
    extra.join('') +
    '</Types>'
  );
}

function relsDocument(...relationships: readonly string[]): string {
  return (
    DECLARATION +
    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
    relationships.join('') +
    '</Relationships>'
  );
}

function relXml(id: string, type: string, target: string): string {
  return '<Relationship Id="' + id + '" Type="' + type + '" Target="' + target + '"/>';
}

function feature(census: PackageCensus, key: string): number {
  return census.features.find((entry) => entry.key === key)?.count ?? 0;
}

function namespaceOf(census: PackageCensus, uri: string) {
  return census.namespaces.find((entry) => entry.uri === uri);
}

function codes(census: PackageCensus): string[] {
  return census.problems.map((problem) => problem.code);
}

describe('a census describes the package it was given', () => {
  it('reads the presentation structure without opening a slide', () => {
    const census = censusPackage(oneSlide(''));

    expect(census.presentation).not.toBeNull();
    expect(census.presentation?.slideWidthEmu).toBe(12192000);
    expect(census.presentation?.slideHeightEmu).toBe(6858000);
    expect(census.presentation?.slideSizeType).toBe('screen16x9');
    expect(census.presentation?.slides).toBe(1);
    expect(census.archive.entries).toBeGreaterThan(0);
  });

  it('counts parts, and knows which of them are XML', () => {
    const bytes = pack(
      [
        {
          name: '/ppt/presentation.xml',
          contentType: CONTENT_TYPE.presentation,
          content: presentationXml(),
        },
        {
          name: '/ppt/media/image1.png',
          contentType: CONTENT_TYPE.png,
          content: new Uint8Array([0x89, 0x50, 0x4e, 0x47]),
        },
      ],
      [[REL_TYPE.officeDocument, '/ppt/presentation.xml']],
    );
    const census = censusPackage(bytes);

    const media = census.parts.find((part) => part.name === '/ppt/media/image1.png');
    expect(media?.kind).toBe('binary');
    expect(media?.elements).toBeNull();
    expect(media?.contentTypeOrigin).toBe('default');

    const presentation = census.parts.find((part) => part.name === '/ppt/presentation.xml');
    expect(presentation?.kind).toBe('xml');
    expect(presentation?.elements).toBeGreaterThan(0);
    expect(presentation?.contentTypeOrigin).toBe('override');
  });

  it('never puts a relationship part in the scanned set', () => {
    // In production a `.rels` part is read by the OPC layer's own flat-XML
    // reader and never reaches the tokenizer. Counting its elements here would
    // describe a code path that does not exist.
    const census = censusPackage(oneSlide(''));
    expect(census.parts.some((part) => part.name.includes('/_rels/'))).toBe(false);
    expect(census.relationships.total).toBeGreaterThan(0);
  });
});

describe('features are found by namespace, never by prefix', () => {
  it('counts tables, geometry and switches with the phase that renders each', () => {
    const census = censusPackage(
      oneSlide(
        '<p:graphicFrame><a:graphic><a:graphicData>' +
          '<a:tbl><a:tblGrid/></a:tbl>' +
          '</a:graphicData></a:graphic></p:graphicFrame>' +
          '<p:sp><p:spPr><a:prstGeom prst="roundRect"><a:avLst/></a:prstGeom></p:spPr></p:sp>' +
          '<p:sp><p:spPr><a:custGeom><a:pathLst/></a:custGeom></p:spPr></p:sp>',
      ),
    );

    expect(feature(census, 'table')).toBe(1);
    expect(feature(census, 'graphicFrame')).toBe(1);
    expect(feature(census, 'shape')).toBe(2);
    expect(feature(census, 'presetGeom')).toBe(1);
    expect(feature(census, 'customGeom')).toBe(1);
    expect(census.presetGeometry).toEqual([{ name: 'roundRect', count: 1 }]);
    expect(census.features.find((entry) => entry.key === 'table')?.phase).toBe('4.1');
  });

  it('finds the same element under a prefix nobody uses for it', () => {
    // The whole point of resolving through the scope: `zz:tbl` is a table if
    // `zz` is bound to DrawingML, and `a:tbl` is not one if `a` is bound to
    // something else.
    const disguised = pack(
      [
        {
          name: '/ppt/presentation.xml',
          contentType: CONTENT_TYPE.presentation,
          content: presentationXml(),
          rels: [[REL_TYPE.slide, '/ppt/slides/slide1.xml']],
        },
        {
          name: '/ppt/slides/slide1.xml',
          contentType: CONTENT_TYPE.slide,
          content:
            DECLARATION +
            '<zz:sld xmlns:zz="' +
            OOXML_NS.p +
            '" xmlns:qq="' +
            OOXML_NS.a +
            '"><qq:tbl/><zz:sp/></zz:sld>',
        },
      ],
      [[REL_TYPE.officeDocument, '/ppt/presentation.xml']],
    );
    const census = censusPackage(disguised);

    expect(feature(census, 'table')).toBe(1);
    expect(feature(census, 'shape')).toBe(1);
    expect(namespaceOf(census, OOXML_NS.a)?.prefixes).toContain('qq');
  });

  it('is not fooled by a prefix bound to a namespace that is not DrawingML', () => {
    const census = censusPackage(
      pack(
        [
          {
            name: '/ppt/presentation.xml',
            contentType: CONTENT_TYPE.presentation,
            content: presentationXml(),
            rels: [[REL_TYPE.slide, '/ppt/slides/slide1.xml']],
          },
          {
            name: '/ppt/slides/slide1.xml',
            contentType: CONTENT_TYPE.slide,
            content:
              DECLARATION +
              '<a:sld xmlns:a="http://example.invalid/not-drawingml"><a:tbl/></a:sld>',
          },
        ],
        [[REL_TYPE.officeDocument, '/ppt/presentation.xml']],
      ),
    );

    expect(feature(census, 'table')).toBe(0);
    expect(namespaceOf(census, 'http://example.invalid/not-drawingml')?.standard).toBe(false);
    expect(namespaceOf(census, 'http://example.invalid/not-drawingml')?.extension).toBe(false);
  });

  it('records a namespace a deck declares and never uses', () => {
    const census = censusPackage(oneSlide('', 'xmlns:p14="' + EXTENSION_NS.p14 + '"'));
    const p14 = namespaceOf(census, EXTENSION_NS.p14);

    expect(p14).toBeDefined();
    expect(p14?.count).toBe(0);
    expect(p14?.prefixes).toEqual(['p14']);
    expect(p14?.extension).toBe(true);
  });

  it('honours xmlns="" undeclaring the default namespace', () => {
    const census = censusPackage(
      pack(
        [
          {
            name: '/ppt/presentation.xml',
            contentType: CONTENT_TYPE.presentation,
            content: presentationXml(),
            rels: [[REL_TYPE.slide, '/ppt/slides/slide1.xml']],
          },
          {
            name: '/ppt/slides/slide1.xml',
            contentType: CONTENT_TYPE.slide,
            // `tbl` inside the reset scope is in no namespace at all, and a
            // scanner that kept the inherited default would count it.
            content:
              DECLARATION +
              '<sld xmlns="' +
              OOXML_NS.a +
              '"><tbl/><wrap xmlns=""><tbl/></wrap></sld>',
          },
        ],
        [[REL_TYPE.officeDocument, '/ppt/presentation.xml']],
      ),
    );

    expect(feature(census, 'table')).toBe(1);
    expect(namespaceOf(census, '')?.count).toBe(2);
  });
});

describe('markup compatibility is read as prefixes resolved in place', () => {
  it('resolves an unprefixed Requires against the scope it was written in', () => {
    const census = censusPackage(
      oneSlide(
        '<mc:AlternateContent xmlns:mc="' +
          OOXML_NS.mc +
          '" xmlns:p14="' +
          EXTENSION_NS.p14 +
          '">' +
          '<mc:Choice Requires="p14"><p:sp/></mc:Choice>' +
          '<mc:Fallback><p:sp/></mc:Fallback>' +
          '</mc:AlternateContent>',
      ),
    );

    expect(feature(census, 'alternateContent')).toBe(1);
    expect(namespaceOf(census, EXTENSION_NS.p14)?.required).toBe(true);
    expect(namespaceOf(census, EXTENSION_NS.p14)?.ignorable).toBe(false);
  });

  it('reads mc:Ignorable as a list of prefixes, not of URIs', () => {
    const census = censusPackage(
      oneSlide(
        '',
        'xmlns:mc="' +
          OOXML_NS.mc +
          '" xmlns:p14="' +
          EXTENSION_NS.p14 +
          '" xmlns:a14="' +
          EXTENSION_NS.a14 +
          '" mc:Ignorable="p14 a14"',
      ),
    );

    expect(namespaceOf(census, EXTENSION_NS.p14)?.ignorable).toBe(true);
    expect(namespaceOf(census, EXTENSION_NS.a14)?.ignorable).toBe(true);
  });

  it('resolves the same prefix to two different namespaces in two parts', () => {
    const other = 'http://schemas.microsoft.com/office/powerpoint/2007/7/12/main';
    const census = censusPackage(
      pack(
        [
          {
            name: '/ppt/presentation.xml',
            contentType: CONTENT_TYPE.presentation,
            content: presentationXml(),
            rels: [
              [REL_TYPE.slide, '/ppt/slides/slide1.xml'],
              [REL_TYPE.slide, '/ppt/slides/slide2.xml'],
            ],
          },
          {
            name: '/ppt/slides/slide1.xml',
            contentType: CONTENT_TYPE.slide,
            content: slideXml('', 'xmlns:p14="' + EXTENSION_NS.p14 + '"'),
          },
          {
            name: '/ppt/slides/slide2.xml',
            contentType: CONTENT_TYPE.slide,
            content: slideXml('', 'xmlns:p14="' + other + '"'),
          },
        ],
        [[REL_TYPE.officeDocument, '/ppt/presentation.xml']],
      ),
    );

    expect(namespaceOf(census, EXTENSION_NS.p14)?.prefixes).toEqual(['p14']);
    expect(namespaceOf(census, other)?.prefixes).toEqual(['p14']);
    expect(namespaceOf(census, EXTENSION_NS.p14)?.uri).not.toBe(namespaceOf(census, other)?.uri);
  });
});

describe('text is counted the way the layout engine will count it', () => {
  it('counts paragraphs, runs and characters inside a:t only', () => {
    const census = censusPackage(
      oneSlide(
        '<p:sp><p:txBody><a:bodyPr/>' +
          '<a:p><a:r><a:rPr lang="en-GB"/><a:t>Hello</a:t></a:r>' +
          '<a:r><a:t> world</a:t></a:r></a:p>' +
          '<a:p><a:r><a:t>second</a:t></a:r><a:br/></a:p>' +
          '</p:txBody></p:sp>',
      ),
    );

    expect(census.text.paragraphs).toBe(2);
    expect(census.text.runs).toBe(3);
    expect(census.text.characters).toBe('Hello'.length + ' world'.length + 'second'.length);
    expect(census.text.hardBreaks).toBe(1);
    expect(census.languages).toEqual([{ name: 'en-GB', count: 1 }]);
  });

  it('counts a character reference as the character it denotes', () => {
    const census = censusPackage(
      oneSlide(
        '<p:sp><p:txBody><a:p><a:r><a:t>a&amp;b&#8212;c</a:t></a:r></a:p></p:txBody></p:sp>',
      ),
    );
    expect(census.text.characters).toBe(5);
  });

  it('collects every typeface a run can name', () => {
    const census = censusPackage(
      oneSlide(
        '<p:sp><p:txBody><a:p><a:r><a:rPr>' +
          '<a:latin typeface="Calibri"/><a:ea typeface="Meiryo"/>' +
          '<a:cs typeface="Arial"/><a:sym typeface="Wingdings"/>' +
          '</a:rPr><a:t>x</a:t></a:r>' +
          '<a:r><a:rPr><a:latin typeface="+mn-lt"/></a:rPr><a:t>y</a:t></a:r>' +
          '</a:p></p:txBody></p:sp>',
      ),
    );

    expect(census.typefaces.map((entry) => entry.name).sort()).toEqual([
      '+mn-lt',
      'Arial',
      'Calibri',
      'Meiryo',
      'Wingdings',
    ]);
  });
});

describe('a census reports what is wrong without refusing to finish', () => {
  it('turns a part that will not tokenize into a problem and carries on', () => {
    const census = censusPackage(
      pack(
        [
          {
            name: '/ppt/presentation.xml',
            contentType: CONTENT_TYPE.presentation,
            content: presentationXml(),
            rels: [
              [REL_TYPE.slide, '/ppt/slides/slide1.xml'],
              [REL_TYPE.slide, '/ppt/slides/slide2.xml'],
            ],
          },
          {
            name: '/ppt/slides/slide1.xml',
            contentType: CONTENT_TYPE.slide,
            content: DECLARATION + '<p:sld xmlns:p="' + OOXML_NS.p + '"><p:cSld>',
          },
          {
            name: '/ppt/slides/slide2.xml',
            contentType: CONTENT_TYPE.slide,
            content: slideXml('<p:sp/>'),
          },
        ],
        [[REL_TYPE.officeDocument, '/ppt/presentation.xml']],
      ),
    );

    expect(codes(census)).toContain('PART_NOT_WELL_FORMED');
    expect(census.problems.some((problem) => problem.part === '/ppt/slides/slide1.xml')).toBe(true);
    expect(
      census.problems.every((problem) => problem.severity !== 'error' || problem.part !== null),
    ).toBe(true);
    // The undamaged slide was still scanned.
    expect(feature(census, 'shape')).toBe(1);
    expect(census.presentation?.slides).toBe(2);
  });

  it('names a dangling relationship, which PowerPoint tolerates and we preserve', () => {
    // Assembled at the ZIP layer rather than through `PartStore`, which refuses
    // to write an edge *we* broke. That refusal is correct and it is also why
    // this fixture cannot go through it: the deck being described here is one
    // somebody else shipped, and PowerPoint opens it.
    const census = censusPackage(
      rawPack([
        ['[Content_Types].xml', contentTypes()],
        [
          '_rels/.rels',
          relsDocument(relXml('rId1', REL_TYPE.officeDocument, 'ppt/presentation.xml')),
        ],
        ['ppt/presentation.xml', presentationXml()],
        [
          'ppt/_rels/presentation.xml.rels',
          relsDocument(relXml('rId1', REL_TYPE.image, '../media/gone.png')),
        ],
      ]),
    );

    expect(codes(census)).toContain('DANGLING_RELATIONSHIP');
    expect(census.relationships.dangling).toHaveLength(1);
    expect(census.relationships.dangling[0]?.target).toContain('gone.png');
  });

  it('lists parts nothing relates to, without suggesting they are deletable', () => {
    const census = censusPackage(
      pack(
        [
          {
            name: '/ppt/presentation.xml',
            contentType: CONTENT_TYPE.presentation,
            content: presentationXml(),
          },
          {
            name: '/ppt/slideLayouts/slideLayout1.xml',
            contentType: CONTENT_TYPE.slideLayout,
            content: DECLARATION + '<p:sldLayout xmlns:p="' + OOXML_NS.p + '"/>',
          },
        ],
        [[REL_TYPE.officeDocument, '/ppt/presentation.xml']],
      ),
    );

    expect(census.relationships.unreachableParts).toEqual(['/ppt/slideLayouts/slideLayout1.xml']);
    expect(codes(census)).toContain('UNREACHABLE_PARTS');
  });

  it('catches the missing fntdata Default, which is the canonical repair bug', () => {
    const store = PartStore.create();
    store.contentTypes.setDefault('rels', CONTENT_TYPE.relationships);
    store.contentTypes.setDefault('xml', CONTENT_TYPE.xml);
    store.addPart(
      '/ppt/presentation.xml',
      CONTENT_TYPE.presentation,
      encoder.encode(presentationXml()),
    );
    // An Override gives the part a content type, so the package is writable -
    // and PowerPoint still refuses it, because it wants the Default.
    store.addPart('/ppt/fonts/font1.fntdata', 'application/x-fontdata', new Uint8Array([1, 2, 3]));
    store.rootRelationships().addTo(REL_TYPE.officeDocument, '/ppt/presentation.xml');
    store.relationships('/ppt/presentation.xml').addTo(REL_TYPE.font, '/ppt/fonts/font1.fntdata');

    expect(codes(censusPackage(store.write()))).toContain('FNTDATA_DEFAULT_MISSING');
  });

  it('separates having embedded fonts from PowerPoint being willing to read them', () => {
    const fontList =
      '<p:embeddedFontLst><p:embeddedFont>' +
      '<p:font typeface="ProbeAlpha" panose="02000000000000000000" pitchFamily="2" charset="0"/>' +
      '<p:regular r:id="rId9"/>' +
      '</p:embeddedFont></p:embeddedFontLst>';

    const without = censusPackage(oneSlide('', '', fontList));
    expect(without.presentation?.embeddedFonts).toHaveLength(1);
    expect(without.presentation?.embeddedFonts[0]?.typeface).toBe('ProbeAlpha');
    expect(without.presentation?.embeddedFonts[0]?.slots).toEqual(['regular']);
    expect(without.presentation?.embedTrueTypeFonts).toBe(false);
    expect(codes(without)).toContain('EMBED_FLAG_MISSING');

    const withFlag = censusPackage(
      pack(
        [
          {
            name: '/ppt/presentation.xml',
            contentType: CONTENT_TYPE.presentation,
            content: presentationXml(fontList, 'embedTrueTypeFonts="1" saveSubsetFonts="1"'),
          },
        ],
        [[REL_TYPE.officeDocument, '/ppt/presentation.xml']],
      ),
    );
    expect(withFlag.presentation?.embedTrueTypeFonts).toBe(true);
    expect(withFlag.presentation?.saveSubsetFonts).toBe(true);
    expect(codes(withFlag)).not.toContain('EMBED_FLAG_MISSING');
  });

  it('reports a slide list that disagrees with the slide parts', () => {
    const census = censusPackage(
      pack(
        [
          {
            name: '/ppt/presentation.xml',
            contentType: CONTENT_TYPE.presentation,
            content: presentationXml(
              '<p:sldIdLst><p:sldId id="256" r:id="rId1"/><p:sldId id="257" r:id="rId2"/></p:sldIdLst>',
            ),
            rels: [[REL_TYPE.slide, '/ppt/slides/slide1.xml']],
          },
          {
            name: '/ppt/slides/slide1.xml',
            contentType: CONTENT_TYPE.slide,
            content: slideXml(''),
          },
        ],
        [[REL_TYPE.officeDocument, '/ppt/presentation.xml']],
      ),
    );

    expect(codes(census)).toContain('SLIDE_COUNT_MISMATCH');
  });
});

describe('the result is a message, not a document', () => {
  it('survives JSON and structured clone unchanged, which is what both consumers do', () => {
    const census = censusPackage(
      oneSlide('<p:sp><p:txBody><a:p><a:r><a:t>x</a:t></a:r></a:p></p:txBody></p:sp>'),
    );

    // `postMessage` uses structured clone; `inspect --json` uses JSON. A census
    // that means different things through the two would be a census whose
    // definition depends on who asked.
    expect(structuredClone(census)).toEqual(census);
    expect(JSON.parse(JSON.stringify(census))).toEqual(census);
  });

  it('holds no XNode: nothing in it has a parent, children or a source', () => {
    // The tree is what must not escape. A shallow property sweep is enough to
    // catch it, because an XNode is recognisable by any of three field names.
    const census = censusPackage(oneSlide('<p:sp/>'));
    const seen = new Set<unknown>();
    const banned = ['parent', 'children', 'qname', 'source'];

    const walk = (value: unknown, path: string): void => {
      if (value === null || typeof value !== 'object' || seen.has(value)) return;
      seen.add(value);
      for (const [key, child] of Object.entries(value)) {
        expect(banned, path + '.' + key).not.toContain(key);
        walk(child, path + '.' + key);
      }
    };
    walk(census, 'census');
    expect(seen.size).toBeGreaterThan(10);
  });

  it('formats to stable text that says what it found', () => {
    const census = censusPackage(
      oneSlide('<p:sp><p:spPr><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr></p:sp>'),
    );
    const text = formatCensus(census, { namespaces: true, parts: true });

    expect(text).toContain('PACKAGE');
    expect(text).toContain('PRESENTATION');
    expect(text).toContain('slide size     13.33 x 7.50 in (screen16x9)');
    expect(text).toContain('Preset geometry');
    expect(text).toContain('/ppt/slides/slide1.xml');
    expect(text).toContain(OOXML_NS.a);
    // Two runs of the same census must produce the same bytes, or it cannot be
    // committed or diffed.
    expect(formatCensus(census, { namespaces: true, parts: true })).toBe(text);
  });
});

describe('timings are measured, not estimated', () => {
  it('reports scan throughput against an injected clock', () => {
    let clock = 0;
    const census = censusPackage(oneSlide('<p:sp/>'), {
      now: () => (clock += 10),
    });

    expect(census.timings.totalMs).toBeGreaterThan(0);
    expect(census.timings.partsScanned).toBe(2);
    expect(census.timings.xmlBytesScanned).toBeGreaterThan(0);
  });

  it('calls back with progress, because a 200 MB deck needs a bar', () => {
    const seen: number[] = [];
    const census = censusPackage(oneSlide('<p:sp/>'), {
      onProgress: (done, total) => {
        seen.push(done);
        expect(total).toBe(2);
      },
    });

    expect(seen).toEqual([0, 1, 2]);
    expect(census.archive.parts).toBe(2);
  });
});
