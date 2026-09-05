import { describe, expect, it } from 'vitest';
import { REL_TYPE } from './constants.js';
import { isOpcError } from './errors.js';
import {
  isValidRelationshipId,
  relativeTargetFor,
  Relationships,
  sourcePartNameForRels,
} from './relationships.js';
import { relsXml, relXml } from './testing/build-package.js';

const encoder = new TextEncoder();
const decoder = new TextDecoder();

function parse(xml: string, source: string): Relationships {
  return Relationships.parse(encoder.encode(xml), source);
}

function codeOf(fn: () => unknown): string {
  try {
    fn();
  } catch (error) {
    if (isOpcError(error)) return error.code;
    return 'NOT_AN_OPC_ERROR: ' + (error instanceof Error ? error.constructor.name : typeof error);
  }
  return 'DID_NOT_THROW';
}

describe('reading a relationship part', () => {
  it('reads the root relationships PowerPoint writes, out of numeric order and all', () => {
    // Transcribed from a file PowerPoint 365 saved: rId3, rId2, rId1, in that
    // document order. Nothing about relationship order is meaningful.
    const rels = parse(
      relsXml(
        relXml('rId3', REL_TYPE.extendedProperties, 'docProps/app.xml'),
        relXml('rId2', REL_TYPE.coreProperties, 'docProps/core.xml'),
        relXml('rId1', REL_TYPE.officeDocument, 'ppt/presentation.xml'),
      ),
      '/',
    );
    expect(rels.all.map((r) => r.id)).toEqual(['rId3', 'rId2', 'rId1']);
    expect(rels.firstOfType(REL_TYPE.officeDocument)?.id).toBe('rId1');
    expect(rels.partName).toBe('/_rels/.rels');
  });

  it('resolves a root target against the package root, not against _rels', () => {
    const rels = parse(
      relsXml(relXml('rId1', REL_TYPE.officeDocument, 'ppt/presentation.xml')),
      '/',
    );
    expect(rels.resolve(rels.all[0]!)).toBe('/ppt/presentation.xml');
  });

  it('resolves a slide target against the slide’s folder, not the .rels folder', () => {
    const rels = parse(
      relsXml(relXml('rId1', REL_TYPE.slideLayout, '../slideLayouts/slideLayout6.xml')),
      '/ppt/slides/slide1.xml',
    );
    expect(rels.partName).toBe('/ppt/slides/_rels/slide1.xml.rels');
    expect(rels.resolve(rels.all[0]!)).toBe('/ppt/slideLayouts/slideLayout6.xml');
  });

  it('reads an absolute target, which three quarters of one real corpus uses', () => {
    const rels = parse(
      relsXml(relXml('rId1', REL_TYPE.slideLayout, '/ppt/slideLayouts/slideLayout1.xml')),
      '/ppt/slides/slide1.xml',
    );
    expect(rels.resolve(rels.all[0]!)).toBe('/ppt/slideLayouts/slideLayout1.xml');
  });

  it('defaults TargetMode to Internal and keeps an external target verbatim', () => {
    const rels = parse(
      relsXml(
        relXml('rId1', REL_TYPE.slide, 'slides/slide1.xml'),
        // A hyperlink target carries `&amp;`, never a bare ampersand - a bare
        // one is not well-formed XML and the reader refuses it, as it should.
        relXml('rId2', REL_TYPE.hyperlink, 'https://example.com/a?x=1&amp;y=2', 'External'),
      ),
      '/ppt/presentation.xml',
    );
    expect(rels.byId('rId1')?.targetMode).toBe('Internal');
    expect(rels.byId('rId2')?.targetMode).toBe('External');
    expect(rels.byId('rId2')?.target).toBe('https://example.com/a?x=1&y=2');
    expect(codeOf(() => rels.resolve(rels.byId('rId2')!))).toBe('ERR_INVALID_PART_NAME');
    expect(rels.targetOf('rId2')).toBeUndefined();
  });

  it('collects relationships by type', () => {
    const rels = parse(
      relsXml(
        relXml('rId1', REL_TYPE.slide, 'slides/slide1.xml'),
        relXml('rId2', REL_TYPE.slide, 'slides/slide2.xml'),
        relXml('rId3', REL_TYPE.theme, 'theme/theme1.xml'),
      ),
      '/ppt/presentation.xml',
    );
    expect(rels.byType(REL_TYPE.slide).map((r) => r.id)).toEqual(['rId1', 'rId2']);
    expect(rels.size).toBe(3);
  });

  it('gives a part with no .rels an empty collection rather than nothing', () => {
    const rels = Relationships.empty('/ppt/slides/slide1.xml');
    expect(rels.size).toBe(0);
    expect(rels.dirty).toBe(false);
  });
});

describe('the failures PowerPoint refuses a package over', () => {
  it('refuses an Id that is not an NCName', () => {
    // Verified against the installed PowerPoint: Id="1rId" makes it report the
    // whole file as corrupted and unreadable.
    for (const bad of ['1rId', 'r Id', 'a:b', '', '-x', '.x']) {
      expect(
        codeOf(() => parse(relsXml(relXml(bad, REL_TYPE.slide, 'a.xml')), '/p.xml')),
        bad,
      ).toBe('ERR_INVALID_RELATIONSHIP_ID');
    }
  });

  it('accepts the NCNames that are legal, including ones nobody writes', () => {
    for (const good of ['rId1', '_x', 'a.b-c', 'docRel', 'xé']) {
      expect(isValidRelationshipId(good), good).toBe(true);
    }
    for (const bad of ['1a', 'a b', 'a:b', '']) {
      expect(isValidRelationshipId(bad), bad).toBe(false);
    }
  });

  it('refuses a duplicate Id in one part', () => {
    expect(
      codeOf(() =>
        parse(
          relsXml(relXml('rId1', REL_TYPE.slide, 'a.xml'), relXml('rId1', REL_TYPE.slide, 'b.xml')),
          '/p.xml',
        ),
      ),
    ).toBe('ERR_DUPLICATE_RELATIONSHIP_ID');
  });

  it('lets the same Id exist in two different parts, because its scope is one file', () => {
    const a = parse(relsXml(relXml('rId1', REL_TYPE.slide, 'a.xml')), '/ppt/presentation.xml');
    const b = parse(relsXml(relXml('rId1', REL_TYPE.theme, 'b.xml')), '/ppt/slides/slide1.xml');
    expect(a.byId('rId1')?.type).toBe(REL_TYPE.slide);
    expect(b.byId('rId1')?.type).toBe(REL_TYPE.theme);
  });

  it('refuses a target that climbs out of the package', () => {
    const rels = parse(
      relsXml(relXml('rId1', REL_TYPE.image, '../../../../etc/passwd')),
      '/ppt/slides/slide1.xml',
    );
    expect(codeOf(() => rels.resolve(rels.all[0]!))).toBe('ERR_TARGET_ESCAPES_PACKAGE');
  });

  it('refuses missing attributes, an empty internal target and a bogus TargetMode', () => {
    expect(codeOf(() => parse(relsXml('<Relationship Id="rId1" Type="t"/>'), '/p.xml'))).toBe(
      'ERR_MALFORMED_XML',
    );
    expect(codeOf(() => parse(relsXml(relXml('rId1', 't', '')), '/p.xml'))).toBe(
      'ERR_MALFORMED_XML',
    );
    expect(
      codeOf(() =>
        parse(
          relsXml('<Relationship Id="rId1" Type="t" Target="a" TargetMode="Sideways"/>'),
          '/p.xml',
        ),
      ),
    ).toBe('ERR_MALFORMED_XML');
  });

  it('refuses a root element that is not Relationships', () => {
    expect(codeOf(() => parse('<Types/>', '/p.xml'))).toBe('ERR_MALFORMED_XML');
  });
});

describe('allocating an id', () => {
  it('takes one past the highest number, not one past the count', () => {
    // 99 of 1296 real relationship parts have gaps in their numbering. Counting
    // would have collided in every one of them.
    const rels = parse(
      relsXml(
        relXml('rId1', REL_TYPE.slide, 'a.xml'),
        relXml('rId7', REL_TYPE.slide, 'b.xml'),
        relXml('rId3', REL_TYPE.slide, 'c.xml'),
      ),
      '/p.xml',
    );
    expect(rels.nextId()).toBe('rId8');
    expect(rels.add(REL_TYPE.slide, 'd.xml').id).toBe('rId8');
    expect(rels.add(REL_TYPE.slide, 'e.xml').id).toBe('rId9');
  });

  it('starts at rId1 in an empty collection', () => {
    expect(Relationships.empty('/p.xml').nextId()).toBe('rId1');
  });

  it('steps over ids that are not rIdN but would collide', () => {
    const rels = parse(
      relsXml(relXml('docRel', REL_TYPE.slide, 'a.xml'), relXml('rId2', REL_TYPE.slide, 'b.xml')),
      '/p.xml',
    );
    expect(rels.nextId()).toBe('rId3');
  });

  it('removes a relationship and keeps the rest addressable', () => {
    const rels = parse(
      relsXml(
        relXml('rId1', REL_TYPE.slide, 'a.xml'),
        relXml('rId2', REL_TYPE.slide, 'b.xml'),
        relXml('rId3', REL_TYPE.slide, 'c.xml'),
      ),
      '/p.xml',
    );
    expect(rels.remove('rId1')).toBe(true);
    expect(rels.byId('rId2')?.target).toBe('b.xml');
    expect(rels.byId('rId3')?.target).toBe('c.xml');
    expect(rels.remove('rId1')).toBe(false);
    expect(rels.dirty).toBe(true);
  });

  it('addTo computes the relative target Office would have written', () => {
    const rels = Relationships.empty('/ppt/slides/slide1.xml');
    const rel = rels.addTo(REL_TYPE.slideLayout, '/ppt/slideLayouts/slideLayout2.xml');
    expect(rel.target).toBe('../slideLayouts/slideLayout2.xml');
    expect(rels.resolve(rel)).toBe('/ppt/slideLayouts/slideLayout2.xml');
  });
});

describe('naming', () => {
  it('maps a .rels part back to the part it describes', () => {
    expect(sourcePartNameForRels('/_rels/.rels')).toBe('/');
    expect(sourcePartNameForRels('/ppt/_rels/presentation.xml.rels')).toBe('/ppt/presentation.xml');
    expect(sourcePartNameForRels('/ppt/slides/_rels/slide1.xml.rels')).toBe(
      '/ppt/slides/slide1.xml',
    );
    expect(codeOf(() => sourcePartNameForRels('/ppt/slides/slide1.xml'))).toBe(
      'ERR_INVALID_PART_NAME',
    );
  });

  it('computes relative targets in both directions', () => {
    expect(relativeTargetFor('/', '/ppt/presentation.xml')).toBe('ppt/presentation.xml');
    expect(relativeTargetFor('/ppt/presentation.xml', '/ppt/slides/slide1.xml')).toBe(
      'slides/slide1.xml',
    );
    expect(relativeTargetFor('/ppt/slides/slide1.xml', '/ppt/slides/slide2.xml')).toBe(
      'slide2.xml',
    );
    expect(relativeTargetFor('/ppt/slides/slide1.xml', '/ppt/media/image1.png')).toBe(
      '../media/image1.png',
    );
    expect(relativeTargetFor('/ppt/presentation.xml', '/docProps/core.xml')).toBe(
      '../docProps/core.xml',
    );
  });
});

describe('writing a relationship part back', () => {
  it('reproduces Office’s layout, with TargetMode written only when External', () => {
    const xml = relsXml(
      relXml('rId1', REL_TYPE.officeDocument, 'ppt/presentation.xml'),
      relXml('rId2', REL_TYPE.hyperlink, 'https://example.com/', 'External'),
    );
    expect(decoder.decode(parse(xml, '/').serialize())).toBe(xml);
  });

  it('drops a redundant TargetMode="Internal", as Office does', () => {
    const rels = parse(
      relsXml(relXml('rId1', REL_TYPE.slide, 'a.xml', 'Internal')),
      '/ppt/presentation.xml',
    );
    expect(decoder.decode(rels.serialize())).not.toContain('TargetMode');
  });

  it('escapes an external target so an ampersand survives the trip', () => {
    const rels = Relationships.empty('/ppt/slides/slide1.xml');
    rels.add(REL_TYPE.hyperlink, 'https://example.com/?a=1&b=2', 'External');
    const back = Relationships.parse(rels.serialize(), '/ppt/slides/slide1.xml');
    expect(back.byId('rId1')?.target).toBe('https://example.com/?a=1&b=2');
    expect(decoder.decode(rels.serialize())).toContain('&amp;');
  });
});
