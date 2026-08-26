import { describe, expect, it } from 'vitest';
import { isOpcError, OpcError } from './errors.js';
import {
  checkPartNameCollisions,
  isContentTypesStreamName,
  isRelationshipPartName,
  isValidPartName,
  normalizePartName,
  partDirectory,
  partExtension,
  partNameFromZipEntry,
  relsPartNameFor,
  resolveRelativeTarget,
  toPartName,
  validatePartName,
  zipEntryNameFor,
  type PartNameRule,
} from './pack-uri.js';

function rules(name: string): PartNameRule[] {
  return validatePartName(name).map((v) => v.rule);
}

function codeOf(fn: () => unknown): string {
  try {
    fn();
  } catch (error) {
    return isOpcError(error) ? error.code : 'NOT_AN_OPC_ERROR:' + String(error);
  }
  return 'DID_NOT_THROW';
}

describe('part names that a real deck contains', () => {
  it.each([
    '/ppt/presentation.xml',
    '/ppt/slides/slide1.xml',
    '/ppt/slides/_rels/slide1.xml.rels',
    '/ppt/slideLayouts/slideLayout11.xml',
    '/ppt/media/image1.png',
    '/ppt/fonts/font1.fntdata',
    '/ppt/embeddings/oleObject1.bin',
    '/docProps/app.xml',
    '/_rels/.rels',
    '/customXml/item1.xml',
    '/ppt/media/image%201.png',
  ])('accepts %s', (name) => {
    expect(validatePartName(name)).toEqual([]);
    expect(isValidPartName(name)).toBe(true);
  });
});

describe('the OPC grammar, one rule at a time', () => {
  it('M1.1 - a part name shall not be empty', () => {
    expect(rules('')).toEqual(['M1.1']);
    expect(rules('/')).toEqual(['M1.1']);
  });

  it('M1.3 - a part name shall not have empty segments', () => {
    expect(rules('/ppt//slide1.xml')).toContain('M1.3');
    // A ZIP entry name that already starts with a slash produces `//`.
    expect(rules('//ppt/slide1.xml')).toContain('M1.3');
  });

  it('M1.4 - a part name shall start with a forward slash', () => {
    expect(rules('ppt/slides/slide1.xml')).toEqual(['M1.4']);
  });

  it('M1.5 - a part name shall not end with a forward slash', () => {
    // This is also how ZIP directory entries present themselves.
    expect(rules('/ppt/media/')).toEqual(['M1.5']);
  });

  it('M1.6 - a segment holds pchar only', () => {
    const spaced = validatePartName('/ppt/media/my image.png');
    expect(spaced.map((v) => v.rule)).toContain('M1.6');
    // A literal space is out of spec but harmless, and failing a whole deck
    // over one media file name would be worse than tolerating it.
    expect(spaced.every((v) => v.severity === 'warning')).toBe(true);
    expect(rules('/ppt/slides/slide[1].xml')).toContain('M1.6');
    expect(rules('/ppt/media/image%zz.png')).toContain('M1.6');
  });

  it('M1.7 - no percent-encoded separators', () => {
    expect(rules('/ppt%2fslides%2fslide1.xml')).toContain('M1.7');
    expect(rules('/ppt/..%2f..%2fetc/passwd')).toContain('M1.7');
    expect(rules('/ppt%5cslides')).toContain('M1.7');
  });

  it('M1.8 - unreserved characters appear literally', () => {
    const encoded = validatePartName('/ppt/slides/slide%31.xml');
    expect(encoded.map((v) => v.rule)).toEqual(['M1.8']);
    expect(encoded[0]?.severity).toBe('warning');
  });

  it('M1.9 - a segment shall not end with a dot', () => {
    expect(rules('/ppt/slides.')).toContain('M1.9');
    expect(rules('/ppt/slide1.xml.')).toContain('M1.9');
  });

  it('M1.10 and X1.4 - "." and ".." are not part-name segments at all', () => {
    // The two rules together are what makes zip-slip unrepresentable in a
    // conformant part name. X1.4 is reported instead of them because
    // "path traversal" is the useful thing to read in a bug report.
    expect(rules('/ppt/../../../etc/passwd')).toContain('X1.4');
    expect(rules('/ppt/./slide1.xml')).toContain('X1.4');
    expect(rules('/..')).toEqual(['X1.4']);
    expect(rules('/ppt/.../x')).toContain('X1.4');
  });

  it('M1.11 - one name cannot be another name plus segments', () => {
    const found = checkPartNameCollisions([
      '/ppt/media',
      '/ppt/media/image1.png',
      '/ppt/presentation.xml',
    ]);
    expect(found.map((f) => f.rule)).toEqual(['M1.11']);
    expect(found[0]?.names).toEqual(['/ppt/media', '/ppt/media/image1.png']);
  });

  it('M1.12 - equivalence is case-insensitive', () => {
    const found = checkPartNameCollisions(['/ppt/slides/slide1.xml', '/PPT/Slides/Slide1.XML']);
    expect(found.map((f) => f.rule)).toEqual(['M1.12']);
    expect(normalizePartName('/PPT/Slide1.XML')).toBe('/ppt/slide1.xml');
  });

  it('M1.12 folds ASCII only', () => {
    // `toLowerCase()` would fold characters that are not in the grammar,
    // making two names collide that a conformant comparison keeps apart.
    expect(normalizePartName('/ppt/\u0130.xml')).toBe('/ppt/\u0130.xml');
  });

  it('a clean package produces no collisions', () => {
    expect(
      checkPartNameCollisions([
        '/ppt/presentation.xml',
        '/ppt/slides/slide1.xml',
        '/ppt/slides/_rels/slide1.xml.rels',
      ]),
    ).toEqual([]);
  });
});

describe('hostility the spec never contemplated', () => {
  it('X1.1 - control characters and NUL', () => {
    expect(rules('/ppt/slide1.xml\u0000.png')).toContain('X1.1');
    expect(rules('/ppt/slide%001.xml')).toContain('X1.1');
    expect(rules('/ppt/sl\u001bide.xml')).toContain('X1.1');
  });

  it('X1.2 - a backslash is a separator to some readers and not to others', () => {
    expect(rules('/ppt\\slides\\slide1.xml')).toContain('X1.2');
    expect(rules('/..\\..\\evil.xml')).toContain('X1.2');
  });

  it('X1.3 - a drive letter is legal pchar and still hostile', () => {
    // A colon is a pchar, so `/C:/Windows/System32/evil.dll` satisfies the
    // OPC grammar completely. It is rejected here rather than in the spec.
    expect(rules('/C:/Windows/System32/evil.dll')).toContain('X1.3');
    expect(rules('/ppt/C:/x.xml')).not.toContain('X1.3');
  });

  it('X1.5 - names have a ceiling', () => {
    expect(rules('/' + 'a'.repeat(2000))).toEqual(['X1.5']);
  });
});

describe('toPartName', () => {
  it('throws OpcError with the offending rules named', () => {
    expect(codeOf(() => toPartName('/ppt/../etc/passwd'))).toBe('ERR_INVALID_PART_NAME');
    try {
      toPartName('/ppt/../etc/passwd');
    } catch (error) {
      expect(error).toBeInstanceOf(OpcError);
      expect((error as OpcError).message).toContain('[X1.4]');
      expect((error as OpcError).detail.entry).toBe('/ppt/../etc/passwd');
    }
  });

  it('tolerates warnings', () => {
    expect(toPartName('/ppt/media/my image.png')).toBe('/ppt/media/my image.png');
  });
});

describe('ZIP entry names and part names are different namespaces', () => {
  it('adds and removes the leading slash', () => {
    const name = partNameFromZipEntry('ppt/slides/slide1.xml');
    expect(name).toBe('/ppt/slides/slide1.xml');
    expect(zipEntryNameFor(name)).toBe('ppt/slides/slide1.xml');
  });

  it('rejects a traversing entry name', () => {
    expect(codeOf(() => partNameFromZipEntry('../../../etc/passwd'))).toBe('ERR_INVALID_PART_NAME');
    expect(codeOf(() => partNameFromZipEntry('ppt/../../evil.xml'))).toBe('ERR_INVALID_PART_NAME');
    expect(codeOf(() => partNameFromZipEntry('/etc/passwd'))).toBe('ERR_INVALID_PART_NAME');
  });

  it('the content-type stream is a ZIP entry name and not a part name', () => {
    // `[` and `]` are not pchar, but M1.6 is only a warning - it has to be, or
    // a media file with a space in its name would fail a whole deck. So the
    // grammar alone would let this through as a part name, and the ZIP-to-OPC
    // boundary is where the refusal belongs.
    expect(validatePartName('/[Content_Types].xml').map((v) => v.rule)).toEqual(['M1.6', 'M1.6']);
    expect(validatePartName('/[Content_Types].xml').every((v) => v.severity === 'warning')).toBe(
      true,
    );
    expect(codeOf(() => partNameFromZipEntry('[Content_Types].xml'))).toBe('ERR_INVALID_PART_NAME');
    expect(codeOf(() => partNameFromZipEntry('[content_types].XML'))).toBe('ERR_INVALID_PART_NAME');
    expect(isContentTypesStreamName('[Content_Types].xml')).toBe(true);
    expect(isContentTypesStreamName('[content_types].xml')).toBe(true);
    expect(isContentTypesStreamName('ppt/[Content_Types].xml')).toBe(false);
  });
});

describe('naming helpers', () => {
  it('splits directory, extension and rels name', () => {
    expect(partDirectory('/ppt/slides/slide1.xml')).toBe('/ppt/slides/');
    expect(partExtension('/ppt/media/image1.PNG')).toBe('PNG');
    expect(partExtension('/ppt/fonts/font1.fntdata')).toBe('fntdata');
    expect(partExtension('/_rels/.rels')).toBe('');
    expect(partExtension('/ppt/noextension')).toBe('');
  });

  it('derives the relationship part for a part', () => {
    expect(relsPartNameFor('/ppt/slides/slide1.xml')).toBe('/ppt/slides/_rels/slide1.xml.rels');
    expect(relsPartNameFor('/ppt/presentation.xml')).toBe('/ppt/_rels/presentation.xml.rels');
    expect(relsPartNameFor('/')).toBe('/_rels/.rels');
  });

  it('recognises relationship parts', () => {
    expect(isRelationshipPartName('/_rels/.rels')).toBe(true);
    expect(isRelationshipPartName('/ppt/slides/_rels/slide1.xml.rels')).toBe(true);
    expect(isRelationshipPartName('/ppt/slides/slide1.xml')).toBe(false);
    expect(isRelationshipPartName('/ppt/_rels_/x.rels')).toBe(false);
  });
});

describe('resolveRelativeTarget', () => {
  it('resolves against the source part folder, not the package root', () => {
    // This is the one place `..` is legitimate: a slide binds its layout only
    // through this relationship, so getting it wrong loses the whole
    // inheritance chain.
    expect(
      resolveRelativeTarget('/ppt/slides/slide1.xml', '../slideLayouts/slideLayout1.xml'),
    ).toBe('/ppt/slideLayouts/slideLayout1.xml');
    expect(resolveRelativeTarget('/ppt/slides/slide1.xml', 'notesSlide1.xml')).toBe(
      '/ppt/slides/notesSlide1.xml',
    );
    expect(resolveRelativeTarget('/ppt/presentation.xml', 'slides/slide1.xml')).toBe(
      '/ppt/slides/slide1.xml',
    );
  });

  it('handles absolute targets, "." segments and fragments', () => {
    expect(resolveRelativeTarget('/ppt/slides/slide1.xml', '/ppt/media/image1.png')).toBe(
      '/ppt/media/image1.png',
    );
    expect(resolveRelativeTarget('/ppt/slides/slide1.xml', './slide2.xml')).toBe(
      '/ppt/slides/slide2.xml',
    );
    expect(resolveRelativeTarget('/ppt/presentation.xml', 'slides/slide1.xml#anchor')).toBe(
      '/ppt/slides/slide1.xml',
    );
  });

  it('clamps at the package root', () => {
    expect(
      codeOf(() => resolveRelativeTarget('/ppt/slides/slide1.xml', '../../../../etc/passwd')),
    ).toBe('ERR_TARGET_ESCAPES_PACKAGE');
    expect(codeOf(() => resolveRelativeTarget('/ppt/presentation.xml', '../../x'))).toBe(
      'ERR_TARGET_ESCAPES_PACKAGE',
    );
    // Exactly at the root is fine; one more is not.
    expect(resolveRelativeTarget('/ppt/presentation.xml', '../docProps/app.xml')).toBe(
      '/docProps/app.xml',
    );
  });

  it('rejects an empty target', () => {
    expect(codeOf(() => resolveRelativeTarget('/ppt/presentation.xml', ''))).toBe(
      'ERR_INVALID_PART_NAME',
    );
    expect(codeOf(() => resolveRelativeTarget('/ppt/presentation.xml', '#only-a-fragment'))).toBe(
      'ERR_INVALID_PART_NAME',
    );
  });
});
