import { PartStore } from '@pptx-studio/opc';
import { parseXmlString } from '@pptx-studio/xml';
import { describe, expect, it } from 'vitest';
import { isValidateError } from '../errors.js';
import { attributeLocation, elementLocation, lineColumn, xpathOf } from './location.js';
import { formatReport, isReport } from './report.js';
import { deck, deckBytes, minimalDeck } from '../testing/deck.js';
import { assertValid, validatePackage } from '../validate.js';

const encoder = new TextEncoder();

/** A deck carrying one fatal defect from the moment it was opened. */
function defective(): Uint8Array {
  const parts = minimalDeck();
  return deckBytes({
    parts: {
      // A dangling relationship: PowerPoint tolerates one, and files that have
      // been through three other tools routinely have one.
      'ppt/slides/_rels/slide1.xml.rels': parts['ppt/slides/_rels/slide1.xml.rels']!.replace(
        '</Relationships>',
        '<Relationship Id="rId2" ' +
          'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" ' +
          'Target="../media/image1.png"/></Relationships>',
      ),
    },
  });
}

describe('origin', () => {
  it('marks a defect that was already there as inherited, and lets the export through', () => {
    const bytes = defective();
    const store = PartStore.open(bytes);
    const baseline = PartStore.open(bytes);

    const report = validatePackage({ bytes, store, baseline, baselineBytes: bytes });
    const dangling = report.findings.filter((finding) => finding.rule === 'V008');

    expect(dangling).toHaveLength(1);
    expect(dangling[0]!.severity).toBe('fatal');
    expect(dangling[0]!.origin).toBe('inherited');
    // The point of the whole mechanism: a user can still save their own file.
    expect(report.blocking).toBe(0);
    expect(report.ok).toBe(true);
  });

  it('marks the same defect as introduced when we are the ones who made it', () => {
    const clean = deckBytes();
    const store = PartStore.open(defective());
    const baseline = PartStore.open(clean);

    const report = validatePackage({ store, baseline });
    const dangling = report.findings.filter((finding) => finding.rule === 'V008');

    expect(dangling).toHaveLength(1);
    expect(dangling[0]!.origin).toBe('introduced');
    expect(report.ok).toBe(false);
  });

  it('separates the two in one report', () => {
    // The case the mechanism exists for: a file that arrived with a defect, and
    // an edit that added a different one. Only the second may refuse the export.
    //
    // The edit goes through `replacePart`, which is what an edit is. Building
    // the "after" package as a second archive instead would be a different
    // scenario with a different right answer - `V027` would fire, correctly,
    // because a part would have changed with nothing claiming to have edited it.
    const parts = minimalDeck();
    const before = defective();
    const after = PartStore.open(before);
    after.replacePart(
      '/ppt/presentation.xml',
      encoder.encode(
        parts['ppt/presentation.xml']!.replace('<p:sldId id="256"', '<p:sldId id="7"'),
      ),
    );

    const report = validatePackage({ store: after, baseline: PartStore.open(before) });
    const byRule = new Map(report.findings.map((finding) => [finding.rule, finding.origin]));

    expect(byRule.get('V008')).toBe('inherited');
    expect(byRule.get('V018')).toBe('introduced');
    expect(report.blocking).toBe(1);
  });

  it('says `unknown` when there is no baseline to ask about', () => {
    const bytes = defective();
    const report = validatePackage({ bytes });
    expect(report.findings.every((finding) => finding.origin === 'unknown')).toBe(true);
    // With nothing to compare against, a fatal is treated as ours. Refusing is
    // the safe direction when the question cannot be answered.
    expect(report.ok).toBe(false);
  });
});

describe('assertValid', () => {
  it('returns the report when nothing blocking was found', () => {
    const { bytes, store } = deck();
    expect(assertValid({ bytes, store }).ok).toBe(true);
  });

  it('throws with the first finding named, and carries the whole report', () => {
    const parts = minimalDeck();
    const { bytes, store } = deck({
      parts: {
        'ppt/presentation.xml': parts['ppt/presentation.xml']!.replace(
          '<p:notesSz cx="6858000" cy="9144000"/>',
          '',
        ),
      },
    });

    try {
      assertValid({ bytes, store });
      expect.unreachable('assertValid should have refused');
    } catch (error) {
      expect(isValidateError(error)).toBe(true);
      if (!isValidateError(error)) return;
      expect(error.code).toBe('ERR_VALIDATION_FAILED');
      expect(error.message).toContain('V013');
      expect(error.message).toContain('/ppt/presentation.xml');
      // The report rides along, so a caller does not have to validate twice to
      // find out what else was wrong.
      expect(isReport(error.detail.report)).toBe(true);
    }
  });

  it('does not throw over a defect the file arrived with', () => {
    const bytes = defective();
    const store = PartStore.open(bytes);
    expect(() => assertValid({ store, baseline: PartStore.open(bytes) })).not.toThrow();
  });

  it('refuses to guess when it is given neither bytes nor a store', () => {
    expect(() => validatePackage({})).toThrow(/neither/);
  });

  it('says so when the bytes are not a package at all', () => {
    const notAZip = encoder.encode('this is not a zip file, not even slightly');
    try {
      validatePackage({ bytes: notAZip });
      expect.unreachable('should not have validated');
    } catch (error) {
      expect(isValidateError(error) && error.code).toBe('ERR_UNVALIDATABLE');
    }
  });
});

describe('formatReport', () => {
  it('groups by part, marks inherited findings, and counts what blocks', () => {
    const bytes = defective();
    const text = formatReport(
      validatePackage({ bytes, store: PartStore.open(bytes), baseline: PartStore.open(bytes) }),
    );

    expect(text).toContain('/ppt/slides/_rels/slide1.xml.rels');
    expect(text).toContain('V008  fatal (already there)');
    expect(text).toContain('/Relationships/Relationship[2]');
    expect(text).toContain('0 blocking');
  });

  it('can hide what a caller does not want, and says nothing is shown rather than nothing exists', () => {
    const bytes = defective();
    const report = validatePackage({
      bytes,
      store: PartStore.open(bytes),
      baseline: PartStore.open(bytes),
    });

    expect(formatReport(report, { inherited: false })).toContain('no findings shown');
  });

  it('explains a rule once, on request', () => {
    const bytes = defective();
    const text = formatReport(validatePackage({ bytes }), { explain: true });
    expect(text).toContain('V008 - internal targets resolve');
    expect(text).toContain('classic reason a rebuilt deck loses all its images');
  });
});

describe('locations', () => {
  const document = parseXmlString(
    '<p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">' +
      '<p:cSld><p:spTree><p:sp id="1"/><p:sp id="2"/><p:sp id="3"/></p:spTree></p:cSld>' +
      '</p:sld>',
  );
  const tree = document.root.children[0]!;
  const spTree = (tree as { children: unknown[] }).children[0] as {
    children: { qname: string }[];
  };

  it('numbers only what needs numbering', () => {
    // `p:cSld` and `p:spTree` are alone among their siblings, so no predicate.
    // The three `p:sp` are not, so all three get one.
    const paths = spTree.children.map((child) => xpathOf(child as never));
    expect(paths).toEqual([
      '/p:sld/p:cSld/p:spTree/p:sp[1]',
      '/p:sld/p:cSld/p:spTree/p:sp[2]',
      '/p:sld/p:cSld/p:spTree/p:sp[3]',
    ]);
  });

  it('uses the prefixes the document used, not ones we prefer', () => {
    // A part is free to bind PresentationML to any prefix, and `mc:Ignorable`
    // and `mc:Choice/@Requires` hold prefixes rather than URIs - so a prefix is
    // load-bearing data here and a path that canonicalised it would point at
    // something the reader cannot find in their file.
    const odd = parseXmlString(
      '<pp:sld xmlns:pp="http://schemas.openxmlformats.org/presentationml/2006/main">' +
        '<pp:cSld/></pp:sld>',
    );
    expect(xpathOf(odd.root.children[0] as never)).toBe('/pp:sld/pp:cSld');
  });

  it('points at the attribute, and at the element when the attribute is not there', () => {
    const sp = spTree.children[1] as never;
    expect(attributeLocation('/ppt/slides/slide1.xml', sp, 'id').xpath).toBe(
      '/p:sld/p:cSld/p:spTree/p:sp[2]/@id',
    );
    const missing = attributeLocation('/ppt/slides/slide1.xml', sp, 'name');
    expect(missing.xpath).toBe('/p:sld/p:cSld/p:spTree/p:sp[2]/@name');
    expect(missing.offset).toBe(elementLocation('/ppt/slides/slide1.xml', sp).offset);
  });

  it('turns an offset into a line and column when somebody asks', () => {
    expect(lineColumn('a\nbc\ndef', 0)).toEqual({ line: 1, column: 1 });
    expect(lineColumn('a\nbc\ndef', 5)).toEqual({ line: 3, column: 1 });
    expect(lineColumn('a\nbc\ndef', 999)).toEqual({ line: 3, column: 4 });
  });
});
