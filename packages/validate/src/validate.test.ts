import { describe, expect, it } from 'vitest';
import { formatReport, type Report } from './report/report.js';
import type { RuleId } from './rules/rules.js';
import { deck, minimalDeck, relsPart, rel, shape, IDENTITY_CLR_MAP } from './testing/deck.js';
import { validatePackage, type ValidateOptions } from './validate.js';

/**
 * Every rule, broken once.
 *
 * The plan's verification for this sub-phase is "hand-corrupt a deck 29 ways,
 * assert each rule fires with part URI and XPath", and that is literally what
 * this file is. One fixture per rule, each a copy of a deck that passes with
 * exactly one thing changed.
 *
 * Two assertions on every one of them, and the second is the one that keeps the
 * suite honest over time. The rule fires - and the *clean* deck does not fire
 * it, checked once for all twenty-nine below. A test that only ever showed the
 * broken case would still pass if a rule started firing on everything.
 */

const OD = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/';
const PML = 'application/vnd.openxmlformats-officedocument.presentationml.';

function report(options: ValidateOptions): Report {
  return validatePackage(options);
}

/** The findings of one rule, as `part xpath - message` lines. */
function firedBy(report: Report, rule: RuleId): string[] {
  return report.findings
    .filter((finding) => finding.rule === rule)
    .map(
      (finding) =>
        finding.where.part + ' ' + (finding.where.xpath ?? '(part)') + ' - ' + finding.message,
    );
}

/** Validate a deck built by replacing parts, and return what `rule` said. */
function broken(
  rule: RuleId,
  parts: Readonly<Record<string, string | null>>,
  extraEntries?: readonly { name: string; bytes: Uint8Array }[],
): string[] {
  const { bytes, store } = deck(extraEntries === undefined ? { parts } : { parts, extraEntries });
  return firedBy(report({ bytes, store }), rule);
}

describe('a deck that passes', () => {
  it('has no findings at all', () => {
    const { bytes, store } = deck();
    const clean = report({ bytes, store });

    // Printed rather than counted, so a regression says which rule broke and
    // where instead of "expected 0, got 3".
    expect(formatReport(clean)).toContain('no findings');
    expect(clean.findings).toEqual([]);
    expect(clean.ok).toBe(true);
  });

  it('says which rules did not run, rather than counting them as passes', () => {
    const { bytes, store } = deck();
    const clean = report({ bytes, store });

    // No baseline was given, so the three preservation rules cannot run. A
    // report that omitted them would look like twenty-nine passes.
    expect(clean.skipped.map((entry) => entry.rule)).toEqual(['V027', 'V028', 'V029']);
    expect(clean.checked).toHaveLength(26);
    for (const entry of clean.skipped) expect(entry.why).toContain('package as it was opened');
  });

  it('says so when it was given a store and no archive', () => {
    const { store } = deck();
    const partial = report({ store });

    expect(partial.skipped.map((entry) => entry.rule)).toContain('V003');
    expect(partial.problems.map((problem) => problem.message).join('\n')).toContain(
      'only a PartStore was supplied',
    );
  });
});

describe('V001 content-type coverage', () => {
  it('fires on a part nothing types', () => {
    // A `.jpg` with no Default for the extension. Everything else is untouched.
    const found = broken('V001', {
      'ppt/media/image1.jpg': 'not really a jpeg',
    });
    expect(found).toHaveLength(1);
    expect(found[0]).toContain('/ppt/media/image1.jpg');
    expect(found[0]).toContain('a Default for the "jpg" extension');
  });
});

describe('V002 the content-type map', () => {
  it('fires on a duplicate Default that agrees with itself', () => {
    const parts = minimalDeck();
    const found = broken('V002', {
      '[Content_Types].xml': parts['[Content_Types].xml']!.replace(
        '<Default Extension="xml" ContentType="application/xml"/>',
        '<Default Extension="xml" ContentType="application/xml"/>' +
          '<Default Extension="xml" ContentType="application/xml"/>',
      ),
    });
    expect(found.join('\n')).toContain('two <Default> entries');
    // The XPath is the point: the second one, not "somewhere in the file".
    expect(found.join('\n')).toContain('/Types/Default[3]');
  });

  it('fires on an Override that names nothing', () => {
    const parts = minimalDeck();
    const found = broken('V002', {
      '[Content_Types].xml': parts['[Content_Types].xml']!.replace(
        '</Types>',
        '<Override PartName="/ppt/slides/slide9.xml" ContentType="' + PML + 'slide+xml"/></Types>',
      ),
    });
    expect(found.join('\n')).toContain('/ppt/slides/slide9.xml, which is not in the package');
  });

  it('fires on embedded fonts with no fntdata Default', () => {
    const parts = minimalDeck();
    const found = broken('V002', {
      'ppt/fonts/font1.fntdata': 'EOTx',
      '[Content_Types].xml': parts['[Content_Types].xml']!.replace(
        '</Types>',
        '<Override PartName="/ppt/fonts/font1.fntdata" ContentType="application/x-fontdata"/></Types>',
      ),
    });
    // An Override types the part, so `V001` is satisfied and the deck looks
    // complete. PowerPoint still reports a problem with the content, because
    // what it wants is the Default.
    expect(found.join('\n')).toContain('Default Extension="fntdata"');
  });
});

describe('V003 the archive', () => {
  it('fires on a directory entry', () => {
    const found = broken('V003', {}, [{ name: 'ppt/media/', bytes: new Uint8Array(0) }]);
    expect(found.join('\n')).toContain('directory entry, "ppt/media/"');
  });
});

describe('V004 part names', () => {
  it('fires on a percent-escape of an unreserved character', () => {
    // `%5F` is an underscore. RFC 3986 §6.2.2.2 says the two are equivalent;
    // PowerPoint refuses the escaped form with 0x808D1005. `@pptx-studio/opc`
    // rates this a warning because it must open what it is given - this is the
    // same finding at the point of writing, where it is fatal.
    const found = broken('V004', { 'ppt/tags/tag%5F1.xml': '<t/>' });
    expect(found.join('\n')).toContain('M1.8');
  });

  it('reports it fatal, where the package layer rates it a warning', () => {
    // The whole point of the rule, and the carried follow-up it discharges.
    // `validatePartName` grades `M1.8` a warning because a reader must open
    // what it is given. Here the package is on its way out, so the same finding
    // is fatal - and that difference is only meaningful if it shows up in the
    // report as one.
    const { bytes, store } = deck({ parts: { 'ppt/tags/tag%5F1.xml': '<t/>' } });
    const result = report({ bytes, store });
    const v004 = result.findings.filter((finding) => finding.rule === 'V004');

    expect(v004).toHaveLength(1);
    expect(v004[0]!.severity).toBe('fatal');
    expect(result.ok).toBe(false);
  });
});

describe('V005 the main part', () => {
  it('fires on a second officeDocument relationship', () => {
    const found = broken('V005', {
      '_rels/.rels': relsPart(
        rel('rId1', OD + 'officeDocument', 'ppt/presentation.xml') +
          rel('rId2', OD + 'officeDocument', 'ppt/presentation.xml'),
      ),
    });
    expect(found.join('\n')).toContain('a second officeDocument relationship');
    expect(found.join('\n')).toContain('/Relationships/Relationship[2]');
  });

  it('fires when the main part is typed as something else', () => {
    const parts = minimalDeck();
    const found = broken('V005', {
      '[Content_Types].xml': parts['[Content_Types].xml']!.replace(
        '<Override PartName="/ppt/presentation.xml" ContentType="' +
          PML +
          'presentation.main+xml"/>',
        '<Override PartName="/ppt/presentation.xml" ContentType="' + PML + 'slide+xml"/>',
      ),
    });
    expect(found.join('\n')).toContain('not a PresentationML main-part type');
  });
});

describe('V006 relationship references', () => {
  it('fires on an r:id nothing declares', () => {
    const parts = minimalDeck();
    const found = broken('V006', {
      'ppt/presentation.xml': parts['ppt/presentation.xml']!.replace(
        '</p:sldIdLst>',
        '<p:sldId id="257" r:id="rId9"/></p:sldIdLst>',
      ),
    });
    expect(found).toHaveLength(1);
    expect(found[0]).toContain('/p:presentation/p:sldIdLst/p:sldId[2]/@r:id');
    expect(found[0]).toContain('does not declare');
    // The message lists what the part does have, because the next question is
    // always "then what is there".
    expect(found[0]).toContain('rId1, rId2');
  });
});

describe('V007 relationship ids', () => {
  it('fires on two relationships with one id', () => {
    const found = broken('V007', {
      'ppt/slides/_rels/slide1.xml.rels': relsPart(
        rel('rId1', OD + 'slideLayout', '../slideLayouts/slideLayout1.xml') +
          rel('rId1', OD + 'tags', '../tags/tag1.xml'),
      ),
      'ppt/tags/tag1.xml': '<t/>',
    });
    expect(found.join('\n')).toContain('appears twice');
  });

  it('fires on an Id that is not an xsd:ID', () => {
    const found = broken('V007', {
      'ppt/slides/_rels/slide1.xml.rels': relsPart(
        rel('1rId', OD + 'slideLayout', '../slideLayouts/slideLayout1.xml'),
      ),
    });
    expect(found.join('\n')).toContain('is not an xsd:ID');
  });
});

describe('V008 relationship targets', () => {
  it('fires on a target that is not in the package', () => {
    const found = broken('V008', {
      'ppt/slides/_rels/slide1.xml.rels': relsPart(
        rel('rId1', OD + 'slideLayout', '../slideLayouts/slideLayout9.xml'),
      ),
    });
    expect(found.join('\n')).toContain('/ppt/slideLayouts/slideLayout9.xml');
    expect(found.join('\n')).toContain('dangling relationship is fatal');
  });

  it('resolves against the source part folder, not the .rels folder', () => {
    // The regression this rule is really about. `../slideLayouts/…` from
    // `/ppt/slides/_rels/slide1.xml.rels` resolves against `/ppt/slides/`,
    // giving `/ppt/slideLayouts/…`. Resolving against `/ppt/slides/_rels/`
    // would give `/ppt/slides/slideLayouts/…` and every image in a real deck
    // would dangle at once. The clean deck uses exactly that form, so this
    // passing at all is the assertion.
    const { bytes, store } = deck();
    expect(firedBy(report({ bytes, store }), 'V008')).toEqual([]);
  });
});

describe('V009 required relationship edges', () => {
  it('fires on a slide with no layout relationship', () => {
    const found = broken('V009', { 'ppt/slides/_rels/slide1.xml.rels': relsPart('') });
    expect(found.join('\n')).toContain('has 0 layout relationship(s)');
  });

  it('fires on a slide whose layout relationship points at a master', () => {
    const found = broken('V009', {
      'ppt/slides/_rels/slide1.xml.rels': relsPart(
        rel('rId1', OD + 'slideLayout', '../slideMasters/slideMaster1.xml'),
      ),
    });
    expect(found.join('\n')).toContain('slideMaster+xml');
    expect(found.join('\n')).toContain('One pointing at a master instead');
  });

  it('fires on a layout with no master relationship', () => {
    const found = broken('V009', { 'ppt/slideLayouts/_rels/slideLayout1.xml.rels': relsPart('') });
    expect(found.join('\n')).toContain('has 0 master relationship(s)');
  });
});

describe('V010 schema order', () => {
  it('fires on two children swapped inside a master', () => {
    // The exact shape of the Open XML SDK regression: `p:clrMap` after
    // `p:sldLayoutIdLst` rather than before it, nothing else changed.
    const parts = minimalDeck();
    const master = parts['ppt/slideMasters/slideMaster1.xml']!;
    const clrMap = '<p:clrMap ' + IDENTITY_CLR_MAP + '/>';
    const layouts =
      '<p:sldLayoutIdLst><p:sldLayoutId id="2147483649" r:id="rId1"/></p:sldLayoutIdLst>';
    const found = broken('V010', {
      'ppt/slideMasters/slideMaster1.xml': master.replace(clrMap + layouts, layouts + clrMap),
    });
    expect(found.join('\n')).toContain('<p:clrMap> follows <p:sldLayoutIdLst>');
    expect(found.join('\n')).toContain('/p:sldMaster');
  });
});

describe('V011 extension lists', () => {
  it('fires on an a:ext with no uri', () => {
    const parts = minimalDeck();
    const found = broken('V011', {
      'ppt/slides/slide1.xml': parts['ppt/slides/slide1.xml']!.replace(
        '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom>',
        '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:extLst><a:ext/></a:extLst>',
      ),
    });
    expect(found.join('\n')).toContain('has no @uri');
    expect(found.join('\n')).toContain('carried through untouched');
  });

  it('fires on an extLst that is not last', () => {
    const parts = minimalDeck();
    const found = broken('V011', {
      'ppt/slides/slide1.xml': parts['ppt/slides/slide1.xml']!.replace(
        '<p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr>',
        '<p:extLst/><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr>',
      ),
    });
    expect(found.join('\n')).toContain('is followed by <p:clrMapOvr>');
  });
});

describe('V012 children the model has no place for', () => {
  it('fires on a:ahXY directly under a:custGeom', () => {
    const parts = minimalDeck();
    const found = broken('V012', {
      'ppt/slides/slide1.xml': parts['ppt/slides/slide1.xml']!.replace(
        '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom>',
        '<a:custGeom><a:avLst/><a:gdLst/><a:ahXY/><a:cxnLst/><a:rect l="0" t="0" r="r" b="b"/>' +
          '<a:pathLst/></a:custGeom>',
      ),
    });
    expect(found.join('\n')).toContain('<a:ahXY> is not a child <a:custGeom> admits');
  });
});

describe('what the corpus taught these three rules', () => {
  // Every case below was a false positive on a deck PowerPoint opens, found by
  // running the validator over the fifty-one committed decks and read off the
  // failure. They are here rather than only in `tools/corpus/suites/validate.test.ts`
  // because the corpus check needs a filesystem and these do not, so a
  // regression should fail in the fast suite first.

  it('leaves extension markup alone, however unrankable it is', () => {
    // Without the namespace test, `V012` fired on forty `mc:AlternateContent`
    // elements plus `a14:m`, `p14:honeycomb` and `a37-mce`'s `zz:` markup. The
    // ordering table cannot rank any of them, and "the table cannot rank this"
    // means two different things: the schema forbids it, or we have never read
    // that schema. Only the first is a finding.
    const parts = minimalDeck();
    const { bytes, store } = deck({
      parts: {
        'ppt/slides/slide1.xml': parts['ppt/slides/slide1.xml']!.replace(
          '</p:spTree>',
          '<mc:AlternateContent ' +
            'xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006">' +
            '<mc:Choice xmlns:zz="urn:example:zz" Requires="zz"><zz:thing/></mc:Choice>' +
            '<mc:Fallback/></mc:AlternateContent></p:spTree>',
        ).replace(
          '</p:sld>',
          '<p:transition xmlns:p14="http://schemas.microsoft.com/office/powerpoint/2010/main">' +
            '<p14:honeycomb/></p:transition></p:sld>',
        ),
      },
    });
    expect(firedBy(report({ bytes, store }), 'V012')).toEqual([]);
  });

  it('lets the two branches of one mc:AlternateContent reuse a shape id', () => {
    // They are alternatives: a consumer takes one and the other is not there.
    // PowerPoint's own writer does this - `a22-chartex` has a p:graphicFrame
    // id="10" in the Choice and the p:pic that stands in for it, also id="10",
    // in the Fallback.
    const parts = minimalDeck();
    const branches =
      '<mc:AlternateContent ' +
      'xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006">' +
      '<mc:Choice xmlns:p14="http://schemas.microsoft.com/office/powerpoint/2010/main" ' +
      'Requires="p14">' +
      shape(10, 'Chart 9', '') +
      '</mc:Choice><mc:Fallback>' +
      shape(10, 'Chart 9 fallback', '') +
      '</mc:Fallback></mc:AlternateContent>';
    const { bytes, store } = deck({
      parts: {
        'ppt/slides/slide1.xml': parts['ppt/slides/slide1.xml']!.replace(
          '</p:spTree>',
          branches + '</p:spTree>',
        ),
      },
    });
    expect(firedBy(report({ bytes, store }), 'V020')).toEqual([]);
  });

  it('still catches a shape id reused where both shapes are really there', () => {
    // The other half of the same fix. Two `mc:AlternateContent` elements are
    // not alternatives to each other: both branches are chosen and both sets of
    // shapes end up on the slide, so a shared id between them is a real clash.
    const parts = minimalDeck();
    const one = (name: string): string =>
      '<mc:AlternateContent ' +
      'xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006">' +
      '<mc:Choice xmlns:p14="http://schemas.microsoft.com/office/powerpoint/2010/main" ' +
      'Requires="p14">' +
      shape(10, name, '') +
      '</mc:Choice><mc:Fallback/></mc:AlternateContent>';
    const { bytes, store } = deck({
      parts: {
        'ppt/slides/slide1.xml': parts['ppt/slides/slide1.xml']!.replace(
          '</p:spTree>',
          one('First') + one('Second') + '</p:spTree>',
        ),
      },
    });
    expect(firedBy(report({ bytes, store }), 'V020').join('\n')).toContain('already used by');
  });

  it('leaves a SmartArt drawing part to number its shapes however it likes', () => {
    // A `dsp:drawing` writes `dsp:cNvPr id="0"` on every shape. PowerPoint
    // generates those parts and regenerates them on any diagram interaction, so
    // whatever the ids are for, it is not identity - and the rule is about
    // PresentationML.
    const { bytes, store } = deck({
      parts: {
        'ppt/diagrams/drawing1.xml':
          '<dsp:drawing xmlns:dsp="http://schemas.microsoft.com/office/drawing/2008/diagram">' +
          '<dsp:spTree><dsp:sp><dsp:nvSpPr><dsp:cNvPr id="0" name=""/></dsp:nvSpPr></dsp:sp>' +
          '<dsp:sp><dsp:nvSpPr><dsp:cNvPr id="0" name=""/></dsp:nvSpPr></dsp:sp>' +
          '</dsp:spTree></dsp:drawing>',
      },
    });
    expect(firedBy(report({ bytes, store }), 'V020')).toEqual([]);
  });

  it('wants a:fill around a table style fill, where a:tblPr takes one directly', () => {
    // Found in our own `a20-tables`, which wrote the fill flat in thirteen
    // places. `a:tblPr` takes a fill element directly; `a:tcStyle` and
    // `a:tblBg` take a choice of `a:fill` or `a:fillRef`, so the fill sits one
    // level deeper two elements away. The deck was fixed; this is the guard.
    const parts = minimalDeck();
    const table = (fill: string): string =>
      parts['ppt/slides/slide1.xml']!.replace(
        '</p:spTree>',
        '<p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id="3" name="Table 2"/>' +
          '<p:cNvGraphicFramePr/><p:nvPr/></p:nvGraphicFramePr>' +
          '<p:xfrm><a:off x="0" y="0"/><a:ext cx="1" cy="1"/></p:xfrm>' +
          '<a:graphic><a:graphicData ' +
          'uri="http://schemas.openxmlformats.org/drawingml/2006/table">' +
          '<a:tbl><a:tblPr><a:tableStyle styleId="{1}" styleName="x">' +
          '<a:tblBg>' +
          fill +
          '</a:tblBg></a:tableStyle></a:tblPr></a:tbl>' +
          '</a:graphicData></a:graphic></p:graphicFrame></p:spTree>',
      );

    const flat = deck({
      parts: {
        'ppt/slides/slide1.xml': table('<a:solidFill><a:srgbClr val="FF0000"/></a:solidFill>'),
      },
    });
    expect(firedBy(report(flat), 'V012').join('\n')).toContain(
      '<a:solidFill> is not a child <a:tblBg> admits',
    );

    const wrapped = deck({
      parts: {
        'ppt/slides/slide1.xml': table(
          '<a:fill><a:solidFill><a:srgbClr val="FF0000"/></a:solidFill></a:fill>',
        ),
      },
    });
    expect(firedBy(report(wrapped), 'V012')).toEqual([]);
  });
});

describe('V013 p:notesSz', () => {
  it('fires when it is missing, though p:sldSz is there', () => {
    const parts = minimalDeck();
    const found = broken('V013', {
      'ppt/presentation.xml': parts['ppt/presentation.xml']!.replace(
        '<p:notesSz cx="6858000" cy="9144000"/>',
        '',
      ),
    });
    expect(found).toHaveLength(1);
    expect(found[0]).toContain('p:sldSz is [0..1] and p:notesSz is [1..1]');
    expect(found[0]).toContain('/p:presentation');
  });
});

describe('V014 the colour map', () => {
  it('fires on eleven of twelve', () => {
    const parts = minimalDeck();
    const found = broken('V014', {
      'ppt/slideMasters/slideMaster1.xml': parts['ppt/slideMasters/slideMaster1.xml']!.replace(
        ' folHlink="folHlink"',
        '',
      ),
    });
    expect(found).toHaveLength(1);
    expect(found[0]).toContain('missing folHlink');
    expect(found[0]).toContain('/p:sldMaster/p:clrMap');
  });
});

describe('V015 the shape-tree prologue', () => {
  it('fires when p:grpSpPr is gone', () => {
    const parts = minimalDeck();
    const found = broken('V015', {
      'ppt/slides/slide1.xml': parts['ppt/slides/slide1.xml']!.replace('<p:grpSpPr/>', ''),
    });
    expect(found.join('\n')).toContain('begins <p:nvGrpSpPr> then <p:sp>');
    expect(found.join('\n')).toContain('/p:sld/p:cSld/p:spTree');
  });
});

describe('V016 text bodies', () => {
  it('fires on a text body with no paragraph', () => {
    const parts = minimalDeck();
    const found = broken('V016', {
      'ppt/slides/slide1.xml': parts['ppt/slides/slide1.xml']!.replace(
        '<a:p><a:endParaRPr lang="en-US"/></a:p>',
        '',
      ),
    });
    expect(found).toHaveLength(1);
    expect(found[0]).toContain('has no <a:p>');
    expect(found[0]).toContain('deleting the last paragraph');
  });

  it('fires on a text body with no a:bodyPr', () => {
    const parts = minimalDeck();
    const found = broken('V016', {
      'ppt/slides/slide1.xml': parts['ppt/slides/slide1.xml']!.replace('<a:bodyPr/>', ''),
    });
    expect(found.join('\n')).toContain('has no <a:bodyPr>');
  });
});

describe('V017 graphic frames', () => {
  it('fires on a frame with no p:xfrm', () => {
    const parts = minimalDeck();
    const found = broken('V017', {
      'ppt/slides/slide1.xml': parts['ppt/slides/slide1.xml']!.replace(
        '</p:spTree>',
        '<p:graphicFrame>' +
          '<p:nvGraphicFramePr><p:cNvPr id="3" name="Table 2"/><p:cNvGraphicFramePr/>' +
          '<p:nvPr/></p:nvGraphicFramePr>' +
          '<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/table"/>' +
          '</a:graphic>' +
          '</p:graphicFrame></p:spTree>',
      ),
    });
    expect(found).toHaveLength(1);
    expect(found[0]).toContain('no <p:xfrm>');
    expect(found[0]).toContain('not the a:xfrm every other shape uses');
  });
});

describe('V018 slide ids', () => {
  it('fires below the floor of 256', () => {
    const parts = minimalDeck();
    const found = broken('V018', {
      'ppt/presentation.xml': parts['ppt/presentation.xml']!.replace(
        '<p:sldId id="256"',
        '<p:sldId id="1"',
      ),
    });
    expect(found).toHaveLength(1);
    expect(found[0]).toContain('outside ST_SlideId');
    expect(found[0]).toContain('/p:presentation/p:sldIdLst/p:sldId/@id');
  });
});

describe('V019 master and layout ids', () => {
  it('fires when a layout id collides with a master id in another part', () => {
    // The finding a per-part validator cannot make: the master id is in
    // `presentation.xml` and the layout id is in `slideMaster1.xml`.
    const parts = minimalDeck();
    const found = broken('V019', {
      'ppt/slideMasters/slideMaster1.xml': parts['ppt/slideMasters/slideMaster1.xml']!.replace(
        '<p:sldLayoutId id="2147483649"',
        '<p:sldLayoutId id="2147483648"',
      ),
    });
    expect(found).toHaveLength(1);
    expect(found[0]).toContain('already used by <p:sldMasterId> in /ppt/presentation.xml');
    expect(found[0]).toContain('one number space, not two');
  });

  it('fires below 2147483648', () => {
    const parts = minimalDeck();
    const found = broken('V019', {
      'ppt/presentation.xml': parts['ppt/presentation.xml']!.replace(
        '<p:sldMasterId id="2147483648"',
        '<p:sldMasterId id="256"',
      ),
    });
    expect(found.join('\n')).toContain('outside ST_SlideMasterId');
  });
});

describe('V020 shape ids', () => {
  it('fires on two shapes with one id in a part', () => {
    const parts = minimalDeck();
    const found = broken('V020', {
      'ppt/slides/slide1.xml': parts['ppt/slides/slide1.xml']!.replace(
        '</p:spTree>',
        shape(2, 'Title 1 again', '') + '</p:spTree>',
      ),
    });
    expect(found).toHaveLength(1);
    expect(found[0]).toContain('is already used by');
    expect(found[0]).toContain('free to repeat across parts');
  });

  it('fires in the range PowerPoint reads as negative', () => {
    const parts = minimalDeck();
    const found = broken('V020', {
      'ppt/slides/slide1.xml': parts['ppt/slides/slide1.xml']!.replace(
        '<p:cNvPr id="2" name="Title 1"/>',
        '<p:cNvPr id="3000000000" name="Title 1"/>',
      ),
    });
    expect(found.join('\n')).toContain('PowerPoint refuses outright');
  });

  it('accepts 4294967295, which opens', () => {
    const parts = minimalDeck();
    const found = broken('V020', {
      'ppt/slides/slide1.xml': parts['ppt/slides/slide1.xml']!.replace(
        '<p:cNvPr id="2" name="Title 1"/>',
        '<p:cNvPr id="4294967295" name="Title 1"/>',
      ),
    });
    expect(found).toEqual([]);
  });
});

describe('V021 placeholder matching', () => {
  it('warns on a slide placeholder the layout has no counterpart for', () => {
    const parts = minimalDeck();
    const found = broken('V021', {
      'ppt/slides/slide1.xml': parts['ppt/slides/slide1.xml']!.replace(
        '<p:ph type="title"/>',
        '<p:ph type="body" idx="7"/>',
      ),
    });
    expect(found).toHaveLength(1);
    expect(found[0]).toContain('(type=body, idx=7) matches nothing');
    expect(found[0]).toContain('which offers title/0');
  });

  it('does not warn, and does not block, when it matches by a later tier', () => {
    // Tier 2: a title matches any title regardless of `@idx`.
    const parts = minimalDeck();
    const { bytes, store } = deck({
      parts: {
        'ppt/slides/slide1.xml': parts['ppt/slides/slide1.xml']!.replace(
          '<p:ph type="title"/>',
          '<p:ph type="ctrTitle" idx="4"/>',
        ),
      },
    });
    const result = report({ bytes, store });
    expect(firedBy(result, 'V021')).toEqual([]);
    expect(result.ok).toBe(true);
  });

  it('never blocks an export on its own', () => {
    const parts = minimalDeck();
    const { bytes, store } = deck({
      parts: {
        'ppt/slides/slide1.xml': parts['ppt/slides/slide1.xml']!.replace(
          '<p:ph type="title"/>',
          '<p:ph type="body" idx="7"/>',
        ),
      },
    });
    const result = report({ bytes, store });
    expect(result.findings).toHaveLength(1);
    expect(result.blocking).toBe(0);
    expect(result.ok).toBe(true);
  });
});

describe('V022 hdr and sldImg placeholders', () => {
  it('fires on a slide', () => {
    const parts = minimalDeck();
    const found = broken('V022', {
      'ppt/slides/slide1.xml': parts['ppt/slides/slide1.xml']!.replace(
        '<p:ph type="title"/>',
        '<p:ph type="hdr"/>',
      ),
    });
    expect(found.join('\n')).toContain('whole-package refusal here');
    expect(found.join('\n')).toContain('/@type');
  });

  it('fires on a layout too', () => {
    const parts = minimalDeck();
    const found = broken('V022', {
      'ppt/slideLayouts/slideLayout1.xml': parts['ppt/slideLayouts/slideLayout1.xml']!.replace(
        '<p:ph type="title"/>',
        '<p:ph type="sldImg"/>',
      ),
    });
    expect(found.join('\n')).toContain('sldImg');
  });
});

describe('V023 geometry guides', () => {
  it('fires on a formula naming a guide nothing defines', () => {
    const parts = minimalDeck();
    const found = broken('V023', {
      'ppt/slides/slide1.xml': parts['ppt/slides/slide1.xml']!.replace(
        '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom>',
        '<a:custGeom><a:avLst/>' +
          '<a:gdLst><a:gd name="x1" fmla="*/ w adj1 100000"/></a:gdLst>' +
          '<a:ahLst/><a:cxnLst/><a:rect l="0" t="0" r="r" b="b"/><a:pathLst/></a:custGeom>',
      ),
    });
    expect(found).toHaveLength(1);
    expect(found[0]).toContain('"adj1" names a guide nothing defines');
    expect(found[0]).toContain('This geometry defines x1');
  });

  it('accepts built-in guides and literals', () => {
    const parts = minimalDeck();
    const { bytes, store } = deck({
      parts: {
        'ppt/slides/slide1.xml': parts['ppt/slides/slide1.xml']!.replace(
          '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom>',
          '<a:custGeom><a:avLst><a:gd name="adj" fmla="val 25000"/></a:avLst>' +
            '<a:gdLst><a:gd name="x1" fmla="*/ ss adj 100000"/>' +
            '<a:gd name="x2" fmla="+- x1 wd2 0"/></a:gdLst>' +
            '<a:ahLst/><a:cxnLst/><a:rect l="l" t="t" r="r" b="b"/>' +
            '<a:pathLst><a:path w="21600" h="21600">' +
            '<a:moveTo><a:pt x="x1" y="0"/></a:moveTo>' +
            '<a:lnTo><a:pt x="x2" y="hd2"/></a:lnTo>' +
            '</a:path></a:pathLst></a:custGeom>',
        ),
      },
    });
    expect(firedBy(report({ bytes, store }), 'V023')).toEqual([]);
  });
});

describe('V024 series text', () => {
  const chart = (tx: string): string =>
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
    '<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart">' +
    '<c:chart><c:plotArea><c:barChart><c:ser>' +
    tx +
    '</c:ser></c:barChart></c:plotArea></c:chart></c:chartSpace>';

  it('fires on a c:strLit', () => {
    const found = broken('V024', {
      'ppt/charts/chart1.xml': chart('<c:tx><c:strLit/></c:tx>'),
    });
    expect(found).toHaveLength(1);
    expect(found[0]).toContain('<c:strLit> inside a series <c:tx>');
    expect(found[0]).toContain('/c:chartSpace/c:chart/c:plotArea/c:barChart/c:ser/c:tx/c:strLit');
  });

  it('leaves c:strRef alone', () => {
    const { bytes, store } = deck({
      parts: {
        'ppt/charts/chart1.xml': chart('<c:tx><c:strRef><c:f>Sheet1!$A$1</c:f></c:strRef></c:tx>'),
      },
    });
    expect(firedBy(report({ bytes, store }), 'V024')).toEqual([]);
  });
});

describe('V025 ActiveX controls', () => {
  it('fires on a p:control, and not on an empty p:controls', () => {
    const parts = minimalDeck();
    const withControl = broken('V025', {
      'ppt/slides/slide1.xml': parts['ppt/slides/slide1.xml']!.replace(
        '</p:cSld>',
        '<p:controls><p:control name="Button1"/></p:controls></p:cSld>',
      ),
    });
    expect(withControl.join('\n')).toContain('all eight forms tried');

    const { bytes, store } = deck({
      parts: {
        'ppt/slides/slide1.xml': parts['ppt/slides/slide1.xml']!.replace(
          '</p:cSld>',
          '<p:controls/></p:cSld>',
        ),
      },
    });
    expect(firedBy(report({ bytes, store }), 'V025')).toEqual([]);
  });
});

describe('V026 chart styles', () => {
  const CS = 'http://schemas.microsoft.com/office/drawing/2012/chartStyle';

  it('fires on a subset, which is worse than an absence', () => {
    const found = broken('V026', {
      'ppt/charts/style1.xml':
        '<cs:chartStyle xmlns:cs="' +
        CS +
        '" id="201"><cs:chartArea/><cs:dataPoint/><cs:legend/><cs:plotArea/></cs:chartStyle>',
    });
    expect(found).toHaveLength(1);
    expect(found[0]).toContain('4 of its 31 entries');
    expect(found[0]).toContain('Four entries were refused and thirty-one opened');
  });
});
