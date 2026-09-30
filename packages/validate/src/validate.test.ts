import { attributeValue, childElements, parseXmlString, type XElement } from '@pptx-studio/xml';
import { beforeAll, describe, expect, it } from 'vitest';

import cascade from '../../../corpus/ground-truth/table-cascade.json' with { type: 'json' };
import tableStyles from '../../../corpus/ground-truth/table-styles.json' with { type: 'json' };
import tables from '../../../corpus/ground-truth/tables.json' with { type: 'json' };

import { formatReport, type Report } from './report/report.js';
import { RULES, type RuleId } from './rules/rules.js';
import { deck, minimalDeck, relsPart, rel, shape, IDENTITY_CLR_MAP } from './testing/deck.js';
import { validatePackage, type ValidateOptions } from './validate.js';

/**
 * Every rule, broken once: a deck that passes, with exactly one thing changed per rule.
 * The clean deck is checked once for every rule too, so a rule that fires on everything fails.
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

    // No baseline was given, so the three preservation rules are named as skipped, not passed.
    expect(clean.skipped.map((entry) => entry.rule)).toEqual(['V027', 'V028', 'V029']);
    expect(clean.checked).toHaveLength(RULES.length - 3);
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

describe('V030 and V031 tables, against what PowerPoint wrote back in C7', () => {
  const TABLE_URI = 'http://schemas.openxmlformats.org/drawingml/2006/table';
  /** The C5 test slide with a probe's table on it. */
  const withTable = (markup: string, xfrm = true): Record<string, string> => {
    const parts = minimalDeck();
    return {
      'ppt/slides/slide1.xml': parts['ppt/slides/slide1.xml']!.replace(
        '</p:spTree>',
        '<p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id="3" name="table"/>' +
          '<p:cNvGraphicFramePr><a:graphicFrameLocks noGrp="1"/></p:cNvGraphicFramePr><p:nvPr/>' +
          '</p:nvGraphicFramePr>' +
          (xfrm
            ? '<p:xfrm><a:off x="914400" y="914400"/><a:ext cx="5486400" cy="1371600"/></p:xfrm>'
            : '') +
          '<a:graphic><a:graphicData uri="' +
          TABLE_URI +
          '">' +
          markup +
          '</a:graphicData></a:graphic>' +
          '</p:graphicFrame></p:spTree>',
      ),
    };
  };
  const NS =
    'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" ' +
    'xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"';

  /** A row's cells as PowerPoint would compare them: spans as counts, flags as booleans. */
  type Cell = readonly [number, number, boolean, boolean];
  const normal = (attrs: Readonly<Record<string, string | undefined>>): Cell => [
    Number(attrs['gridSpan'] ?? 1),
    Number(attrs['rowSpan'] ?? 1),
    attrs['hMerge'] === '1' || attrs['hMerge'] === 'true',
    attrs['vMerge'] === '1' || attrs['vMerge'] === 'true',
  ];
  const writtenGrid = (markup: string): Cell[][] => {
    const tbl = parseXmlString('<x ' + NS + '>' + markup + '</x>').root.children.find(
      (n): n is XElement => n.type === 'element',
    )!;
    return childElements(tbl)
      .filter((child) => child.qname === 'a:tr')
      .map((row) =>
        childElements(row)
          .filter((child) => child.qname === 'a:tc')
          .map((cell) =>
            normal({
              gridSpan: attributeValue(cell, 'gridSpan'),
              rowSpan: attributeValue(cell, 'rowSpan'),
              hMerge: attributeValue(cell, 'hMerge'),
              vMerge: attributeValue(cell, 'vMerge'),
            }),
          ),
      );
  };
  const asWritten = tables.probes.filter((p) => p.repaired === false && p.resaved !== null);

  it('V030 fires on exactly the probes whose grid PowerPoint read back differently', () => {
    // Differently in what it reads: a lexical form normalised, an explicit 1 dropped, or its
    // own repeated span added to a flagged cell is the same grid written PowerPoint's way.
    const disagreements: string[] = [];
    for (const probe of asWritten) {
      const written = writtenGrid(probe.markup);
      const resaved = probe.resaved.rows.map((row) => row.cells.map((cell) => normal(cell)));
      // A span PowerPoint added to a flagged cell is its own redundant form, not a disagreement.
      const padded = resaved.map((row, r) =>
        row.map((cell, c): Cell => {
          const before = written[r]?.[c];
          if (before === undefined || !(cell[2] || cell[3])) return cell;
          return [before[0] === 1 ? 1 : cell[0], before[1] === 1 ? 1 : cell[1], cell[2], cell[3]];
        }),
      );
      const rewritten =
        probe.resaved.cols.length !== (probe.markup.match(/<a:gridCol\b/g) ?? []).length ||
        JSON.stringify(written) !== JSON.stringify(padded);
      const fired = broken('V030', withTable(probe.markup)).length > 0;
      if (fired !== rewritten) {
        disagreements.push(
          probe.id + ': ' + (rewritten ? 'rewritten and silent' : 'kept and fired'),
        );
      }
    }
    expect(disagreements).toEqual([]);
    expect(asWritten.length).toBe(76);
  });

  it('V030 names the disagreement', () => {
    const probe = (id: string): string => tables.probes.find((p) => p.id === id)!.markup;
    expect(broken('V030', withTable(probe('span-no-flag-h'))).join('\n')).toContain(
      'cell (1, 2) is covered by the span of cell (1, 1) and does not say hMerge="1"',
    );
    expect(broken('V030', withTable(probe('flag-no-span-h'))).join('\n')).toContain(
      'cell (1, 2) says hMerge="1" and no span covers it',
    );
    expect(broken('V030', withTable(probe('cross'))).join('\n')).toContain(
      'cell (2, 1) says gridSpan="2" and covers 1',
    );
    expect(broken('V030', withTable(probe('row-short'))).join('\n')).toContain(
      'row 1 holds 3 cell(s) for 4 column(s)',
    );
    expect(broken('V030', withTable(probe('no-cols'))).join('\n')).toContain(
      '0 column(s) and 3 row(s)',
    );
    expect(broken('V030', withTable(probe('covered-anchor-v'))).join('\n')).toContain(
      'cell (1, 2) is covered and says rowSpan="2"',
    );
    expect(broken('V030', withTable(probe('cross'))).join('\n')).toContain(
      'cell (2, 2) says hMerge="1" and the span of cell (1, 2) covers it the other way',
    );
  });

  it('V030 is silent on the form PowerPoint writes for its own merges and splits', () => {
    for (const slide of tables.findings.canonical) {
      const cols = Math.max(...slide.rows.map((row) => row.length));
      const markup =
        '<a:tbl><a:tblPr/><a:tblGrid>' +
        '<a:gridCol w="1371600"/>'.repeat(cols) +
        '</a:tblGrid>' +
        slide.rows
          .map(
            (row) =>
              '<a:tr h="457200">' +
              row
                .map(
                  (cell) =>
                    '<a:tc' +
                    Object.entries(cell)
                      .map(([k, v]) => ' ' + k + '="' + String(v) + '"')
                      .join('') +
                    '><a:txBody><a:bodyPr/><a:lstStyle/><a:p/></a:txBody><a:tcPr/></a:tc>',
                )
                .join('') +
              '</a:tr>',
          )
          .join('') +
        '</a:tbl>';
      expect(broken('V030', withTable(markup)), slide.name).toEqual([]);
    }
    expect(tables.findings.canonical.length).toBe(12);
  });

  it('V031 fires on the four lexical repairs and on nothing PowerPoint opened as written', () => {
    const probe = (id: string) => tables.probes.find((p) => p.id === id)!;
    for (const [id, text] of [
      ['gridcol-no-w', '<a:gridCol> has no @w'],
      ['tr-no-h', '<a:tr> has no @h'],
      ['merge-on', '<a:tc>/@hMerge is "on", which is not an xsd:boolean'],
      ['span-float', '<a:tc>/@gridSpan is "2.0", which is not an xsd:int'],
    ] as const) {
      const found = broken('V031', withTable(probe(id).markup));
      expect(found, id).toHaveLength(1);
      expect(found[0], id).toContain(text);
    }
    for (const p of asWritten) expect(broken('V031', withTable(p.markup)), p.id).toEqual([]);
    expect(broken('V016', withTable(probe('tc-body-no-p').markup)).join('\n')).toContain(
      'has no <a:p>',
    );
    expect(broken('V017', withTable(probe('frame-missing').markup, false)).join('\n')).toContain(
      'no <p:xfrm>',
    );
  });
});

describe('V032 and V033 table styles, against what PowerPoint drew and wrote back in C8', () => {
  type StyleProbe = (typeof tableStyles.probes)[number];
  const single = tableStyles.probes.filter(
    (p) => !(p.markup.tblPr ?? '').includes('slides, one per'),
  );
  /** The controls a single-table probe's picture matched. */
  const matchesOf = (probe: StyleProbe): readonly string[] => {
    const drawn: unknown = probe.drawn;
    if (typeof drawn !== 'object' || drawn === null || !('matches' in drawn)) return [];
    const matches: unknown = drawn.matches;
    return Array.isArray(matches) ? matches.filter((m): m is string => typeof m === 'string') : [];
  };
  interface Markup {
    readonly markup: {
      readonly tblPr: string | null;
      readonly tableStyles: string | null;
      readonly partName: string | null;
      readonly rel: boolean;
    };
  }
  /** The minimal deck with the probe's table on its slide and its table-style part beside it. */
  const withStyles = (probe: Markup): Record<string, string> => {
    const parts = minimalDeck();
    const table =
      '<a:tbl>' +
      (probe.markup.tblPr ?? '') +
      '<a:tblGrid><a:gridCol w="1371600"/></a:tblGrid><a:tr h="457200"><a:tc><a:txBody>' +
      '<a:bodyPr/><a:lstStyle/><a:p/></a:txBody><a:tcPr/></a:tc></a:tr></a:tbl>';
    const out: Record<string, string> = {
      'ppt/slides/slide1.xml': parts['ppt/slides/slide1.xml']!.replace(
        '</p:spTree>',
        '<p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id="3" name="table"/>' +
          '<p:cNvGraphicFramePr><a:graphicFrameLocks noGrp="1"/></p:cNvGraphicFramePr><p:nvPr/>' +
          '</p:nvGraphicFramePr><p:xfrm><a:off x="914400" y="914400"/>' +
          '<a:ext cx="1371600" cy="457200"/></p:xfrm><a:graphic>' +
          '<a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/table">' +
          table +
          '</a:graphicData></a:graphic></p:graphicFrame></p:spTree>',
      ),
    };
    const name = (probe.markup.partName ?? '/ppt/tableStyles.xml').slice(1);
    if (probe.markup.tableStyles !== null) {
      out[name] = probe.markup.tableStyles;
      out['[Content_Types].xml'] = parts['[Content_Types].xml']!.replace(
        '</Types>',
        '<Override PartName="/' + name + '" ContentType="' + PML + 'tableStyles+xml"/></Types>',
      );
    }
    if (probe.markup.rel) {
      out['ppt/_rels/presentation.xml.rels'] = relsPart(
        rel('rId1', OD + 'slideMaster', 'slideMasters/slideMaster1.xml') +
          rel('rId2', OD + 'slide', 'slides/slide1.xml') +
          rel('rId3', OD + 'tableStyles', name.slice('ppt/'.length)),
      );
    }
    return out;
  };
  /** What PowerPoint drew for a table naming no built-in, in each theme C8 drew in. */
  const DEFAULT_GRID: Readonly<Record<string, string>> = {
    s1: 'sweep-ctl#7',
    s2: 'ctl2-black-grid',
    s3: 'ctl3-black-grid',
  };
  const clean = single.filter((p) => p.repaired === false);
  const repaired = single.filter((p) => p.repaired === true);

  it('V032 fires on exactly the tables that name an id and draw the default grid', () => {
    const disagreements: string[] = [];
    for (const probe of clean) {
      const names = /<a:tableStyle(Id>|\s)/.test(probe.markup.tblPr ?? '');
      const grid = matchesOf(probe).includes(DEFAULT_GRID[probe.theme] ?? '');
      const fired = broken('V032', withStyles(probe)).length > 0;
      if (fired !== (names && grid)) disagreements.push(probe.id + (fired ? ' fired' : ' silent'));
    }
    expect(disagreements).toEqual([]);
    expect(clean.filter((p) => broken('V032', withStyles(p)).length > 0).length).toBeGreaterThan(
      20,
    );
  });

  it('V032 names the id, and is silent on a built-in written in lower case', () => {
    const probe = (id: string) => single.find((p) => p.id === id)!;
    expect(broken('V032', withStyles(probe('custom-X'))).join('\n')).toContain(
      "names {C8000000-0000-4000-8000-00000000A001}, none of PowerPoint's 74 built-in table styles",
    );
    expect(broken('V032', withStyles(probe('lex-lower-G2')))).toEqual([]);
    expect(broken('V032', withStyles(probe('inline-X'))).join('\n')).toContain('<a:tableStyle>');
  });

  it('V033 fires on the table-style forms PowerPoint repaired, and on nothing it opened as written', () => {
    const probe = (id: string) => single.find((p) => p.id === id)!;
    for (const [id, text] of [
      ['lex-nobrace-G2', 'which is not a GUID in braces'],
      ['lex-space-G2', 'which is not a GUID in braces'],
      ['lex-newline-G2', 'which is not a GUID in braces'],
      ['def-missing', '<a:tblStyleLst> has no @def'],
      ['styleid-missing', '<a:tblStyle> has no @styleId'],
      ['b-yes', '<a:tcTxStyle>/@b is "yes", not on, off or def'],
      ['edge-empty', '<a:left> has neither <a:ln> nor <a:lnRef>'],
      ['fill-empty', '<a:fill> holds no fill'],
    ] as const) {
      const found = broken('V033', withStyles(probe(id)));
      expect(found, id).toHaveLength(1);
      expect(found[0], id).toContain(text);
    }
    // ST_Guid wants both braces; the nobrace probe measured the pair, this holds each alone.
    const g2 = probe('lex-nobrace-G2');
    for (const id of [
      'E8034E78-7F5D-4C2E-B375-FC64B27BC917}',
      '{E8034E78-7F5D-4C2E-B375-FC64B27BC917',
    ]) {
      const oneBrace = {
        ...g2,
        markup: {
          ...g2.markup,
          tblPr: `<a:tblPr><a:tableStyleId>${id}</a:tableStyleId></a:tblPr>`,
        },
      };
      expect(broken('V033', withStyles(oneBrace)), id).toHaveLength(1);
    }
    for (const p of clean) expect(broken('V033', withStyles(p)), p.id).toEqual([]);
  });

  it('refuses every package PowerPoint repaired', () => {
    const unrefused: string[] = [];
    for (const probe of repaired) {
      const { bytes, store } = deck({ parts: withStyles(probe) });
      const fatal = report({ bytes, store }).findings.filter((f) => f.severity === 'fatal');
      if (fatal.length === 0) unrefused.push(probe.id);
    }
    expect(unrefused).toEqual([]);
    expect(repaired.length).toBe(11);
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

describe('V034 table borders, against what PowerPoint drew and wrote in C9', () => {
  type Cells = Readonly<Record<string, { readonly attrs?: string; readonly tcPr?: string }>>;
  interface Slide {
    readonly name?: string;
    readonly rows: number;
    readonly cols: number;
    readonly cells?: Cells;
    readonly codes: string;
  }
  interface Deck {
    readonly id: string;
    readonly group: string;
    readonly read: string;
    readonly width: number;
    readonly values: readonly string[];
    readonly slides: readonly Slide[];
  }
  interface Extras {
    readonly name?: string;
    readonly cells?: Cells;
  }
  interface Packed extends Omit<Deck, 'slides'> {
    readonly shapes: string;
    /** What else each slide says, by index: a key of the file's `templates`, or inline. */
    readonly extras?: Readonly<Record<string, Extras | string>>;
    readonly codes: string;
  }
  const templates = cascade.templates as unknown as Readonly<Record<string, Extras>>;
  const said = (extras: Extras | string | undefined): Extras | undefined => {
    if (typeof extras !== 'string') return extras;
    const found = templates[extras];
    if (found === undefined) throw new Error(`no template ${extras}`);
    return found;
  };
  /** The fixture's decks, each slide's codes cut from the deck's raw DEFLATE stream by its shape. */
  const unpack = async (deck: Packed): Promise<Deck> => {
    const bytes = Uint8Array.from(atob(deck.codes), (ch) => ch.charCodeAt(0));
    const inflated = new Blob([bytes]).stream().pipeThrough(new DecompressionStream('deflate-raw'));
    const codes = await new Response(inflated).text();
    let at = 0;
    const slides = deck.shapes.split(' ').map((shape, k): Slide => {
      const [, rows = 0, cols = 0] = (/^(\d+)x(\d+)\//.exec(shape) ?? []).map(Number);
      const units =
        rows * cols * (deck.read === 'full' ? 4 : 3) + (rows + 1) * cols + rows * (cols + 1);
      const own = codes.slice(at, at + units * deck.width);
      at += units * deck.width;
      return { ...said(deck.extras?.[String(k)]), rows, cols, codes: own };
    });
    if (at !== codes.length) throw new Error(`${deck.id}: ${String(codes.length - at)} codes over`);
    return { ...deck, slides };
  };
  let decks: readonly Deck[] = [];
  let direct: readonly Deck[] = [];
  beforeAll(async () => {
    decks = await Promise.all((cascade.decks as unknown as readonly Packed[]).map(unpack));
    direct = decks.filter((d) => d.group === 'direct');
  });
  const TABLE_URI = 'http://schemas.openxmlformats.org/drawingml/2006/table';

  /** The minimal deck with a table of `rows` x `cols` whose cells carry `cells`' markup. */
  const withTable = (
    rows: number,
    cols: number,
    cells: Readonly<Record<string, { readonly attrs?: string; readonly tcPr?: string }>>,
  ): Record<string, string> => {
    const body = '<a:txBody><a:bodyPr/><a:lstStyle/><a:p/></a:txBody>';
    const trs = Array.from({ length: rows }, (_, r) => {
      const tcs = Array.from({ length: cols }, (_, c) => {
        const cell = cells[`${String(r)},${String(c)}`] ?? {};
        return `<a:tc${cell.attrs ?? ''}>${body}${cell.tcPr ?? '<a:tcPr/>'}</a:tc>`;
      });
      return `<a:tr h="457200">${tcs.join('')}</a:tr>`;
    });
    const tbl =
      `<a:tbl><a:tblPr/><a:tblGrid>${'<a:gridCol w="914400"/>'.repeat(cols)}</a:tblGrid>` +
      `${trs.join('')}</a:tbl>`;
    const parts = minimalDeck();
    return {
      'ppt/slides/slide1.xml': parts['ppt/slides/slide1.xml']!.replace(
        '</p:spTree>',
        '<p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id="3" name="table"/><p:cNvGraphicFramePr/>' +
          '<p:nvPr/></p:nvGraphicFramePr><p:xfrm><a:off x="914400" y="914400"/>' +
          '<a:ext cx="4572000" cy="2286000"/></p:xfrm><a:graphic>' +
          `<a:graphicData uri="${TABLE_URI}">${tbl}</a:graphicData></a:graphic></p:graphicFrame></p:spTree>`,
      ),
    };
  };

  /** Each grid edge's pixels, as the fixture holds them: `a|b|descriptor`, horizontal edges first. */
  const edgesOf = (deck: Deck, slide: Slide): Map<string, string> => {
    const perCell = deck.read === 'full' ? 4 : 3;
    const digits = cascade.encoding.digits;
    const valueAt = (k: number): string => {
      const code = slide.codes.slice(k * deck.width, (k + 1) * deck.width);
      const n = [...code].reduce((acc, ch) => acc * digits.length + digits.indexOf(ch), 0);
      return deck.values[n] ?? '';
    };
    const keys: string[] = [];
    for (let r = 0; r <= slide.rows; r++)
      for (let c = 0; c < slide.cols; c++) keys.push(`h${String(r)},${String(c)}`);
    for (let r = 0; r < slide.rows; r++)
      for (let c = 0; c <= slide.cols; c++) keys.push(`v${String(r)},${String(c)}`);
    const first = slide.rows * slide.cols * perCell;
    return new Map(keys.map((k, i) => [k, valueAt(first + i)]));
  };

  /**
   * Every coloured side an anchor cell writes, and whether PowerPoint drew its colour on every segment
   * of it. A covered cell's own `a:tcPr` is never read (C9), and `V034` does not look at it.
   */
  const writtenLines = (
    deck: Deck,
    slide: Slide,
  ): { readonly key: string; readonly drawn: boolean }[] => {
    const edges = edgesOf(deck, slide);
    const out: { key: string; drawn: boolean }[] = [];
    for (const [at, cell] of Object.entries(slide.cells ?? {})) {
      if (/Merge=/.test(cell.attrs ?? '')) continue;
      const [r, c] = at.split(',').map(Number) as [number, number];
      const span = (name: string): number =>
        Number(new RegExp(`${name}="(\\d+)"`).exec(cell.attrs ?? '')?.[1] ?? 1);
      const [rows, cols] = [span('rowSpan'), span('gridSpan')];
      for (const tag of ['lnT', 'lnL', 'lnB', 'lnR'] as const) {
        const line = new RegExp(`<a:${tag}\\b[^>]*>.*?</a:${tag}>`).exec(cell.tcPr ?? '')?.[0];
        const colour =
          line === undefined ? undefined : /<a:srgbClr val="([0-9A-F]{6})"/.exec(line)?.[1];
        if (colour === undefined) continue;
        const across = tag === 'lnT' || tag === 'lnB' ? cols : rows;
        const segments = Array.from({ length: across }, (_, k) => {
          if (tag === 'lnT') return `h${String(r)},${String(c + k)}`;
          if (tag === 'lnB') return `h${String(r + rows)},${String(c + k)}`;
          if (tag === 'lnL') return `v${String(r + k)},${String(c)}`;
          return `v${String(r + k)},${String(c + cols)}`;
        });
        out.push({
          key: `cell (${String(r + 1)}, ${String(c + 1)}) writes <a:${tag}>`,
          drawn: segments.every((s) =>
            (edges.get(s) ?? '').split('|').slice(2).join('|').includes(colour),
          ),
        });
      }
    }
    return out;
  };

  it('fires on exactly the lines PowerPoint did not draw, in direct and merged cells alike', () => {
    const disagreements: string[] = [];
    let fired = 0;
    let silent = 0;
    for (const deck of decks.filter((d) => d.group === 'direct' || d.group === 'merge')) {
      for (const slide of deck.slides) {
        const findings = broken('V034', withTable(slide.rows, slide.cols, slide.cells ?? {}));
        const found = findings.join('\n');
        const lines = writtenLines(deck, slide);
        for (const { key, drawn } of lines) {
          if (drawn) silent += 1;
          else fired += 1;
          if (found.includes(key) === drawn)
            disagreements.push(`${deck.id} ${String(slide.name)}: ${key}, drawn ${String(drawn)}`);
        }
        // One finding per line not drawn, and nothing else.
        if (findings.length !== lines.filter((l) => !l.drawn).length)
          disagreements.push(
            `${deck.id} ${String(slide.name)}: ${String(findings.length)} findings`,
          );
      }
    }
    expect(disagreements).toEqual([]);
    expect(fired).toBeGreaterThan(20);
    expect(silent).toBeGreaterThan(0);
  });

  it('is silent beside a merged neighbour’s segment that is not level with its anchor', () => {
    const merged = decks.filter((d) => d.group === 'merge');
    expect(merged).toHaveLength(4);
    for (const deck of merged) {
      const slide = deck.slides.find((s) => s.name === 'right-authored')!;
      const lines = writtenLines(deck, slide);
      expect(lines).toEqual([
        { key: 'cell (2, 2) writes <a:lnR>', drawn: false },
        { key: 'cell (2, 3) writes <a:lnL>', drawn: false },
        { key: 'cell (3, 3) writes <a:lnL>', drawn: true },
      ]);
      const found = broken('V034', withTable(5, 5, slide.cells ?? {})).join('\n');
      expect(found).toContain('cell (2, 3) writes <a:lnL>');
      expect(found).not.toContain('cell (3, 3) writes <a:lnL>');
    }
  });

  it('is silent where both sides draw nothing, however each says so', () => {
    // A line with no fill draws nothing, as an a:noFill does (C9's partial lines).
    const cells = {
      '1,1': { tcPr: '<a:tcPr><a:lnB w="12700"><a:noFill/></a:lnB></a:tcPr>' },
      '2,1': { tcPr: '<a:tcPr><a:lnT w="38100" cmpd="dbl"/></a:tcPr>' },
    };
    expect(broken('V034', withTable(4, 4, cells))).toEqual([]);
    const drawnAbove = {
      ...cells,
      '1,1': {
        tcPr:
          '<a:tcPr><a:lnB w="12700">' +
          '<a:solidFill><a:srgbClr val="FF0000"/></a:solidFill></a:lnB></a:tcPr>',
      },
    };
    expect(broken('V034', withTable(4, 4, drawnAbove))).toHaveLength(1);
  });

  it('names the cell, the segment, and the cell whose line PowerPoint draws there', () => {
    const lower = direct[0]!.slides.find((s) => s.name === 'lower-only')!;
    expect(broken('V034', withTable(5, 5, lower.cells ?? {})).join('\n')).toContain(
      'cell (3, 3) writes <a:lnT>, but above column 3 PowerPoint draws the <a:lnB> of cell (2, 3), ' +
        'which writes none - measured in C9.',
    );
    const wide = decks
      .find((d) => d.group === 'merge')!
      .slides.find((s) => s.name === 'wide-bottom')!;
    expect(broken('V034', withTable(5, 5, wide.cells ?? {})).join('\n')).toContain(
      'cell (2, 2) writes <a:lnB>, but below column 3 PowerPoint draws the <a:lnT> of cell (3, 3), ' +
        'which writes none - measured in C9.',
    );
  });

  it('is silent on the markup PowerPoint wrote when it set each border itself', () => {
    const authored = cascade.authored as unknown as readonly {
      readonly name: string;
      readonly rows: readonly (readonly { readonly attrs: string; readonly tcPr: string }[])[];
    }[];
    expect(authored.length).toBe(12);
    for (const op of authored) {
      const cells: Record<string, { attrs: string; tcPr: string }> = {};
      op.rows.forEach((row, r) =>
        row.forEach((cell, c) => {
          cells[`${String(r)},${String(c)}`] = {
            attrs: cell.attrs === '' ? '' : ` ${cell.attrs}`,
            tcPr: cell.tcPr,
          };
        }),
      );
      expect(
        broken('V034', withTable(op.rows.length, op.rows[0]?.length ?? 0, cells)),
        op.name,
      ).toEqual([]);
    }
  });
});
