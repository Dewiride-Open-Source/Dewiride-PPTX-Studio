import { PartStore } from '@pptx-studio/opc';
import { describe, expect, it } from 'vitest';
import type { Report } from './report.js';
import type { RuleId } from './rules.js';
import { deck, deckBytes, minimalDeck } from './testing/deck.js';
import { validatePackage } from './validate.js';

/**
 * `V027` … `V029`, which need two packages to say anything.
 *
 * Two shapes of test, and they are not interchangeable, because the two failure
 * modes they model are not.
 *
 * **A writer that re-serialised a clean part** is a package whose bytes differ
 * from the original in a part nobody edited. Modelled by opening two stores
 * from two different archives: nothing is marked edited in either, so any
 * difference is one the writer invented.
 *
 * **A command that rewrote markup it should have carried through** is a package
 * where a part *was* edited and the edit reached into something held opaque.
 * Modelled by opening one store and calling `replacePart`, which is exactly
 * what an edit does.
 */

const encoder = new TextEncoder();

function firedBy(report: Report, rule: RuleId): string[] {
  return report.findings.filter((finding) => finding.rule === rule).map((f) => f.message);
}

/** A deck, and the same deck with one part's markup changed by a hand edit. */
function edited(
  part: string,
  replace: [string, string],
): {
  store: PartStore;
  baseline: PartStore;
} {
  const parts = minimalDeck();
  const original = deckBytes();
  const store = PartStore.open(original);
  const changed = parts[part]!.replace(replace[0], replace[1]);
  expect(changed, 'the fixture edit found nothing to replace').not.toBe(parts[part]);
  store.replacePart('/' + part, encoder.encode(changed));
  return { store, baseline: PartStore.open(original) };
}

describe('V027 parts nobody edited', () => {
  it('fires when a part changed and nothing says it was edited', () => {
    const parts = minimalDeck();
    const before = PartStore.open(deckBytes());
    const after = PartStore.open(
      deckBytes({
        parts: {
          'ppt/slides/slide1.xml': parts['ppt/slides/slide1.xml']!.replace(
            'name="Title 1"',
            'name="Title 1 "',
          ),
        },
      }),
    );

    const report = validatePackage({ store: after, baseline: before });
    const found = firedBy(report, 'V027');
    expect(found).toHaveLength(1);
    expect(found[0]).toContain('nobody edited this part and its bytes changed anyway');
    expect(report.ok).toBe(false);
  });

  it('says nothing about a part that was edited on purpose', () => {
    const { store, baseline } = edited('ppt/slides/slide1.xml', [
      'name="Title 1"',
      'name="Renamed by the user"',
    ]);
    expect(firedBy(validatePackage({ store, baseline }), 'V027')).toEqual([]);
  });

  it('fires on an OLE embedding that was replaced, edited or not', () => {
    const original = deckBytes({ parts: { 'ppt/embeddings/oleObject1.bin': 'CFB-ish bytes' } });
    const store = PartStore.open(original);
    store.replacePart('/ppt/embeddings/oleObject1.bin', encoder.encode('rewritten'));

    const found = firedBy(validatePackage({ store, baseline: PartStore.open(original) }), 'V027');
    expect(found).toHaveLength(1);
    expect(found[0]).toContain('OLE2/CFB compound files');
  });

  it('says nothing about a part that was added this session', () => {
    const original = deckBytes();
    const store = PartStore.open(original);
    store.addPart('/ppt/tags/tag1.xml', 'application/xml', encoder.encode('<t/>'));
    expect(firedBy(validatePackage({ store, baseline: PartStore.open(original) }), 'V027')).toEqual(
      [],
    );
  });
});

describe('V028 containers held opaque', () => {
  const EXT =
    '<a:extLst><a:ext uri="{FF2B5EF4-FFF2-40B4-BE49-F238E27FC236}">' +
    '<a16:creationId xmlns:a16="http://schemas.microsoft.com/office/drawing/2014/main" id="{1}"/>' +
    '</a:ext></a:extLst>';

  it('fires when an extension comes back different', () => {
    const parts = minimalDeck();
    const withExt = parts['ppt/slides/slide1.xml']!.replace(
      '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom>',
      '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom>' + EXT,
    );
    const original = deckBytes({ parts: { 'ppt/slides/slide1.xml': withExt } });
    const store = PartStore.open(original);
    // A "tidy-up" that rebuilt the extension from a typed model: same uri,
    // different bytes. This is the shape of the mistake, not a strawman.
    store.replacePart(
      '/ppt/slides/slide1.xml',
      encoder.encode(withExt.replace('id="{1}"', 'id="{2}"')),
    );

    const found = firedBy(validatePackage({ store, baseline: PartStore.open(original) }), 'V028');
    expect(found).toHaveLength(1);
    expect(found[0]).toContain('not markup this part arrived with');
    expect(found[0]).toContain('{FF2B5EF4-FFF2-40B4-BE49-F238E27FC236}');
  });

  it('says nothing when the extension is carried through and the shape moves', () => {
    const parts = minimalDeck();
    const withExt = parts['ppt/slides/slide1.xml']!.replace(
      '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom>',
      '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom>' + EXT,
    );
    const original = deckBytes({ parts: { 'ppt/slides/slide1.xml': withExt } });
    const store = PartStore.open(original);
    // An ordinary edit: the transform changes, the extension does not. The
    // element's position in the tree changes too, which is why this rule is a
    // multiset test rather than a comparison by path.
    store.replacePart(
      '/ppt/slides/slide1.xml',
      encoder.encode(withExt.replace('<a:off x="838200" y="365125"/>', '<a:off x="0" y="0"/>')),
    );

    expect(firedBy(validatePackage({ store, baseline: PartStore.open(original) }), 'V028')).toEqual(
      [],
    );
  });

  it('fires on an mc:AlternateContent branch that was rewritten', () => {
    const parts = minimalDeck();
    const MC =
      '<mc:AlternateContent xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006">' +
      '<mc:Choice xmlns:a14="http://schemas.microsoft.com/office/drawing/2010/main" Requires="a14">' +
      '<p:extLst/></mc:Choice><mc:Fallback><p:extLst/></mc:Fallback></mc:AlternateContent>';
    const withMc = parts['ppt/slides/slide1.xml']!.replace('</p:spTree>', MC + '</p:spTree>');
    const original = deckBytes({ parts: { 'ppt/slides/slide1.xml': withMc } });
    const store = PartStore.open(original);
    // The specific disaster: the prefix `a14` is renamed, which `Requires`
    // resolves by prefix rather than by URI, so ignorable markup becomes a hard
    // error.
    store.replacePart('/ppt/slides/slide1.xml', encoder.encode(withMc.split('a14').join('a99')));

    const found = firedBy(validatePackage({ store, baseline: PartStore.open(original) }), 'V028');
    expect(found.join('\n')).toContain('<mc:AlternateContent>');
  });
});

describe('V029 text and field identity', () => {
  const FIELD =
    '<a:fld id="{B7C4D8E1-0000-4000-8000-000000000001}" type="datetime1">' +
    '<a:t>28/08/2026</a:t></a:fld>';

  it('fires when xml:space="preserve" is dropped from an a:t that had it', () => {
    const parts = minimalDeck();
    const withSpace = parts['ppt/slides/slide1.xml']!.replace(
      '<a:p><a:endParaRPr lang="en-US"/></a:p>',
      '<a:p><a:r><a:rPr lang="en-US"/><a:t xml:space="preserve"> lead</a:t></a:r></a:p>',
    );
    const original = deckBytes({ parts: { 'ppt/slides/slide1.xml': withSpace } });
    const store = PartStore.open(original);
    store.replacePart(
      '/ppt/slides/slide1.xml',
      encoder.encode(withSpace.replace(' xml:space="preserve"', '')),
    );

    const found = firedBy(validatePackage({ store, baseline: PartStore.open(original) }), 'V029');
    expect(found).toHaveLength(1);
    expect(found[0]).toContain('no longer has it');
    expect(found[0]).toContain('120 elements in our corpus');
  });

  it('does not demand the attribute on text that never had it', () => {
    // The measured half: 120 `<a:t>` in the corpus carry edge whitespace with
    // no `xml:space`, and PowerPoint round-trips every one. A rule that
    // required the attribute would fire on real files that work.
    const parts = minimalDeck();
    const bare = parts['ppt/slides/slide1.xml']!.replace(
      '<a:p><a:endParaRPr lang="en-US"/></a:p>',
      '<a:p><a:r><a:rPr lang="en-US"/><a:t> lead</a:t></a:r></a:p>',
    );
    const original = deckBytes({ parts: { 'ppt/slides/slide1.xml': bare } });
    const store = PartStore.open(original);
    store.replacePart('/ppt/slides/slide1.xml', encoder.encode(bare.replace('lead', 'trail')));

    expect(firedBy(validatePackage({ store, baseline: PartStore.open(original) }), 'V029')).toEqual(
      [],
    );
  });

  it('fires on a regenerated a:fld GUID', () => {
    const parts = minimalDeck();
    const withField = parts['ppt/slides/slide1.xml']!.replace(
      '<a:p><a:endParaRPr lang="en-US"/></a:p>',
      '<a:p>' + FIELD + '</a:p>',
    );
    const original = deckBytes({ parts: { 'ppt/slides/slide1.xml': withField } });
    const store = PartStore.open(original);
    store.replacePart(
      '/ppt/slides/slide1.xml',
      encoder.encode(withField.replace('000000000001', '000000000002')),
    );

    const found = firedBy(validatePackage({ store, baseline: PartStore.open(original) }), 'V029');
    expect(found.join('\n')).toContain('is not one this part arrived with');
    expect(found.join('\n')).toContain('required ST_Guid');
  });

  it('fires when a field loses its cached text', () => {
    const parts = minimalDeck();
    const withField = parts['ppt/slides/slide1.xml']!.replace(
      '<a:p><a:endParaRPr lang="en-US"/></a:p>',
      '<a:p>' + FIELD + '</a:p>',
    );
    const original = deckBytes({ parts: { 'ppt/slides/slide1.xml': withField } });
    const store = PartStore.open(original);
    store.replacePart(
      '/ppt/slides/slide1.xml',
      encoder.encode(withField.replace('<a:t>28/08/2026</a:t>', '')),
    );

    const found = firedBy(validatePackage({ store, baseline: PartStore.open(original) }), 'V029');
    expect(found.join('\n')).toContain('lost its cached text ("28/08/2026")');
  });
});

describe('the preservation rules on a deck nobody touched', () => {
  it('find nothing, which is the only thing that makes them useful', () => {
    const { store } = deck();
    const report = validatePackage({ store, baseline: store });
    expect(report.findings).toEqual([]);
    // V003 is the one rule that reads the archive rather than the store, and
    // no bytes were passed here. It is named in `skipped`, not counted as a pass.
    expect(report.skipped.map((entry) => entry.rule)).toEqual(['V003']);
    expect(report.checked).toHaveLength(28);
  });
});
