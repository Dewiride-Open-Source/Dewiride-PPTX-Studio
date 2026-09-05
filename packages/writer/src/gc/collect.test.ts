import { CONTENT_TYPE, isOpcError, PartStore, REL_TYPE } from '@pptx-studio/opc';
import { describe, expect, it } from 'vitest';
import { collectGarbage, isMediaPart, planCollection } from './collect.js';
import { isWriterError } from './errors.js';
import { fixture, fixtureBytes, MAIN_PART } from './testing/package.js';

/** Drop the one relationship from `/doc.xml` to `target`. */
function unlink(store: PartStore, target: string): void {
  const rels = store.relationships(MAIN_PART);
  const rel = rels.all.find((r) => r.target.endsWith(target.split('/').pop()!));
  expect(rel, 'fixture should have a relationship to ' + target).toBeDefined();
  rels.remove(rel!.id);
}

describe('what the sweep collects', () => {
  it('collects an image this session unreferenced', () => {
    const { store, baseline } = fixture({
      media: ['/ppt/media/image1.png', '/ppt/media/image2.png'],
    });
    unlink(store, '/ppt/media/image2.png');

    const plan = planCollection({ store, baseline });
    expect(plan.safe).toBe(true);
    expect(plan.collect).toEqual(['/ppt/media/image2.png']);
    expect(store.has('/ppt/media/image2.png')).toBe(false);
    expect(store.has('/ppt/media/image1.png')).toBe(true);
  });

  it('keeps an image that arrived orphaned, and says why', () => {
    // The whole of the default policy. A deck that has been through several
    // editors often carries one of these, and collecting it would mean that
    // merely opening and saving a file changes it.
    const { store, baseline } = fixture({
      media: ['/ppt/media/image1.png'],
      orphanMedia: ['/ppt/media/stale.png'],
    });

    const plan = planCollection({ store, baseline });
    expect(plan.collect).toEqual([]);
    expect(plan.kept).toHaveLength(1);
    expect(plan.kept[0]!.part).toBe('/ppt/media/stale.png');
    expect(plan.kept[0]!.why).toContain('not ours to remove');
    expect(store.has('/ppt/media/stale.png')).toBe(true);
  });

  it('keeps an image the other user of it still needs', () => {
    // The bug that eats a shared logo: delete one of two pictures, collect the
    // image, and the surviving slide shows a red X. It cannot happen here,
    // because the surviving relationship still marks the image live - which is
    // a property of rooting the sweep in the relationship graph rather than a
    // case anyone had to remember.
    const bytes = fixtureBytes({ media: ['/ppt/media/shared.png'] });
    const store = PartStore.open(bytes);
    const baseline = PartStore.open(bytes);
    store.addPart('/ppt/chart.xml', CONTENT_TYPE.chart, new TextEncoder().encode('<c/>'));
    store.relationships(MAIN_PART).addTo(REL_TYPE.chart, '/ppt/chart.xml');
    store.relationships('/ppt/chart.xml').addTo(REL_TYPE.image, '/ppt/media/shared.png');
    unlink(store, '/ppt/media/shared.png');

    expect(planCollection({ store, baseline }).collect).toEqual([]);
    expect(store.has('/ppt/media/shared.png')).toBe(true);
  });

  it('never leaves a relationship dangling behind it', () => {
    // A part is collected only when nothing resolves to it, so there is no edge
    // left to break - and `PartStore.write` refuses a package where we broke
    // one, which makes this checkable rather than merely argued.
    const { store, baseline } = fixture({ media: ['/ppt/media/image1.png'] });
    unlink(store, '/ppt/media/image1.png');

    expect(planCollection({ store, baseline }).collect).toEqual(['/ppt/media/image1.png']);
    expect(() => store.write()).not.toThrow();
  });

  it('cascades to a part that only the collected part referenced', () => {
    // Media has no relationships of its own, so the fixed point is invisible on
    // the default sweep set. It matters for the sweep sets later sub-phases
    // pass - a collected chart owns a `.rels` naming its style parts.
    const bytes = fixtureBytes();
    const store = PartStore.open(bytes);
    const baseline = PartStore.open(bytes);
    store.addPart('/ppt/a.xml', CONTENT_TYPE.chart, new TextEncoder().encode('<a/>'));
    store.addPart('/ppt/b.xml', CONTENT_TYPE.chart, new TextEncoder().encode('<b/>'));
    store.relationships(MAIN_PART).addTo(REL_TYPE.chart, '/ppt/a.xml');
    store.relationships('/ppt/a.xml').addTo(REL_TYPE.chartStyle, '/ppt/b.xml');
    // Both are reachable in the baseline only if the baseline has them, so
    // sweep with the policy that does not need one.
    unlink(store, '/ppt/a.xml');

    const plan = planCollection({
      store,
      baseline,
      policy: 'every-orphan',
      sweepable: (name) => name.startsWith('/ppt/') && name.endsWith('.xml'),
    });
    expect([...plan.collect].sort()).toEqual(['/ppt/a.xml', '/ppt/b.xml']);
  });
});

describe('what the sweep will not touch', () => {
  it('leaves an unreferenced part outside the sweep set alone', () => {
    // PowerPoint keeps layouts no slide uses, and so do we: that is what gives
    // the layout picker anything to offer. The sweep set is media.
    const bytes = fixtureBytes();
    const store = PartStore.open(bytes);
    const baseline = PartStore.open(bytes);
    store.addPart(
      '/ppt/slideLayouts/slideLayout1.xml',
      CONTENT_TYPE.slideLayout,
      new TextEncoder().encode('<l/>'),
    );

    const plan = planCollection({ store, baseline, policy: 'every-orphan' });
    expect(plan.collect).toEqual([]);
    expect(store.has('/ppt/slideLayouts/slideLayout1.xml')).toBe(true);
  });

  it('collects nothing at all under the none policy', () => {
    const { store, baseline } = fixture({ orphanMedia: ['/ppt/media/stale.png'] });
    const plan = planCollection({ store, baseline, policy: 'none' });

    expect(plan).toEqual({ collect: [], kept: [], safe: true, blockedBy: [] });
    expect(store.has('/ppt/media/stale.png')).toBe(true);
  });

  it('collects nothing when there is no baseline, and explains itself', () => {
    // With no baseline there is no "we", so `orphaned-here` has no question to
    // answer. Reporting it as a kept part rather than silently doing nothing is
    // the difference between "there was nothing to collect" and "I could not
    // tell", which are opposite answers.
    const { store } = fixture({ orphanMedia: ['/ppt/media/stale.png'] });
    const plan = planCollection({ store });

    expect(plan.collect).toEqual([]);
    expect(plan.kept[0]!.why).toContain('no baseline');
    expect(store.has('/ppt/media/stale.png')).toBe(true);
  });

  it('recognises the default sweep set', () => {
    expect(isMediaPart('/ppt/media/image1.png')).toBe(true);
    expect(isMediaPart('/PPT/Media/image1.png')).toBe(true);
    expect(isMediaPart('/ppt/embeddings/oleObject1.bin')).toBe(false);
    expect(isMediaPart('/ppt/slides/slide1.xml')).toBe(false);
    expect(isMediaPart('/ppt/mediaX/image1.png')).toBe(false);
  });
});

describe('when the graph cannot be trusted', () => {
  it('collects nothing and reports what blocked it', () => {
    const { store, baseline } = fixture({
      media: ['/ppt/media/image1.png'],
      orphanMedia: ['/ppt/media/stale.png'],
    });
    store.replacePart('/_rels/doc.xml.rels', new TextEncoder().encode('<Relationships'));

    const plan = planCollection({ store, baseline, policy: 'every-orphan' });
    expect(plan.safe).toBe(false);
    expect(plan.collect).toEqual([]);
    expect(plan.blockedBy).toEqual(['/_rels/doc.xml.rels']);
    expect(store.has('/ppt/media/stale.png')).toBe(true);
  });

  it('makes collectGarbage refuse rather than return an empty plan', () => {
    // The two entry points differ on purpose. A caller who asked for a sweep
    // and got silence would conclude there was nothing to sweep; `exportPackage`
    // uses the planning form instead, because refusing to save a file over a
    // defect it did not cause would be worse than saving it uncollected.
    const { store, baseline } = fixture({ media: ['/ppt/media/image1.png'] });
    store.replacePart('/_rels/doc.xml.rels', new TextEncoder().encode('<Relationships'));

    try {
      collectGarbage({ store, baseline, policy: 'every-orphan' });
      expect.unreachable('should have refused');
    } catch (error) {
      expect(isWriterError(error) && error.code).toBe('ERR_COLLECTION_UNSAFE');
      expect(isOpcError(error)).toBe(false);
    }
  });
});
