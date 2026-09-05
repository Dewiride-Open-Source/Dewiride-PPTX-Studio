import { CONTENT_TYPE, PartStore, REL_TYPE } from '@pptx-studio/opc';
import { describe, expect, it } from 'vitest';
import { orphanedParts, reachableParts } from './reachability.js';
import { fixture, fixtureBytes, MAIN_PART } from '../testing/package.js';

describe('the walk', () => {
  it('reaches the main part and everything it relates to', () => {
    const { store } = fixture({ media: ['/ppt/media/image1.png', '/ppt/media/image2.png'] });
    const walk = reachableParts(store);

    expect(walk.complete).toBe(true);
    expect(orphanedParts(store, walk)).toEqual([]);
    expect(walk.reachable.has('/doc.xml')).toBe(true);
    expect(walk.reachable.has('/ppt/media/image1.png')).toBe(true);
  });

  it('counts a relationship part as reachable when its source is', () => {
    // Nothing ever relates to a `.rels` - the spec forbids it and PartStore
    // refuses a package that tries - so a walk that followed only edges would
    // report every relationship part in the package as orphaned. True of the
    // graph, false of the container, and it would make the sweep delete the
    // relationships it just used to decide.
    const { store } = fixture({ media: ['/ppt/media/image1.png'] });
    const walk = reachableParts(store);

    expect(store.partNames).toContain('/_rels/.rels');
    expect(store.partNames).toContain('/_rels/doc.xml.rels');
    expect(orphanedParts(store, walk)).toEqual([]);
  });

  it('finds a media part nothing points at', () => {
    const { store } = fixture({
      media: ['/ppt/media/image1.png'],
      orphanMedia: ['/ppt/media/image9.png'],
    });
    const walk = reachableParts(store);

    expect(walk.complete).toBe(true);
    expect(orphanedParts(store, walk)).toEqual(['/ppt/media/image9.png']);
  });

  it('stops being complete when a relationship part will not parse', () => {
    // The whole contract. An unreadable `.rels` hides edges, hidden edges make
    // their targets look unreferenced, and unreferenced is what a sweep deletes
    // on. `@pptx-studio/census` swallows this and carries on, which is right for
    // a report and would be a data-loss bug here.
    const { store } = fixture({ media: ['/ppt/media/image1.png'] });
    store.replacePart('/_rels/doc.xml.rels', new TextEncoder().encode('<Relationships'));

    const walk = reachableParts(store);
    expect(walk.complete).toBe(false);
    expect(walk.blockedBy).toEqual(['/_rels/doc.xml.rels']);
  });

  it('refuses to reason about a package with no root relationships', () => {
    // Every part would look orphaned, which is the one wrong answer that has
    // consequences. `PartStore.write` refuses such a package a moment later; we
    // decline to act first rather than racing it.
    const store = PartStore.open(fixtureBytes());
    store.rootRelationships().remove(store.rootRelationships().all[0]!.id);

    const walk = reachableParts(store);
    expect(walk.complete).toBe(false);
    expect(walk.reachable.size).toBe(0);
  });

  it('follows a chain rather than only the root', () => {
    const bytes = fixtureBytes();
    const store = PartStore.open(bytes);
    store.addPart('/ppt/chart.xml', CONTENT_TYPE.chart, new TextEncoder().encode('<c/>'));
    store.relationships(MAIN_PART).addTo(REL_TYPE.chart, '/ppt/chart.xml');
    store.addPart('/ppt/media/deep.png', CONTENT_TYPE.png, new Uint8Array([1]));
    store.relationships('/ppt/chart.xml').addTo(REL_TYPE.image, '/ppt/media/deep.png');

    const walk = reachableParts(store);
    expect(walk.complete).toBe(true);
    expect(orphanedParts(store, walk)).toEqual([]);
  });

  it('ignores external targets without treating them as a failure', () => {
    const { store } = fixture();
    store.relationships(MAIN_PART).add(REL_TYPE.hyperlink, 'https://example.invalid/', 'External');

    const walk = reachableParts(store);
    expect(walk.complete).toBe(true);
    expect(walk.blockedBy).toEqual([]);
  });
});
