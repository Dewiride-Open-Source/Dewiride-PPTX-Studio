import { PartStore, readZip } from '@pptx-studio/opc';
import { isValidateError } from '@pptx-studio/validate';
import { describe, expect, it } from 'vitest';
import { exportPackage, openPackage } from './export.js';
import { isWriterError } from './errors.js';
import type { PrepareHook } from './prepare.js';
import { assertPreserved } from './preserve.js';
import { fixture, fixtureBytes, MAIN_PART } from './testing/package.js';

const NOT_A_PRESENTATION = { validate: false } as const;

function entryNames(bytes: Uint8Array): string[] {
  return readZip(bytes).entries.map((entry) => entry.name);
}

describe('a no-op export', () => {
  it('serialises nothing and streams everything', () => {
    // The claim the package exists to make, stated as two numbers. `rewritten`
    // empty is what "we did not re-serialise anything" looks like from inside;
    // the preservation check below is what it looks like from outside.
    const pkg = fixture({ media: ['/ppt/media/image1.png', '/ppt/media/image2.png'] });
    const result = exportPackage({ ...pkg, ...NOT_A_PRESENTATION });

    expect(result.rewritten).toEqual([]);
    expect(result.streamed).toBe(pkg.store.partNames.length);
    expect(result.preservation.skipped).toBeNull();
    expect(result.preservation.checked).toBeGreaterThan(0);
    expect(result.collection.collect).toEqual([]);
    expect(result.prepared).toEqual([]);
  });

  it('produces the same entries, in the same order', () => {
    const pkg = fixture({ media: ['/ppt/media/image1.png'] });
    const result = exportPackage({ ...pkg, ...NOT_A_PRESENTATION });

    expect(entryNames(result.bytes)).toEqual(entryNames(pkg.baselineBytes));
  });

  it('checks every entry, including the content-type stream', () => {
    const pkg = fixture({ media: ['/ppt/media/image1.png'] });
    const result = exportPackage({ ...pkg, ...NOT_A_PRESENTATION });

    // Every entry in the archive, `[Content_Types].xml` included - it is passed
    // through untouched when nothing altered the map.
    expect(result.preservation.checked).toBe(entryNames(pkg.baselineBytes).length);
    expect(result.preservation.rewritten).toBe(0);
  });
});

describe('an export with one edit in it', () => {
  it('names the part it re-serialised and nothing else', () => {
    const pkg = fixture({ media: ['/ppt/media/image1.png'] });
    pkg.store.replacePart(MAIN_PART, new TextEncoder().encode('<doc edited="1"/>'));

    const result = exportPackage({ ...pkg, ...NOT_A_PRESENTATION });
    expect(result.rewritten).toEqual([MAIN_PART]);
    expect(result.streamed).toBe(pkg.store.partNames.length - 1);
    expect(result.preservation.skipped).toBeNull();
  });

  it('counts a relationship collection edited in place', () => {
    // `PartInfo.fromArchive` cannot see this one: the parsed collection *is* the
    // edit, `replacePart` was never called, and `write` re-emits it anyway. If
    // the writer reported only `fromArchive`, this part would be listed as
    // streamed and the preservation check would then compare it and fail - so
    // the two halves have to agree, and `PartStore.rewrittenParts` is where
    // they are put together.
    const pkg = fixture({ media: ['/ppt/media/image1.png'] });
    pkg.store.relationships(MAIN_PART).addTo('http://example.invalid/rel', MAIN_PART);

    const result = exportPackage({ ...pkg, ...NOT_A_PRESENTATION });
    expect(result.rewritten).toEqual(['/_rels/doc.xml.rels']);
    expect(result.preservation.skipped).toBeNull();
  });

  it('re-emits the content-type stream only when the map changed', () => {
    const pkg = fixture();
    pkg.store.addPart('/extra.bin', 'application/octet-stream', new Uint8Array([1, 2, 3]));
    pkg.store.relationships(MAIN_PART).addTo('http://example.invalid/rel', '/extra.bin');

    const result = exportPackage({ ...pkg, ...NOT_A_PRESENTATION });
    expect(pkg.store.contentTypes.dirty).toBe(true);
    expect(result.preservation.skipped).toBeNull();
  });
});

describe('the preservation check', () => {
  it('catches a part that was rewritten without being declared', () => {
    // Simulating the failure the check exists for: bytes changed, nothing said
    // so. `V027` cannot catch this, because it compares two stores and both
    // would answer from the same archive - only the emitted bytes show it.
    const bytes = fixtureBytes({ media: ['/ppt/media/image1.png'] });
    const store = PartStore.open(bytes);
    store.replacePart(MAIN_PART, new TextEncoder().encode('<doc rewritten="1"/>'));
    const output = store.write();

    try {
      assertPreserved(output, bytes, new Set());
      expect.unreachable('should have refused');
    } catch (error) {
      expect(isWriterError(error) && error.code).toBe('ERR_PRESERVATION_BROKEN');
      expect((error as Error).message).toContain('doc.xml');
    }
  });

  it('reports itself as not run rather than as passed', () => {
    // "Skipped" and "found nothing" are opposite answers, and a check that
    // reports zero either way is worse than no check.
    const pkg = fixture({ media: ['/ppt/media/image1.png'] });
    const noSource = exportPackage({
      store: pkg.store,
      baseline: pkg.baseline,
      ...NOT_A_PRESENTATION,
    });
    expect(noSource.preservation.skipped).toContain('not supplied');
    expect(noSource.preservation.checked).toBe(0);

    const turnedOff = exportPackage({
      ...fixture({ media: ['/ppt/media/image1.png'] }),
      ...NOT_A_PRESENTATION,
      verifyPreservation: false,
    });
    expect(turnedOff.preservation.skipped).toContain('turned off');
  });
});

describe('prepare hooks', () => {
  function recorder(name: string, log: string[], body?: (store: PartStore) => void): PrepareHook {
    return {
      name,
      run(context) {
        log.push(name);
        body?.(context.store);
        context.note(name + ' ran');
      },
    };
  }

  it('runs in the order given, and reports what each had to say', () => {
    const log: string[] = [];
    const pkg = fixture();
    const result = exportPackage({
      ...pkg,
      ...NOT_A_PRESENTATION,
      prepare: [recorder('second-in-name-only', log), recorder('first-in-name-only', log)],
    });

    expect(log).toEqual(['second-in-name-only', 'first-in-name-only']);
    expect(result.prepared.map((record) => record.hook)).toEqual([
      'second-in-name-only',
      'first-in-name-only',
    ]);
    expect(result.prepared[0]!.notes).toEqual(['second-in-name-only ran']);
  });

  it('leaves a silent hook out of the report', () => {
    const result = exportPackage({
      ...fixture(),
      ...NOT_A_PRESENTATION,
      prepare: [{ name: 'quiet', run: () => undefined }],
    });
    expect(result.prepared).toEqual([]);
  });

  it('runs before collection, so a hook can orphan something', () => {
    // The ordering that matters: 8.7 both adds font parts and drops the ones
    // nothing uses. A sweep that ran first would be answering about the wrong
    // package.
    const pkg = fixture({ media: ['/ppt/media/image1.png'] });
    const result = exportPackage({
      ...pkg,
      ...NOT_A_PRESENTATION,
      prepare: [
        {
          name: 'unlink-the-image',
          run(context) {
            const rels = context.store.relationships(MAIN_PART);
            rels.remove(rels.all[0]!.id);
          },
        },
      ],
    });

    expect(result.collection.collect).toEqual(['/ppt/media/image1.png']);
    expect(entryNames(result.bytes)).not.toContain('ppt/media/image1.png');
  });

  it('stops the export when a hook throws, and names it', () => {
    // Five of the six font artifacts written is a package PowerPoint reports a
    // problem with and no way to find out which.
    try {
      exportPackage({
        ...fixture(),
        ...NOT_A_PRESENTATION,
        prepare: [
          {
            name: 'embed-fonts',
            run() {
              throw new Error('no EOT writer yet');
            },
          },
        ],
      });
      expect.unreachable('should have refused');
    } catch (error) {
      expect(isWriterError(error) && error.code).toBe('ERR_PREPARE_FAILED');
      expect(isWriterError(error) && error.detail.hook).toBe('embed-fonts');
      expect((error as Error).message).toContain('no EOT writer yet');
    }
  });
});

describe('the firewall', () => {
  it('hands back a file whose defects it did not cause', () => {
    // Worth having as a test of the writer and not only of the rules, because
    // it is the behaviour that decides whether this is usable software. The
    // fixture is a well-formed OPC package and not a presentation, so several
    // rules fire on it - and the export succeeds anyway, because every one of
    // them fires on the baseline too. Refusing here would mean a deck with one
    // pre-existing defect could be opened and never saved again: the editor
    // declining to give somebody back their own file over a problem it did not
    // cause.
    const result = exportPackage(fixture());

    expect(result.report?.ok).toBe(true);
    expect(result.report?.blocking).toBe(0);
    const fatal = result.report!.findings.filter((finding) => finding.severity === 'fatal');
    expect(fatal.length).toBeGreaterThan(0);
    expect(fatal.every((finding) => finding.origin === 'inherited')).toBe(true);
  });

  it('refuses the same package when there is no baseline to blame', () => {
    // No history means no way to tell an inherited defect from one we caused,
    // and the conservative reading is the safe one: everything counts as ours.
    // Which is another way of saying that leaving the baseline out is not a
    // shortcut - it makes the firewall stricter, not quieter.
    const { store } = fixture();
    try {
      exportPackage({ store });
      expect.unreachable('should have refused');
    } catch (error) {
      expect(isValidateError(error) && error.code).toBe('ERR_VALIDATION_FAILED');
    }
  });

  it('is on unless the caller turns it off', () => {
    expect(exportPackage({ ...fixture(), ...NOT_A_PRESENTATION }).report).toBeNull();
  });
});

describe('openPackage', () => {
  it('hands back two stores that do not share state', () => {
    const pkg = openPackage(fixtureBytes({ media: ['/ppt/media/image1.png'] }));
    pkg.store.removePart('/ppt/media/image1.png');

    expect(pkg.store.has('/ppt/media/image1.png')).toBe(false);
    expect(pkg.baseline.has('/ppt/media/image1.png')).toBe(true);
  });
});
