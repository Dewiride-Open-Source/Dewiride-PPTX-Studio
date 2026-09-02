import { mkdtempSync, rmSync, writeFileSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { CONTENT_TYPE, PartStore, REL_TYPE, storedEntry, writeZip } from '@pptx-studio/opc';
import { afterAll, beforeAll, describe, expect, it } from 'vitest';
import { main, type Streams } from './main.js';

/**
 * The command line, run in-process.
 *
 * `main` takes its streams and returns an exit code instead of writing to
 * `process.stdout` and calling `process.exit`, which is what makes this a unit
 * test rather than a child-process fixture: no spawning, no captured pipes, no
 * shell quoting, and an assertion on the exit code that is just a number.
 */

const encoder = new TextEncoder();
const NS_P = 'http://schemas.openxmlformats.org/presentationml/2006/main';
const NS_A = 'http://schemas.openxmlformats.org/drawingml/2006/main';
const DECLARATION = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n';

let directory: string;

/** Two arrays and two closures over them. */
function streams(): { streams: Streams; out: () => string; err: () => string } {
  const outParts: string[] = [];
  const errParts: string[] = [];
  return {
    streams: {
      out: (text) => outParts.push(text),
      err: (text) => errParts.push(text),
    },
    out: () => outParts.join(''),
    err: () => errParts.join(''),
  };
}

function writeDeck(name: string, slideBody = ''): string {
  const store = PartStore.create();
  store.contentTypes.setDefault('rels', CONTENT_TYPE.relationships);
  store.contentTypes.setDefault('xml', CONTENT_TYPE.xml);
  store.addPart(
    '/ppt/presentation.xml',
    CONTENT_TYPE.presentation,
    encoder.encode(
      DECLARATION +
        `<p:presentation xmlns:p="${NS_P}" xmlns:a="${NS_A}">` +
        '<p:sldSz cx="12192000" cy="6858000" type="screen16x9"/>' +
        '<p:notesSz cx="6858000" cy="9144000"/>' +
        '</p:presentation>',
    ),
  );
  store.addPart(
    '/ppt/slides/slide1.xml',
    CONTENT_TYPE.slide,
    encoder.encode(
      DECLARATION +
        `<p:sld xmlns:p="${NS_P}" xmlns:a="${NS_A}"><p:cSld><p:spTree>` +
        slideBody +
        '</p:spTree></p:cSld></p:sld>',
    ),
  );
  store.rootRelationships().addTo(REL_TYPE.officeDocument, '/ppt/presentation.xml');
  store.relationships('/ppt/presentation.xml').addTo(REL_TYPE.slide, '/ppt/slides/slide1.xml');

  const path = join(directory, name);
  writeFileSync(path, store.write());
  return path;
}

beforeAll(() => {
  directory = mkdtempSync(join(tmpdir(), 'pptx-studio-cli-'));
});
afterAll(() => {
  rmSync(directory, { recursive: true, force: true });
});

describe('pptx-studio inspect', () => {
  it('reports a deck and exits 0', () => {
    const deck = writeDeck('good.pptx', '<p:sp><p:spPr><a:prstGeom prst="rect"/></p:spPr></p:sp>');
    const io = streams();

    expect(main(['inspect', deck], io.streams)).toBe(0);
    expect(io.out()).toContain('PACKAGE');
    expect(io.out()).toContain('1 slides');
    expect(io.out()).toContain('Preset geometry');
    expect(io.err()).toBe('');
  });

  it('emits JSON that parses back to the same census', () => {
    const deck = writeDeck('json.pptx');
    const io = streams();

    expect(main(['inspect', deck, '--json'], io.streams)).toBe(0);
    const census: unknown = JSON.parse(io.out());
    expect(census).toMatchObject({ censusVersion: 1 });
    expect((census as { parts: unknown[] }).parts.length).toBeGreaterThan(0);
  });

  it('writes to a file when asked, and nothing to stdout', () => {
    const deck = writeDeck('out.pptx');
    const target = join(directory, 'census.json');
    const io = streams();

    expect(main(['inspect', deck, '--json', '--out', target], io.streams)).toBe(0);
    expect(io.out()).toBe('');
    expect(main(['inspect', target], io.streams)).toBe(1); // a JSON file is not a package
  });

  it('exits 1 on a structural error and still prints the report', () => {
    // A slide that stops mid-element. A census does not refuse it - that is the
    // point of a census - but the exit code has to say something was wrong, or
    // the command is useless in a script.
    const store = PartStore.create();
    store.contentTypes.setDefault('rels', CONTENT_TYPE.relationships);
    store.contentTypes.setDefault('xml', CONTENT_TYPE.xml);
    store.addPart(
      '/ppt/presentation.xml',
      CONTENT_TYPE.presentation,
      encoder.encode(DECLARATION + `<p:presentation xmlns:p="${NS_P}"/>`),
    );
    store.addPart(
      '/ppt/slides/slide1.xml',
      CONTENT_TYPE.slide,
      encoder.encode(DECLARATION + `<p:sld xmlns:p="${NS_P}"><p:cSld>`),
    );
    store.rootRelationships().addTo(REL_TYPE.officeDocument, '/ppt/presentation.xml');
    store.relationships('/ppt/presentation.xml').addTo(REL_TYPE.slide, '/ppt/slides/slide1.xml');
    const path = join(directory, 'broken.pptx');
    writeFileSync(path, store.write());

    const io = streams();
    expect(main(['inspect', path], io.streams)).toBe(1);
    expect(io.out()).toContain('PROBLEMS');
    expect(io.out()).toContain('PART_NOT_WELL_FORMED');
  });

  it('reports the opc error code when the file is not a package at all', () => {
    const path = join(directory, 'not-a-zip.pptx');
    writeFileSync(path, 'this is not a zip archive');
    const io = streams();

    expect(main(['inspect', path], io.streams)).toBe(1);
    expect(io.err()).toContain('ERR_');
    expect(io.out()).toBe('');
  });

  it('does not choke on an archive that is a zip but not a presentation', () => {
    const path = join(directory, 'empty.pptx');
    writeFileSync(path, writeZip([storedEntry('hello.txt', encoder.encode('hi'))]));
    const io = streams();

    // No `[Content_Types].xml`, so `opc` refuses to open it and says why.
    expect(main(['inspect', path], io.streams)).toBe(1);
    expect(io.err()).toContain('pptx-studio: ERR_');
  });
});

describe('the command line itself', () => {
  it('prints usage and exits 2 when given nothing', () => {
    const io = streams();
    expect(main([], io.streams)).toBe(2);
    // Two verbs since sub-phase 1.2, so the usage line names neither and the
    // command list names both.
    expect(io.out()).toContain('Usage: pptx-studio <command> <deck.pptx>');
    expect(io.out()).toContain('  inspect  ');
    expect(io.out()).toContain('  validate ');
  });

  it('exits 0 for --help, because asking for help is not an error', () => {
    const io = streams();
    expect(main(['--help'], io.streams)).toBe(0);
    expect(io.out()).toContain('Usage:');
  });

  it('names the sub-phase that brings a verb it does not have yet', () => {
    // This said `bisect` until 1.5 built it, which is the way this test is
    // supposed to fail: a verb that arrives has to be taken off the list, and
    // the list is what the failure points at.
    const io = streams();
    expect(main(['render', 'deck.pptx'], io.streams)).toBe(2);
    expect(io.err()).toContain('sub-phase 3.10');
    expect(io.err()).not.toContain('unknown command');
  });

  it('says unknown for a verb that is not planned either', () => {
    const io = streams();
    expect(main(['frobnicate'], io.streams)).toBe(2);
    expect(io.err()).toContain('unknown command: frobnicate');
  });

  it('rejects a --top that is not a positive integer', () => {
    const io = streams();
    expect(main(['inspect', 'deck.pptx', '--top', 'lots'], io.streams)).toBe(2);
    expect(io.err()).toContain('--top wants a positive integer');
  });

  it('prints a version', () => {
    const io = streams();
    expect(main(['--version'], io.streams)).toBe(0);
    expect(io.out().trim()).toMatch(/^\d+\.\d+\.\d+/);
  });
});
