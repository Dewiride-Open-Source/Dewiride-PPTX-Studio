import { storage, stream, writeCfb } from '../cfb.ts';
import { grid, prstGeom, scheme, shape, solidFill, type Cell } from '../shapes.ts';
import { textLine, txBody } from '../text.ts';
import type { ProbeDeck } from '../types.ts';
import { vbaProjectNodes } from '../vba.ts';

/**
 * A macro-enabled package, and the one deck in the corpus that is not a
 * `.pptx`.
 *
 * ## Two independent things, and the census only sees one of them
 *
 * "Macro-enabled" is **a content type**, and nothing else. Measured on
 * 2026-08-28: PowerPoint 16.0.20326, asked to save a one-slide deck as `.pptm`,
 * wrote a package byte-for-byte the shape of the `.pptx` one but for a single
 * `Override`:
 *
 * ```xml
 * <Override PartName="/ppt/presentation.xml"
 *   ContentType="application/vnd.ms-powerpoint.presentation.macroEnabled.main+xml"/>
 * ```
 *
 * There was **no `ppt/vbaProject.bin`**, because there were no macros. No
 * attribute, no element and no relationship anywhere in the markup says a deck
 * holds macros; the only two signals are that content type and the presence of
 * the part, and they are independent. A `.pptm` with no code is ordinary, and
 * `packages/census`'s `macros` rule keys on the **part** - a name ending
 * `/vbaproject.bin` - so a deck probing it has to carry a real one.
 *
 * ## Why the extension has to change
 *
 * PowerPoint checks the file extension against the content type of the main
 * part and refuses the pair when they disagree. `ProbeDeck.extension` exists
 * for this deck and is expected to stay used by only this deck: `.ppsx`,
 * `.potx` and the rest differ from `.pptx` in the same one content type and
 * would be probes for the same fact.
 *
 * That reaches further than it looks. `outputName()` feeds the writer, the
 * `--check` reader, the manifest's `path`, the manifest's `format` and
 * `C008-orphan`; `.gitattributes` already carries `*.pptm binary`, so
 * `C016-gitattributes` was satisfied before this deck existed. The hazard is
 * that `format` and `path` are two claims about one file and nothing
 * cross-checks them - so both come from `deck.extension` and cannot disagree.
 *
 * ## What is inside the part
 *
 * A real compound file, written by `tools/corpus/gen/cfb.ts`, holding a real
 * MS-OVBA project written by `tools/corpus/gen/vba.ts`:
 *
 * ```
 * /PROJECT            the project manifest, plain text
 * /PROJECTwm          the module name, MBCS then UTF-16
 * /VBA/_VBA_PROJECT   a version stamp
 * /VBA/dir            a compressed record stream
 * /VBA/Module1        Sub Noop() / End Sub, compressed
 * ```
 *
 * ## What is verified, and what is not
 *
 * This is the least-verified deck in Tier A and the module comments say so at
 * every layer rather than once.
 *
 * **Verified.** The package opens in PowerPoint with no repair prompt, with the
 * macro-enabled content type and the `.pptm` extension. The compound file reads
 * back through an independently written reader - `cfb.test.ts` follows the FAT
 * and MiniFAT chains and the directory's red-black tree rather than running the
 * writer backwards. The compressor round-trips its own output through a
 * decompressor written from the format. The `dir` record sequence walks to its
 * terminator and lands exactly on the end of the stream, which is what catches
 * the records whose four-byte field is a reserved constant rather than a size.
 *
 * **Not verified.** Whether the VBA engine would compile the project. Authoring
 * a module through COM needs "Trust access to the VBA project object model",
 * which is off by default and is a setting on the user's machine rather than
 * something this project may change - and PowerPoint does not load a VBA
 * project when it opens a file anyway, because macros are disabled by default
 * and the project is read when the editor opens or a macro runs. So a deck that
 * opens is evidence about the package, not about the project.
 *
 * Six `Reserved` constants inside `dir` and the `_VBA_PROJECT` version word are
 * asserted from the specification and checked by nothing. `ROSTER.md` carries
 * that as a declared gap; it closes with a Tier B deck on a machine where the
 * setting is on.
 *
 * ## No references, and no protection keys
 *
 * The project declares no `PROJECTREFERENCES` and omits `CMG`, `DPB` and `GC`
 * from `/PROJECT`. Both are the smallest honest choice. A reference to
 * `stdole` would be four more asserted constants and a library path that
 * differs by architecture; the three protection keys are the output of a seeded
 * obfuscation whose result differs on every save, so inventing values would
 * assert something specific about a protection state this project has not got.
 * Their absence means "unprotected", which is true.
 */

const VBA_CONTENT_TYPE = 'application/vnd.ms-office.vbaProject';
const VBA_RELATIONSHIP = 'http://schemas.microsoft.com/office/2006/relationships/vbaProject';

const SPEC = {
  projectName: 'VBAProject',
  moduleName: 'Module1',
  source: 'Sub Noop()\r\nEnd Sub\r\n',
} as const;

/** The whole `ppt/vbaProject.bin`: a compound file holding five streams. */
function vbaProjectBin(): Uint8Array {
  const nodes = vbaProjectNodes(SPEC);
  return writeCfb([
    stream('PROJECT', nodes.project),
    stream('PROJECTwm', nodes.projectWm),
    storage('VBA', [
      stream('_VBA_PROJECT', nodes.vbaProject),
      stream('dir', nodes.dir),
      stream(SPEC.moduleName, nodes.module),
    ]),
  ]);
}

const cell = grid(2, 2);

function box(id: number, name: string, c: Cell, accent: string, lines: readonly string[]): string {
  return shape({
    id,
    name,
    x: c.x,
    y: c.y,
    cx: c.cx,
    cy: c.cy,
    geometry: prstGeom('roundRect'),
    fill: solidFill(scheme(accent, '<a:lumMod val="60000"/><a:lumOff val="40000"/>')),
    textBody: txBody({
      bodyPr: '<a:bodyPr wrap="square" anchor="ctr"/>',
      paras: lines
        .map((line, i) => textLine(line, { sz: 1000, b: i === 0 }, { algn: 'ctr' }))
        .join(''),
    }),
  });
}

export const a32Macros: ProbeDeck = {
  id: 'a32-macros',
  title: 'PPTX Studio corpus: a32 macros',
  description:
    'The corpus&apos;s only .pptm, and the only deck whose file extension is not pptx - because ' +
    'PowerPoint checks the extension against the content type of the main part and refuses the pair ' +
    'when they disagree. Macro-enabled is a content type and nothing else: measured, a .pptm saved ' +
    'from PowerPoint with no code is identical to the .pptx but for one Override, and carries no ' +
    'vbaProject.bin at all. So the census&apos;s macros rule, which keys on a part named ' +
    'vbaProject.bin, needs a real one: a compound file holding a real MS-OVBA project whose only ' +
    'body is Sub Noop(). The container and the compression are verified by reading them back ' +
    'through independently written code; whether the VBA engine would compile the project is a ' +
    'declared gap, because that needs a Trust Center setting this project may not change.',
  extension: 'pptm',
  features: {
    shape: 18,
    placeholder: 6,
    presetGeom: 12,
    gradientFill: 2,
    macros: 1,
  },
  build: () => ({
    title: 'PPTX Studio corpus: a32 macros',
    macroEnabled: true,
    parts: [
      {
        name: 'ppt/vbaProject.bin',
        bytes: vbaProjectBin(),
        // A `Default` on `bin`, which is what PowerPoint writes: the extension
        // is shared with `ppt/embeddings/*.bin`, and an `Override` here would
        // be a second claim about the same extension.
        contentType: { kind: 'default', extension: 'bin', type: VBA_CONTENT_TYPE },
      },
    ],
    // On `ppt/presentation.xml`, and under Microsoft's 2006 extension base
    // rather than the ECMA one - which is why it is written out in full.
    presentationRels: [{ id: 'rId100', type: VBA_RELATIONSHIP, target: 'vbaProject.bin' }],
    slides: [
      {
        title: 'a32 — macro-enabled is a content type',
        body:
          box(10, 'The one difference', cell(0), 'accent1', [
            'application/vnd.ms-powerpoint',
            '.presentation.macroEnabled.main+xml',
            'on /ppt/presentation.xml',
          ]) +
          box(11, 'And nothing in the markup', cell(1), 'accent2', [
            'No attribute, no element, no flag.',
            'A deck that holds macros looks',
            'exactly like one that does not.',
          ]) +
          box(12, 'A .pptm with no code', cell(2), 'accent3', [
            'is ordinary. Measured: PowerPoint',
            'saving as .pptm wrote no vbaProject.bin,',
            'because there was nothing to write.',
          ]) +
          box(13, 'So the two are independent', cell(3), 'accent4', [
            'The content type says what the file is;',
            'the part says what is in it.',
            'The census keys on the part.',
          ]),
      },
      {
        title: 'a32 — and the extension has to match',
        body:
          box(10, 'Why this file is .pptm', cell(0), 'accent5', [
            'PowerPoint checks the extension against',
            'the content type of the main part',
            'and refuses the pair when they disagree.',
          ]) +
          box(11, 'What that reaches', cell(1), 'accent6', [
            'the writer, the --check reader,',
            'the manifest path, the manifest format,',
            'C008-orphan and C016-gitattributes',
          ]) +
          box(12, 'Two claims, one source', cell(2), 'accent1', [
            'Nothing cross-checks format against path,',
            'so both come from deck.extension',
            'and cannot disagree.',
          ]) +
          box(13, 'The only one', cell(3), 'accent2', [
            '.ppsx, .potx and the rest differ in',
            'the same one content type, so they',
            'would probe the same fact again.',
          ]),
      },
      {
        title: 'a32 — what is in ppt/vbaProject.bin',
        body:
          box(10, 'A compound file', cell(0), 'accent3', [
            'OLE2: a 512-byte header, sector chains,',
            'a directory of red-black trees, and a',
            'mini stream for anything under 4096 bytes.',
          ]) +
          box(11, 'Five streams', cell(1), 'accent4', [
            'PROJECT, PROJECTwm, and under /VBA:',
            '_VBA_PROJECT, dir, Module1.',
            'dir and Module1 are compressed.',
          ]) +
          box(12, 'Sub Noop()', cell(2), 'accent5', [
            'The whole of the code. Enough to be',
            'a project, small enough that every',
            'byte of the container is reviewable.',
          ]) +
          box(13, 'What is not checked', cell(3), 'accent6', [
            'Whether VBA would compile it.',
            'PowerPoint does not load the project',
            'on open, so opening proves nothing here.',
          ]),
      },
    ],
  }),
};
