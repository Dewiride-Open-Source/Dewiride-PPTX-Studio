import type { PackageCensus } from './types.js';

/**
 * A census as text.
 *
 * Lives in this package rather than in the CLI so that the one thing both
 * consumers show a human - the report - has one implementation and one set of
 * tests. `pptx-studio inspect` prints what this returns; the browser explorer
 * renders the same object as HTML.
 *
 * No colour and no terminal width detection. Output that is stable byte-for-
 * byte is worth more here than output that is pretty: it can be committed, it
 * can be diffed between two runs, and it survives a pipe into a file.
 */

export interface FormatOptions {
  /** Include the full per-part table. Off by default: a real deck has hundreds. */
  readonly parts?: boolean | undefined;
  /** Include the namespace histogram. */
  readonly namespaces?: boolean | undefined;
  /** Rows to show in each histogram before truncating. */
  readonly top?: number | undefined;
}

const KIB = 1024;

/** Sizes a human reads. Deliberately binary units, and always labelled. */
export function humanBytes(bytes: number): string {
  if (bytes < KIB) return String(bytes) + ' B';
  if (bytes < KIB * KIB) return (bytes / KIB).toFixed(1) + ' KiB';
  if (bytes < KIB * KIB * KIB) return (bytes / (KIB * KIB)).toFixed(1) + ' MiB';
  return (bytes / (KIB * KIB * KIB)).toFixed(2) + ' GiB';
}

function ms(value: number): string {
  return value < 1000 ? value.toFixed(1) + ' ms' : (value / 1000).toFixed(2) + ' s';
}

function rule(title: string): string {
  return '\n' + title + '\n' + '-'.repeat(Math.max(title.length, 8));
}

function pad(text: string, width: number): string {
  return text.length >= width ? text : text + ' '.repeat(width - text.length);
}

function padStart(text: string, width: number): string {
  return text.length >= width ? text : ' '.repeat(width - text.length) + text;
}

export function formatCensus(census: PackageCensus, options: FormatOptions = {}): string {
  const top = options.top ?? 15;
  const out: string[] = [];
  const { archive, presentation, relationships, timings } = census;

  out.push('PACKAGE');
  out.push('=======');
  out.push(
    '  archive        ' + humanBytes(archive.bytes) + ' in ' + String(archive.entries) + ' entries',
  );
  out.push(
    '  parts          ' +
      String(archive.parts) +
      ' (' +
      humanBytes(archive.declaredInflatedBytes) +
      ' inflated, ' +
      ratio(archive.declaredInflatedBytes, archive.compressedBytes) +
      ' overall)',
  );
  out.push(
    '  compression    ' +
      String(archive.deflatedEntries) +
      ' deflated, ' +
      String(archive.storedEntries) +
      ' stored' +
      (archive.zip64 ? ', ZIP64' : '') +
      (archive.dataDescriptorEntries > 0
        ? ', ' + String(archive.dataDescriptorEntries) + ' with data descriptors'
        : ''),
  );
  if (archive.largestPart !== null) {
    out.push(
      '  largest part   ' +
        archive.largestPart.name +
        ' (' +
        humanBytes(archive.largestPart.count) +
        ')',
    );
  }
  if (archive.directoryEntries > 0) {
    out.push('  directories    ' + String(archive.directoryEntries) + ' (PowerPoint writes none)');
  }

  if (presentation !== null) {
    out.push(rule('PRESENTATION'));
    const size =
      presentation.slideWidthEmu === null || presentation.slideHeightEmu === null
        ? 'not declared'
        : emuInches(presentation.slideWidthEmu) +
          ' x ' +
          emuInches(presentation.slideHeightEmu) +
          ' in' +
          (presentation.slideSizeType === null ? '' : ' (' + presentation.slideSizeType + ')');
    out.push('  slide size     ' + size);
    out.push(
      '  sheets         ' +
        String(presentation.slides) +
        ' slides, ' +
        String(presentation.slideLayouts) +
        ' layouts, ' +
        String(presentation.slideMasters) +
        ' masters',
    );
    out.push(
      '  notes          ' +
        String(presentation.notesSlides) +
        ' notes slides, ' +
        String(presentation.notesMasters) +
        ' notes masters, ' +
        String(presentation.handoutMasters) +
        ' handout masters',
    );
    if (presentation.sections > 0 || presentation.customShows > 0) {
      out.push(
        '  structure      ' +
          String(presentation.sections) +
          ' sections, ' +
          String(presentation.customShows) +
          ' custom shows',
      );
    }
    if (presentation.embeddedFonts.length > 0 || presentation.embedTrueTypeFonts) {
      out.push(
        '  fonts          embedTrueTypeFonts=' +
          String(presentation.embedTrueTypeFonts) +
          ' saveSubsetFonts=' +
          String(presentation.saveSubsetFonts),
      );
      for (const font of presentation.embeddedFonts) {
        out.push(
          '                 ' +
            pad(font.typeface, 28) +
            font.slots.join(', ') +
            (font.charset === null ? '' : '  charset ' + String(font.charset)),
        );
      }
    }
  }

  out.push(rule('FEATURES'));
  if (census.features.length === 0) out.push('  (none detected)');
  for (const feature of census.features) {
    out.push(
      '  ' +
        padStart(String(feature.count), 7) +
        '  ' +
        pad(feature.label, 34) +
        phaseTag(feature.phase),
    );
  }

  out.push(rule('TEXT'));
  const text = census.text;
  out.push(
    '  ' +
      String(text.paragraphs) +
      ' paragraphs, ' +
      String(text.runs) +
      ' runs, ' +
      String(text.characters) +
      ' characters, ' +
      String(text.fields) +
      ' fields',
  );
  if (text.normAutofit > 0 || text.shapeAutofit > 0) {
    out.push(
      '  autofit: ' +
        String(text.normAutofit) +
        ' normAutofit, ' +
        String(text.shapeAutofit) +
        ' spAutoFit',
    );
  }
  if (census.typefaces.length > 0) {
    out.push('  typefaces: ' + census.typefaces.slice(0, top).map(nameAndCount).join(', '));
  }
  if (census.languages.length > 0) {
    out.push('  languages: ' + census.languages.slice(0, top).map(nameAndCount).join(', '));
  }

  if (census.presetGeometry.length > 0) {
    out.push(rule('GEOMETRY'));
    out.push('  ' + census.presetGeometry.slice(0, top).map(nameAndCount).join(', '));
    if (census.presetGeometry.length > top) {
      out.push('  ... and ' + String(census.presetGeometry.length - top) + ' more presets');
    }
  }

  out.push(rule('RELATIONSHIPS'));
  out.push(
    '  ' +
      String(relationships.total) +
      ' edges (' +
      String(relationships.external) +
      ' external), ' +
      String(relationships.dangling.length) +
      ' dangling, ' +
      String(relationships.unreachableParts.length) +
      ' unreachable parts',
  );
  for (const entry of relationships.byType.slice(0, top)) {
    out.push('  ' + padStart(String(entry.count), 7) + '  ' + entry.name);
  }

  if (options.namespaces === true) {
    out.push(rule('NAMESPACES'));
    for (const namespace of census.namespaces) {
      const flags = [
        namespace.standard ? 'ECMA-376' : namespace.extension ? 'extension' : 'unknown',
        namespace.required ? 'Requires' : '',
        namespace.ignorable ? 'Ignorable' : '',
      ]
        .filter((flag) => flag !== '')
        .join(' ');
      out.push(
        '  ' +
          padStart(String(namespace.count), 8) +
          '  ' +
          pad(namespace.prefixes.join('|') || '(default)', 10) +
          pad(flags, 22) +
          namespace.uri,
      );
    }
  }

  if (options.parts === true) {
    out.push(rule('PARTS'));
    out.push(
      '  ' +
        pad('part', 52) +
        padStart('bytes', 11) +
        padStart('zipped', 11) +
        padStart('elems', 9) +
        padStart('depth', 7) +
        padStart('rels', 6),
    );
    for (const part of census.parts) {
      out.push(
        '  ' +
          pad(part.name, 52) +
          padStart(String(part.bytes), 11) +
          padStart(String(part.compressedBytes), 11) +
          padStart(part.elements === null ? '-' : String(part.elements), 9) +
          padStart(part.maxDepth === null ? '-' : String(part.maxDepth), 7) +
          padStart(String(part.relationships), 6),
      );
    }
  }

  if (census.problems.length > 0) {
    out.push(rule('PROBLEMS'));
    for (const problem of census.problems) {
      out.push(
        '  ' +
          pad(problem.severity, 8) +
          pad(problem.code, 26) +
          problem.message +
          (problem.part === null ? '' : '\n           in ' + problem.part),
      );
    }
  }

  out.push(rule('TIMING'));
  out.push(
    '  open ' +
      ms(timings.openMs) +
      ', inflate ' +
      ms(timings.inflateMs) +
      ', scan ' +
      ms(timings.scanMs) +
      ', total ' +
      ms(timings.totalMs),
  );
  out.push(
    '  ' +
      humanBytes(timings.xmlBytesScanned) +
      ' of XML across ' +
      String(timings.partsScanned) +
      ' parts' +
      (timings.scanMs > 0
        ? ' = ' +
          (timings.xmlBytesScanned / KIB / KIB / (timings.scanMs / 1000)).toFixed(0) +
          ' MiB/s'
        : ''),
  );

  return out.join('\n') + '\n';
}

function nameAndCount(entry: { name: string; count: number }): string {
  return entry.name + ' (' + String(entry.count) + ')';
}

function phaseTag(phase: string): string {
  return phase === 'preserve' ? 'carried across, never rendered' : 'phase ' + phase;
}

function ratio(inflated: number, compressed: number): string {
  if (compressed === 0) return 'n/a';
  return (inflated / compressed).toFixed(1) + ':1';
}

const EMU_PER_INCH = 914400;

function emuInches(emu: number): string {
  return (emu / EMU_PER_INCH).toFixed(2);
}
