import { OpcError } from './errors.js';

/**
 * OPC part names - the grammar from ECMA-376 Part 2 / ISO 29500-2 §9.1.1.
 *
 *   part_name = 1*( "/" segment )
 *   segment   = 1*( pchar )      ; pchar per RFC 3986 §3.3
 *
 * The rule identifiers below (`M1.1` ... `M1.12`) are the spec's own, kept
 * verbatim so a violation can be looked up rather than argued about.
 *
 * The finding worth carrying forward: **this grammar makes zip-slip
 * structurally impossible.** `M1.9` forbids a segment ending in a dot and
 * `M1.10` requires at least one non-dot character in every segment, so `..` and
 * `.` are not merely discouraged - a conformant part name cannot contain them
 * at all. Path traversal is therefore not a special case bolted onto the
 * reader; it is the grammar refusing to parse. `X1.4` exists only so the
 * failure reports itself as traversal rather than as "segment ends with a dot",
 * which is the message someone reading a bug report actually needs.
 *
 * `..` still has to be handled in exactly one place - `resolveRelativeTarget`,
 * because relationship targets in `.rels` parts are ordinary relative URI
 * references and legitimately use it. That resolution is clamped at the package
 * root.
 */

/** OPC's own rule identifiers, plus `X1.*` for hostility the spec never considered. */
export type PartNameRule =
  | 'M1.1'
  | 'M1.3'
  | 'M1.4'
  | 'M1.5'
  | 'M1.6'
  | 'M1.7'
  | 'M1.8'
  | 'M1.9'
  | 'M1.10'
  | 'M1.11'
  | 'M1.12'
  /** Control character or NUL, literal or percent-encoded. */
  | 'X1.1'
  /** Backslash - a Windows separator smuggled through a ZIP entry name. */
  | 'X1.2'
  /** A leading drive letter, e.g. `/C:/...`, which is valid pchar and still hostile. */
  | 'X1.3'
  /** A `.` or `..` segment: path traversal. */
  | 'X1.4'
  /** Longer than the configured ceiling. */
  | 'X1.5';

/**
 * `fatal` means the name is rejected. `warning` means it is out of spec but
 * harmless, and rejecting it would fail a real deck over a hygiene nit - a
 * literal space in a media part name is the common case.
 */
export type PartNameSeverity = 'fatal' | 'warning';

export interface PartNameViolation {
  readonly rule: PartNameRule;
  readonly severity: PartNameSeverity;
  readonly message: string;
}

declare const partNameBrand: unique symbol;

/**
 * A string that has been through `toPartName`. There is no cast-free way to
 * make one, so a function taking a `PartName` can rely on the grammar rather
 * than re-validating.
 */
export type PartName = string & { readonly [partNameBrand]: true };

/** Default ceiling on part-name length. ZIP itself allows 65535. */
export const MAX_PART_NAME_LENGTH = 1024;

/** RFC 3986 unreserved. */
const UNRESERVED = /^[A-Za-z0-9\-._~]$/;

/** RFC 3986 pchar, minus pct-encoded, which is handled separately. */
const PCHAR_LITERAL = /^[A-Za-z0-9\-._~!$&'()*+,;=:@]$/;

const HEX = /^[0-9A-Fa-f]{2}$/;

const ALL_DOTS = /^\.+$/;

const DRIVE_LETTER = /^[A-Za-z]:$/;

function isControl(code: number): boolean {
  return code < 0x20 || code === 0x7f;
}

function checkSegment(segment: string, push: (v: PartNameViolation) => void): void {
  if (segment.length === 0) {
    push({
      rule: 'M1.3',
      severity: 'fatal',
      message: 'a part name shall not have empty segments',
    });
    return;
  }

  // `.` and `..` violate M1.9 and M1.10 too, but traversal is the useful name
  // for what is happening, so report it once and skip the grammar rules.
  if (ALL_DOTS.test(segment)) {
    push({
      rule: 'X1.4',
      severity: 'fatal',
      message:
        'segment ' +
        JSON.stringify(segment) +
        ' is a relative path reference. OPC part names are absolute and cannot contain ' +
        '"." or ".." segments (M1.9, M1.10).',
    });
    return;
  }

  if (segment.endsWith('.')) {
    push({
      rule: 'M1.9',
      severity: 'fatal',
      message: 'segment ' + JSON.stringify(segment) + ' shall not end with a dot',
    });
  }

  for (let i = 0; i < segment.length; i++) {
    const ch = segment[i]!;

    if (ch === '%') {
      const hex = segment.slice(i + 1, i + 3);
      if (!HEX.test(hex)) {
        push({
          rule: 'M1.6',
          severity: 'fatal',
          message: 'malformed percent-encoding at ' + JSON.stringify(segment.slice(i, i + 3)),
        });
        return;
      }
      const decoded = Number.parseInt(hex, 16);
      if (decoded === 0x2f || decoded === 0x5c) {
        push({
          rule: 'M1.7',
          severity: 'fatal',
          message:
            'segment contains a percent-encoded ' +
            (decoded === 0x2f ? 'forward slash' : 'backslash') +
            ' (%' +
            hex +
            '), which would smuggle a separator past the segment split',
        });
      } else if (isControl(decoded)) {
        push({
          rule: 'X1.1',
          severity: 'fatal',
          message: 'segment contains a percent-encoded control character (%' + hex + ')',
        });
      } else if (UNRESERVED.test(String.fromCharCode(decoded))) {
        push({
          rule: 'M1.8',
          severity: 'warning',
          message:
            'segment percent-encodes the unreserved character ' +
            JSON.stringify(String.fromCharCode(decoded)) +
            ' (%' +
            hex +
            '); OPC requires unreserved characters to appear literally',
        });
      }
      i += 2;
      continue;
    }

    const code = ch.charCodeAt(0);
    if (isControl(code)) {
      push({
        rule: 'X1.1',
        severity: 'fatal',
        message: 'segment contains the control character U+' + code.toString(16).padStart(4, '0'),
      });
      return;
    }
    if (ch === '\\') {
      push({
        rule: 'X1.2',
        severity: 'fatal',
        message:
          'segment contains a backslash. ZIP entry names use "/" only; a backslash is a ' +
          'Windows separator that only some readers honour, which is what makes it useful ' +
          'for smuggling a path.',
      });
      return;
    }
    if (!PCHAR_LITERAL.test(ch)) {
      push({
        rule: 'M1.6',
        severity: 'warning',
        message:
          'segment contains ' +
          JSON.stringify(ch) +
          ', which is not a pchar. It should be percent-encoded.',
      });
    }
  }
}

/**
 * Every rule `name` breaks, in reporting order. An empty array means the name
 * is a conformant OPC part name.
 */
export function validatePartName(
  name: string,
  maxLength: number = MAX_PART_NAME_LENGTH,
): PartNameViolation[] {
  const violations: PartNameViolation[] = [];
  const push = (v: PartNameViolation): void => {
    violations.push(v);
  };

  if (name.length === 0 || name === '/') {
    push({ rule: 'M1.1', severity: 'fatal', message: 'a part name shall not be empty' });
    return violations;
  }
  if (name.length > maxLength) {
    push({
      rule: 'X1.5',
      severity: 'fatal',
      message:
        'part name is ' +
        String(name.length) +
        ' characters, over the ' +
        String(maxLength) +
        '-character ceiling',
    });
    return violations;
  }
  if (!name.startsWith('/')) {
    push({
      rule: 'M1.4',
      severity: 'fatal',
      message: 'a part name shall start with a forward slash',
    });
    return violations;
  }
  if (name.endsWith('/')) {
    push({
      rule: 'M1.5',
      severity: 'fatal',
      message:
        'a part name shall not end with a forward slash. OPC packages have no directory ' +
        'entries; folders exist only as a consequence of part names.',
    });
    return violations;
  }

  const segments = name.slice(1).split('/');
  if (DRIVE_LETTER.test(segments[0] ?? '')) {
    push({
      rule: 'X1.3',
      severity: 'fatal',
      message:
        'part name begins with the drive letter ' +
        JSON.stringify(segments[0]) +
        '. A colon is a legal pchar, so the OPC grammar admits this; a filesystem does not.',
    });
  }
  for (const segment of segments) checkSegment(segment, push);

  return violations;
}

/** True when `name` breaks no rule at all, warnings included. */
export function isValidPartName(name: string): boolean {
  return validatePartName(name).length === 0;
}

function throwOnFatal(name: string, violations: readonly PartNameViolation[]): void {
  const fatal = violations.filter((v) => v.severity === 'fatal');
  if (fatal.length === 0) return;
  throw new OpcError(
    'ERR_INVALID_PART_NAME',
    JSON.stringify(name) +
      ' is not a usable OPC part name: ' +
      fatal.map((v) => '[' + v.rule + '] ' + v.message).join('; '),
    { entry: name },
  );
}

/**
 * Validate `name` as a part name, or throw. Warnings are tolerated - see
 * `PartNameSeverity` for why.
 */
export function toPartName(name: string, maxLength: number = MAX_PART_NAME_LENGTH): PartName {
  throwOnFatal(name, validatePartName(name, maxLength));
  return name as PartName;
}

/**
 * Convert a ZIP entry name to a part name and validate it.
 *
 * ZIP stores `ppt/slides/slide1.xml`; OPC names the same thing
 * `/ppt/slides/slide1.xml`. The two namespaces are not the same size, and the
 * content-type stream is the difference: `[Content_Types].xml` is a legal ZIP
 * entry name and is not a part at all - it has no content type of its own and
 * never appears in a relationship. `[` and `]` are not pchar, but the grammar
 * only rates that a warning (see `M1.6`), so the refusal is stated here rather
 * than left to a rule that was written to tolerate a space in a file name.
 * Reach for the content-type stream by ZIP entry name, via `CONTENT_TYPES_PART`.
 */
export function partNameFromZipEntry(
  entryName: string,
  maxLength: number = MAX_PART_NAME_LENGTH,
): PartName {
  if (isContentTypesStreamName(entryName)) {
    throw new OpcError(
      'ERR_INVALID_PART_NAME',
      'the content-type stream ' +
        JSON.stringify(entryName) +
        ' is not a part. It has no content type of its own and nothing relates to it, so it ' +
        'has no part name either.',
      { entry: entryName },
    );
  }
  return toPartName('/' + entryName, maxLength);
}

/** The ZIP entry name for a part: the part name without its leading slash. */
export function zipEntryNameFor(partName: PartName): string {
  return partName.slice(1);
}

/**
 * ASCII-only lowercase, for the `M1.12` equivalence rule.
 *
 * Deliberately not `toLowerCase()`. Part names are ASCII by grammar, but
 * Unicode case folding is not identity on every string a caller might hand us,
 * and a name that folds differently is a name that collides differently.
 */
export function normalizePartName(name: string): string {
  return name.replace(/[A-Z]/g, (c) => String.fromCharCode(c.charCodeAt(0) + 32));
}

/** A package-level naming collision, reported by `checkPartNameCollisions`. */
export interface PartNameCollision {
  readonly rule: 'M1.11' | 'M1.12';
  readonly names: readonly [string, string];
  readonly message: string;
}

/**
 * `M1.11` and `M1.12` across a whole package - neither can be judged from one
 * name alone.
 *
 * `M1.11` is the rule that stops `/ppt/media` and `/ppt/media/image1.png` from
 * both existing: a part name may not be another part name plus appended
 * segments, because a consumer that materialises the package on a filesystem
 * would need `ppt/media` to be a file and a directory at once.
 */
export function checkPartNameCollisions(names: readonly string[]): PartNameCollision[] {
  const found: PartNameCollision[] = [];
  const seen = new Map<string, string>();

  for (const name of names) {
    const key = normalizePartName(name);
    const previous = seen.get(key);
    if (previous !== undefined) {
      found.push({
        rule: 'M1.12',
        names: [previous, name],
        message:
          'part names ' +
          JSON.stringify(previous) +
          ' and ' +
          JSON.stringify(name) +
          ' are equivalent under the case-insensitive comparison OPC requires',
      });
      continue;
    }
    seen.set(key, name);
  }

  const keys = [...seen.keys()].sort();
  for (let i = 0; i < keys.length; i++) {
    const prefix = keys[i]! + '/';
    for (let j = i + 1; j < keys.length && keys[j]!.startsWith(prefix); j++) {
      found.push({
        rule: 'M1.11',
        names: [seen.get(keys[i]!)!, seen.get(keys[j]!)!],
        message:
          'part name ' +
          JSON.stringify(seen.get(keys[j]!)) +
          ' is ' +
          JSON.stringify(seen.get(keys[i]!)) +
          ' with segments appended; one name cannot be both a part and a folder',
      });
    }
  }

  return found;
}

/** The extension of a part name, without the dot, case preserved. Empty when there is none. */
export function partExtension(name: string): string {
  const base = name.slice(name.lastIndexOf('/') + 1);
  const dot = base.lastIndexOf('.');
  return dot <= 0 ? '' : base.slice(dot + 1);
}

/** Everything up to and including the last slash: `/ppt/slides/`. */
export function partDirectory(name: string): string {
  return name.slice(0, name.lastIndexOf('/') + 1);
}

/**
 * The relationship part that describes `partName`'s outgoing relationships.
 *
 * `/ppt/slides/slide1.xml` -> `/ppt/slides/_rels/slide1.xml.rels`, and the
 * package root `/` -> `/_rels/.rels`.
 */
export function relsPartNameFor(partName: string): PartName {
  if (partName === '/' || partName === '') return '/_rels/.rels' as PartName;
  const dir = partDirectory(partName);
  const base = partName.slice(dir.length);
  return toPartName(dir + '_rels/' + base + '.rels');
}

/** True for `/_rels/.rels` and `/<dir>/_rels/<name>.rels`. */
export function isRelationshipPartName(name: string): boolean {
  return /(^|\/)_rels\/[^/]*\.rels$/.test(name);
}

/**
 * The content-type stream is matched case-insensitively. It is not a part, so
 * `M1.12` does not formally apply, but every producer writes it identically and
 * no reader gains anything by being strict here.
 */
export function isContentTypesStreamName(entryName: string): boolean {
  return normalizePartName(entryName) === '[content_types].xml';
}

/**
 * Resolve a relationship `Target` against the part that declared it.
 *
 * Targets in a `.rels` part are ordinary relative URI references resolved
 * against the **source part's folder**, not the package root - which is why
 * `../slideLayouts/slideLayout1.xml` inside `ppt/slides/_rels/slide1.xml.rels`
 * means `/ppt/slideLayouts/slideLayout1.xml`. This is the one place `..` is
 * legitimate, and so the one place it has to be clamped.
 *
 * The caller checks `TargetMode="External"` first; this is for internal
 * targets only.
 */
export function resolveRelativeTarget(sourcePartName: string, target: string): PartName {
  const withoutFragment = target.split('#')[0] ?? '';
  if (withoutFragment === '') {
    throw new OpcError('ERR_INVALID_PART_NAME', 'relationship target is empty', {
      entry: sourcePartName,
    });
  }

  const base = withoutFragment.startsWith('/')
    ? []
    : partDirectory(sourcePartName).slice(1).split('/').filter(Boolean);

  const out = [...base];
  for (const segment of withoutFragment.replace(/^\//, '').split('/')) {
    if (segment === '' || segment === '.') continue;
    if (segment === '..') {
      if (out.length === 0) {
        throw new OpcError(
          'ERR_TARGET_ESCAPES_PACKAGE',
          'relationship target ' +
            JSON.stringify(target) +
            ' in ' +
            sourcePartName +
            ' resolves above the package root. There is nothing above the package root; a ' +
            'target that reaches for it is either broken or naming a file on the machine ' +
            'reading it.',
          { entry: sourcePartName },
        );
      }
      out.pop();
      continue;
    }
    out.push(segment);
  }

  return toPartName('/' + out.join('/'));
}
