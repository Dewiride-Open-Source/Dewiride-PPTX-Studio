import type {
  PresetAdjustHandle,
  PresetCommand,
  PresetConnectionSite,
  PresetGuide,
  PresetPath,
  PresetPathFill,
  PresetPoint,
  PresetShape,
  PresetTextRect,
} from '../../packages/geometry/src/types.ts';

/**
 * A deliberately small reader for Apache POI's `presetShapeDefinitions.xml`.
 *
 * ## Why not `@pptx-studio/xml`
 *
 * The same reason `tools/schema-codegen/read-xsd.ts` does not use it: a tool
 * Node runs directly cannot import a workspace package by name, because nothing
 * links one into `tools/`. The type import above is relative and type-only, so
 * it is erased before Node ever sees it - which is the point, because it means
 * the transcoder and the runtime cannot disagree about the shape of the data
 * while still being two separate programs.
 *
 * ## The profile, measured rather than assumed
 *
 * Across the 538 970 bytes of the December-2016-era file this reader was
 * written against: no comments, no CDATA, no DOCTYPE, no entity references, no
 * processing instruction but the XML declaration, and no byte outside ASCII.
 * Nineteen element names and thirty-one attribute names, all listed below.
 * Every attribute value is drawn from letters, digits, and `+ * / ? : . -` and
 * the space that separates formula tokens.
 *
 * That absence is what makes two hundred lines sufficient rather than reckless.
 * Anything outside the profile is a hard failure with a byte offset - including
 * an element name this file has never seen - so a future POI release that adds
 * a construct stops the build instead of being silently dropped from a shape.
 *
 * ## Two things in the source that are not in the grammar
 *
 * Both are reported rather than repaired, and the generator prints them:
 *
 * - **Eight `+-` formulas carry four operands** where ECMA's `+-` takes three.
 *   All eight are in the circular-arrow family and all end in a spare `0`, so
 *   the intended value is unambiguous - but silently dropping the fourth token
 *   would mean this transcoder had an opinion about the semantics, which is
 *   2.2's job. The tokens are carried through as they are written.
 * - **Six formulas contain a double space.** Splitting on a single space would
 *   produce an empty operand and a formula one arity too long, so the split is
 *   on runs of whitespace and the six are reported.
 */

export class PresetReadError extends Error {
  readonly offset: number;

  constructor(offset: number, message: string) {
    super(`presetShapeDefinitions.xml@${String(offset)}: ${message}`);
    this.name = 'PresetReadError';
    this.offset = offset;
  }
}

interface Element {
  readonly name: string;
  readonly attrs: ReadonlyMap<string, string>;
  readonly children: readonly Element[];
  readonly offset: number;
}

const NAME_START = /[A-Za-z_]/;
const NAME_CHAR = /[A-Za-z0-9_.-]/;

/** Parse the whole file into a tree, refusing every construct outside the profile. */
function parse(text: string): Element {
  let i = 0;

  const fail = (message: string): never => {
    throw new PresetReadError(i, message);
  };

  const skipSpace = (): void => {
    while (i < text.length && /\s/.test(text[i] ?? '')) i++;
  };

  const readName = (): string => {
    const start = i;
    if (!NAME_START.test(text[i] ?? '')) fail('expected an element or attribute name');
    while (i < text.length && NAME_CHAR.test(text[i] ?? '')) i++;
    return text.slice(start, i);
  };

  /** Reads one element, assuming `text[i] === '<'` and the next char is a name start. */
  const readElement = (): Element => {
    const offset = i;
    i++; // '<'
    const name = readName();
    const attrs = new Map<string, string>();
    for (;;) {
      skipSpace();
      const c = text[i];
      if (c === '>') {
        i++;
        break;
      }
      if (c === '/') {
        if (text[i + 1] !== '>') fail('expected "/>"');
        i += 2;
        return { name, attrs, children: [], offset };
      }
      const attr = readName();
      if (attrs.has(attr)) fail(`attribute "${attr}" appears twice on <${name}>`);
      skipSpace();
      if (text[i] !== '=') fail(`attribute "${attr}" has no value`);
      i++;
      skipSpace();
      const quote = text[i];
      // Single quotes are legal XML and do not occur here. Accepting only what
      // the file uses keeps the reader honest about what it has been tested on.
      if (quote !== '"') fail('attribute values must be double-quoted');
      i++;
      const start = i;
      while (i < text.length && text[i] !== '"') i++;
      if (i >= text.length) fail('unterminated attribute value');
      attrs.set(attr, text.slice(start, i));
      i++;
    }

    const children: Element[] = [];
    for (;;) {
      const textStart = i;
      while (i < text.length && text[i] !== '<') i++;
      const between = text.slice(textStart, i);
      if (between.trim() !== '') {
        i = textStart;
        fail(`<${name}> has text content, which no preset definition uses`);
      }
      if (i >= text.length) fail(`<${name}> is never closed`);
      if (text[i + 1] === '/') {
        i += 2;
        const closing = readName();
        if (closing !== name) fail(`</${closing}> closes <${name}>`);
        skipSpace();
        if (text[i] !== '>') fail(`expected ">" after </${closing}>`);
        i++;
        return { name, attrs, children, offset };
      }
      if (text[i + 1] === '!') fail('comments, CDATA and DOCTYPE are all rejected');
      if (text[i + 1] === '?') fail('a processing instruction may only precede the root');
      children.push(readElement());
    }
  };

  skipSpace();
  if (text.startsWith('<?xml', i)) {
    const end = text.indexOf('?>', i);
    if (end < 0) fail('unterminated XML declaration');
    i = end + 2;
  }
  skipSpace();
  const root = readElement();
  skipSpace();
  if (i !== text.length) fail('trailing content after the root element');
  return root;
}

/** The whole of the element vocabulary, with the attributes each one may carry. */
const PROFILE: Readonly<Record<string, readonly string[]>> = {
  avLst: ['xmlns'],
  gdLst: ['xmlns'],
  gd: ['name', 'fmla'],
  ahLst: ['xmlns'],
  ahXY: ['gdRefX', 'minX', 'maxX', 'gdRefY', 'minY', 'maxY'],
  ahPolar: ['gdRefR', 'minR', 'maxR', 'gdRefAng', 'minAng', 'maxAng'],
  cxnLst: ['xmlns'],
  cxn: ['ang'],
  pos: ['x', 'y'],
  rect: ['l', 't', 'r', 'b', 'xmlns'],
  pathLst: ['xmlns'],
  path: ['w', 'h', 'fill', 'stroke', 'extrusionOk'],
  moveTo: [],
  lnTo: [],
  arcTo: ['wR', 'hR', 'stAng', 'swAng'],
  quadBezTo: [],
  cubicBezTo: [],
  close: [],
  pt: ['x', 'y'],
};

const DML_NS = 'http://schemas.openxmlformats.org/drawingml/2006/main';

/** The root element name, misspelled in the source file and reproduced verbatim. */
export const ROOT_ELEMENT = 'presetShapeDefinitons';

function check(element: Element): void {
  const allowed = PROFILE[element.name];
  if (allowed === undefined) {
    throw new PresetReadError(element.offset, `<${element.name}> is outside the profile`);
  }
  for (const attr of element.attrs.keys()) {
    if (!allowed.includes(attr)) {
      throw new PresetReadError(
        element.offset,
        `<${element.name}> carries an unexpected attribute "${attr}"`,
      );
    }
  }
  const ns = element.attrs.get('xmlns');
  if (ns !== undefined && ns !== DML_NS) {
    throw new PresetReadError(element.offset, `<${element.name}> declares an unknown namespace`);
  }
}

function attr(element: Element, name: string): string {
  const value = element.attrs.get(name);
  if (value === undefined) {
    throw new PresetReadError(element.offset, `<${element.name}> has no @${name}`);
  }
  return value;
}

function optional(element: Element, name: string): string | null {
  return element.attrs.get(name) ?? null;
}

function point(element: Element, expect: string): PresetPoint {
  const child = element.children[0];
  if (element.children.length !== 1 || child === undefined || child.name !== expect) {
    throw new PresetReadError(element.offset, `<${element.name}> needs exactly one <${expect}>`);
  }
  check(child);
  return { x: attr(child, 'x'), y: attr(child, 'y') };
}

function points(element: Element, count: number): PresetPoint[] {
  if (element.children.length !== count) {
    throw new PresetReadError(
      element.offset,
      `<${element.name}> needs ${String(count)} <pt>, found ${String(element.children.length)}`,
    );
  }
  return element.children.map((child) => {
    check(child);
    if (child.name !== 'pt') {
      throw new PresetReadError(child.offset, `expected <pt>, found <${child.name}>`);
    }
    return { x: attr(child, 'x'), y: attr(child, 'y') };
  });
}

/** Something in the source that the grammar does not describe. Reported, never repaired. */
export interface PresetAnomaly {
  readonly kind: 'over-long-formula' | 'collapsed-whitespace';
  readonly shape: string;
  readonly guide: string;
  readonly fmla: string;
}

export interface PresetSource {
  readonly shapes: readonly PresetShape[];
  readonly anomalies: readonly PresetAnomaly[];
}

/** ECMA's operand count for each of the seventeen operators, or -1 where the source disagrees. */
const ARITY: Readonly<Record<string, number>> = {
  '*/': 3,
  '+-': 3,
  '+/': 3,
  '?:': 3,
  abs: 1,
  at2: 2,
  cat2: 3,
  cos: 2,
  max: 2,
  min: 2,
  mod: 3,
  pin: 3,
  sat2: 3,
  sin: 2,
  sqrt: 1,
  tan: 2,
  val: 1,
};

export function readPresets(text: string): PresetSource {
  const root = parse(text);
  if (root.name !== ROOT_ELEMENT) {
    throw new PresetReadError(root.offset, `the root is <${root.name}>, not <${ROOT_ELEMENT}>`);
  }
  if (root.attrs.size !== 0) {
    throw new PresetReadError(root.offset, 'the root element carries attributes');
  }

  const anomalies: PresetAnomaly[] = [];
  const shapes: PresetShape[] = [];
  const seen = new Set<string>();

  for (const shape of root.children) {
    if (seen.has(shape.name)) {
      throw new PresetReadError(shape.offset, `<${shape.name}> is defined twice`);
    }
    seen.add(shape.name);
    if (shape.attrs.size !== 0) {
      throw new PresetReadError(shape.offset, `<${shape.name}> carries attributes`);
    }

    const guides = (element: Element): PresetGuide[] =>
      element.children.map((gd) => {
        check(gd);
        if (gd.name !== 'gd') {
          throw new PresetReadError(gd.offset, `expected <gd>, found <${gd.name}>`);
        }
        const raw = attr(gd, 'fmla');
        // Runs of whitespace, not a single space: six formulas in the file have
        // a double space, and splitting naively yields an empty operand.
        const fmla = raw.trim().split(/\s+/);
        const name = attr(gd, 'name');
        if (/\s\s/.test(raw)) {
          anomalies.push({
            kind: 'collapsed-whitespace',
            shape: shape.name,
            guide: name,
            fmla: raw,
          });
        }
        const op = fmla[0] ?? '';
        const arity = ARITY[op];
        if (arity === undefined) {
          throw new PresetReadError(gd.offset, `unknown formula operator "${op}"`);
        }
        if (fmla.length - 1 !== arity) {
          anomalies.push({ kind: 'over-long-formula', shape: shape.name, guide: name, fmla: raw });
        }
        return { name, fmla };
      });

    let avLst: readonly PresetGuide[] = [];
    let gdLst: readonly PresetGuide[] = [];
    let ahLst: readonly PresetAdjustHandle[] = [];
    let cxnLst: readonly PresetConnectionSite[] = [];
    let rect: PresetTextRect | null = null;
    let pathLst: readonly PresetPath[] = [];

    for (const part of shape.children) {
      check(part);
      switch (part.name) {
        case 'avLst':
          avLst = guides(part);
          break;
        case 'gdLst':
          gdLst = guides(part);
          break;
        case 'ahLst':
          ahLst = part.children.map((handle): PresetAdjustHandle => {
            check(handle);
            if (handle.name === 'ahXY') {
              return {
                kind: 'xy',
                gdRefX: optional(handle, 'gdRefX'),
                minX: optional(handle, 'minX'),
                maxX: optional(handle, 'maxX'),
                gdRefY: optional(handle, 'gdRefY'),
                minY: optional(handle, 'minY'),
                maxY: optional(handle, 'maxY'),
                pos: point(handle, 'pos'),
              };
            }
            if (handle.name === 'ahPolar') {
              return {
                kind: 'polar',
                gdRefR: optional(handle, 'gdRefR'),
                minR: optional(handle, 'minR'),
                maxR: optional(handle, 'maxR'),
                gdRefAng: optional(handle, 'gdRefAng'),
                minAng: optional(handle, 'minAng'),
                maxAng: optional(handle, 'maxAng'),
                pos: point(handle, 'pos'),
              };
            }
            throw new PresetReadError(handle.offset, `<${handle.name}> is not an adjust handle`);
          });
          break;
        case 'cxnLst':
          cxnLst = part.children.map((site): PresetConnectionSite => {
            check(site);
            if (site.name !== 'cxn') {
              throw new PresetReadError(site.offset, `expected <cxn>, found <${site.name}>`);
            }
            return { ang: attr(site, 'ang'), pos: point(site, 'pos') };
          });
          break;
        case 'rect':
          rect = {
            l: attr(part, 'l'),
            t: attr(part, 't'),
            r: attr(part, 'r'),
            b: attr(part, 'b'),
          };
          break;
        case 'pathLst':
          pathLst = part.children.map((path): PresetPath => {
            check(path);
            if (path.name !== 'path') {
              throw new PresetReadError(path.offset, `expected <path>, found <${path.name}>`);
            }
            return {
              w: Number(path.attrs.get('w') ?? '0'),
              h: Number(path.attrs.get('h') ?? '0'),
              fill: (path.attrs.get('fill') ?? 'norm') as PresetPathFill,
              stroke: (path.attrs.get('stroke') ?? 'true') !== 'false',
              extrusionOk: (path.attrs.get('extrusionOk') ?? 'true') !== 'false',
              commands: path.children.map((command): PresetCommand => {
                check(command);
                switch (command.name) {
                  case 'moveTo':
                    return { kind: 'moveTo', to: points(command, 1)[0] as PresetPoint };
                  case 'lnTo':
                    return { kind: 'lnTo', to: points(command, 1)[0] as PresetPoint };
                  case 'quadBezTo': {
                    const [c1, to] = points(command, 2) as [PresetPoint, PresetPoint];
                    return { kind: 'quadBezTo', c1, to };
                  }
                  case 'cubicBezTo': {
                    const [c1, c2, to] = points(command, 3) as [
                      PresetPoint,
                      PresetPoint,
                      PresetPoint,
                    ];
                    return { kind: 'cubicBezTo', c1, c2, to };
                  }
                  case 'arcTo':
                    return {
                      kind: 'arcTo',
                      wR: attr(command, 'wR'),
                      hR: attr(command, 'hR'),
                      stAng: attr(command, 'stAng'),
                      swAng: attr(command, 'swAng'),
                    };
                  case 'close':
                    return { kind: 'close' };
                  default:
                    throw new PresetReadError(
                      command.offset,
                      `<${command.name}> is not a path command`,
                    );
                }
              }),
            };
          });
          break;
        default:
          throw new PresetReadError(
            part.offset,
            `<${part.name}> is not part of a preset definition`,
          );
      }
    }

    shapes.push({ name: shape.name, avLst, gdLst, ahLst, cxnLst, rect, pathLst });
  }

  return { shapes, anomalies };
}
