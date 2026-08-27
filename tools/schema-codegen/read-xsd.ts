/**
 * A deliberately small reader for the ECMA-376 Transitional XSDs.
 *
 * ## Why this does not use `@pptx-studio/xml`
 *
 * It would be pleasing symmetry, and it is the wrong dependency. This tool
 * *generates source for* `packages/xml`, so importing that package's build
 * output makes the codegen depend on the artifact it feeds: build xml, run the
 * codegen, regenerate xml's source, build xml again. Importing its `src`
 * instead does not work either - the package's internal imports carry `.js`
 * extensions for the bundler, which Node's type stripper does not remap - and
 * pointing `tools/` at `dist` breaks `pnpm typecheck` on a clean clone, because
 * typecheck runs before build.
 *
 * So this reads the schemas directly, and it is safe to do so for one reason:
 * **it refuses everything it does not expect.** The input is two dozen
 * machine-generated files whose SHA-256 is pinned in the generated output, and
 * the profile below is the whole of what they use. A CDATA section, an entity
 * reference, a DOCTYPE, a namespace declaration below the root, an unbalanced
 * tag or an unknown schema construct is a hard failure with an offset, not a
 * silent misread.
 *
 * The generated table is then checked against every parent/child pair in the
 * corpus, so a misreading here does not survive to the ADR either.
 *
 * ## The profile
 *
 * Measured across the 26 Transitional schemas: `xsd:sequence`, `xsd:choice`,
 * `xsd:group` (definition and `ref`), `xsd:element` (local with `name`+`type`,
 * or `ref` to a global), and `xsd:any`. No `xsd:all`, no `complexContent`, no
 * `simpleContent`, no `substitutionGroup`, no `abstract`, no type derivation of
 * any kind. That absence is what makes a hundred-line reader sufficient rather
 * than reckless.
 */

const XSD_NS = 'http://www.w3.org/2001/XMLSchema';

export class XsdReadError extends Error {
  readonly file: string;
  readonly offset: number;

  constructor(file: string, offset: number, message: string) {
    super(`${file}@${String(offset)}: ${message}`);
    this.name = 'XsdReadError';
    this.file = file;
    this.offset = offset;
  }
}

/** A named element in a content model, with the type it carries at that site. */
export interface XsdElementParticle {
  readonly kind: 'element';
  /** Namespace URI. `elementFormDefault="qualified"` throughout, so this is the target namespace. */
  readonly ns: string;
  readonly local: string;
  /** Namespace URI of the type name, and the type's local name. */
  readonly typeNs: string;
  readonly typeName: string;
  readonly min: number;
  readonly max: number;
}

export interface XsdGroupRef {
  readonly kind: 'group';
  readonly ns: string;
  readonly name: string;
  readonly min: number;
  readonly max: number;
}

export interface XsdCompositor {
  readonly kind: 'sequence' | 'choice';
  readonly min: number;
  readonly max: number;
  readonly children: readonly XsdParticle[];
}

/** `xsd:any`. Three sites in the whole of PresentationML and DrawingML, and all three matter. */
export interface XsdAny {
  readonly kind: 'any';
  readonly min: number;
  readonly max: number;
}

export type XsdParticle = XsdElementParticle | XsdGroupRef | XsdCompositor | XsdAny;

export interface XsdSchema {
  readonly file: string;
  readonly targetNamespace: string;
  /** Complex types by local name. `undefined` content means attribute-only. */
  readonly complexTypes: ReadonlyMap<string, XsdParticle | undefined>;
  readonly groups: ReadonlyMap<string, XsdParticle>;
  /** Global element declarations, by local name, to `{typeNs, typeName}`. */
  readonly globalElements: ReadonlyMap<string, { readonly ns: string; readonly name: string }>;
}

// ------------------------------------------------------------------- scanning

interface RawTag {
  readonly qname: string;
  readonly attrs: ReadonlyMap<string, string>;
  readonly selfClosing: boolean;
  readonly closing: boolean;
  readonly at: number;
}

interface RawNode {
  readonly qname: string;
  readonly attrs: ReadonlyMap<string, string>;
  readonly children: RawNode[];
  readonly at: number;
}

const SPACE = /\s/;

function isNameChar(c: string): boolean {
  return /[A-Za-z0-9_.:-]/.test(c);
}

/**
 * Tags, in order, with nothing between them interpreted.
 *
 * Text content is skipped rather than collected: these schemas carry no
 * `xsd:documentation`, and nothing in the content model we extract lives in a
 * text node.
 */
function* scanTags(text: string, file: string): Generator<RawTag> {
  let i = 0;
  const fail = (at: number, message: string): never => {
    throw new XsdReadError(file, at, message);
  };

  while (i < text.length) {
    const lt = text.indexOf('<', i);
    if (lt < 0) return;
    i = lt;

    if (text.startsWith('<!--', i)) {
      const close = text.indexOf('-->', i + 4);
      if (close < 0) fail(i, 'unterminated comment');
      i = close + 3;
      continue;
    }
    if (text.startsWith('<![CDATA[', i)) fail(i, 'a CDATA section is outside the accepted profile');
    if (text.startsWith('<!', i)) fail(i, 'a DOCTYPE or other declaration is not accepted');
    if (text.startsWith('<?', i)) {
      const close = text.indexOf('?>', i + 2);
      if (close < 0) fail(i, 'unterminated processing instruction');
      i = close + 2;
      continue;
    }

    const at = i;
    i++;
    const closing = text.charAt(i) === '/';
    if (closing) i++;

    const nameStart = i;
    while (i < text.length && isNameChar(text.charAt(i))) i++;
    if (i === nameStart) fail(at, 'a tag with no name');
    const qname = text.slice(nameStart, i);

    const attrs = new Map<string, string>();
    let selfClosing = false;
    for (;;) {
      while (i < text.length && SPACE.test(text.charAt(i))) i++;
      if (i >= text.length) fail(at, 'unterminated tag');
      const c = text.charAt(i);
      if (c === '>') {
        i++;
        break;
      }
      if (c === '/') {
        if (text.charAt(i + 1) !== '>') fail(i, 'a "/" that does not close the tag');
        selfClosing = true;
        i += 2;
        break;
      }
      if (closing) fail(i, 'an end tag may not carry attributes');

      const attrStart = i;
      while (i < text.length && isNameChar(text.charAt(i))) i++;
      if (i === attrStart) fail(i, 'expected an attribute name');
      const name = text.slice(attrStart, i);
      while (i < text.length && SPACE.test(text.charAt(i))) i++;
      if (text.charAt(i) !== '=') fail(i, `attribute "${name}" has no value`);
      i++;
      while (i < text.length && SPACE.test(text.charAt(i))) i++;
      const quote = text.charAt(i);
      if (quote !== '"' && quote !== "'") fail(i, `attribute "${name}" is not quoted`);
      i++;
      const valueStart = i;
      const close = text.indexOf(quote, i);
      if (close < 0) fail(valueStart, `attribute "${name}" is not terminated`);
      const value = text.slice(valueStart, close);
      // Nothing in the profile needs reference expansion, and silently
      // mis-expanding one would corrupt a type name. Refuse instead.
      if (value.includes('&')) fail(valueStart, `attribute "${name}" contains a reference`);
      if (attrs.has(name)) fail(attrStart, `attribute "${name}" appears twice`);
      attrs.set(name, value);
      i = close + 1;
    }

    yield { qname, attrs, selfClosing, closing, at };
  }
}

/** Build the element tree, checking that tags nest. */
function buildTree(text: string, file: string): RawNode {
  const fail = (at: number, message: string): never => {
    throw new XsdReadError(file, at, message);
  };
  const roots: RawNode[] = [];
  const stack: RawNode[] = [];

  for (const tag of scanTags(text, file)) {
    if (tag.closing) {
      const open = stack.pop();
      if (open === undefined) fail(tag.at, `</${tag.qname}> closes nothing`);
      else if (open.qname !== tag.qname) {
        fail(tag.at, `</${tag.qname}> closes <${open.qname}>`);
      }
      continue;
    }
    const node: RawNode = { qname: tag.qname, attrs: tag.attrs, children: [], at: tag.at };
    const parent = stack[stack.length - 1];
    if (parent === undefined) roots.push(node);
    else parent.children.push(node);
    if (!tag.selfClosing) stack.push(node);
  }

  if (stack.length > 0) fail(text.length, `<${stack[stack.length - 1]!.qname}> is never closed`);
  if (roots.length !== 1) fail(0, `expected one root element, found ${String(roots.length)}`);
  return roots[0]!;
}

// ---------------------------------------------------------------- interpreting

function occurs(node: RawNode, name: 'minOccurs' | 'maxOccurs', file: string): number {
  const raw = node.attrs.get(name);
  if (raw === undefined) return 1;
  if (raw === 'unbounded') return Infinity;
  const n = Number(raw);
  if (!Number.isInteger(n) || n < 0) {
    throw new XsdReadError(file, node.at, `${name}="${raw}" is not a count`);
  }
  return n;
}

/**
 * Read one schema.
 *
 * Namespace prefixes are resolved from declarations on the root element only.
 * A declaration anywhere below it is refused rather than handled: none of these
 * files has one, and the alternative is a scope walk that would exist solely to
 * be untested.
 */
export function readXsd(text: string, file: string): XsdSchema {
  const root = buildTree(text, file);
  const fail = (node: RawNode, message: string): never => {
    throw new XsdReadError(file, node.at, message);
  };

  const prefixes = new Map<string, string>();
  let defaultNs = '';
  for (const [name, value] of root.attrs) {
    if (name === 'xmlns') defaultNs = value;
    else if (name.startsWith('xmlns:')) prefixes.set(name.slice(6), value);
  }
  const xsdPrefixes = [...prefixes].filter(([, uri]) => uri === XSD_NS).map(([p]) => p);
  if (defaultNs === XSD_NS) xsdPrefixes.push('');
  if (xsdPrefixes.length !== 1) {
    fail(root, `expected exactly one prefix bound to ${XSD_NS}`);
  }
  const xsd = xsdPrefixes[0] === '' ? '' : `${xsdPrefixes[0]!}:`;

  const targetNamespace = root.attrs.get('targetNamespace');
  if (targetNamespace === undefined) fail(root, 'no targetNamespace');
  if (root.qname !== `${xsd}schema`) fail(root, `root is <${root.qname}>, not <${xsd}schema>`);
  // Qualified local element declarations are what makes `ns` below simply the
  // target namespace. Every one of these schemas says so; assert rather than
  // assume, because the value silently changes what every generated key means.
  if (root.attrs.get('elementFormDefault') !== 'qualified') {
    fail(root, 'elementFormDefault is not "qualified"');
  }

  const resolve = (qname: string, node: RawNode): { ns: string; name: string } => {
    const colon = qname.indexOf(':');
    if (colon < 0) return { ns: defaultNs, name: qname };
    const uri = prefixes.get(qname.slice(0, colon));
    if (uri === undefined) fail(node, `prefix "${qname.slice(0, colon)}" is not bound`);
    return { ns: uri!, name: qname.slice(colon + 1) };
  };

  const assertNoDeclarations = (node: RawNode): void => {
    if (node === root) return;
    for (const name of node.attrs.keys()) {
      if (name === 'xmlns' || name.startsWith('xmlns:')) {
        fail(node, 'a namespace declaration below the root is outside the accepted profile');
      }
    }
  };

  const particle = (node: RawNode): XsdParticle => {
    assertNoDeclarations(node);
    const min = occurs(node, 'minOccurs', file);
    const max = occurs(node, 'maxOccurs', file);
    const tag = node.qname.startsWith(xsd) ? node.qname.slice(xsd.length) : node.qname;

    switch (tag) {
      case 'sequence':
      case 'choice':
        return { kind: tag, min, max, children: node.children.map(particle) };
      case 'any':
        return { kind: 'any', min, max };
      case 'group': {
        const ref = node.attrs.get('ref');
        if (ref === undefined) fail(node, 'a group inside a content model needs @ref');
        const { ns, name } = resolve(ref!, node);
        return { kind: 'group', ns, name, min, max };
      }
      case 'element': {
        const ref = node.attrs.get('ref');
        if (ref !== undefined) {
          // A reference to a global element declaration. Resolved in a second
          // pass, once every schema has been read, because it may cross files.
          const { ns, name } = resolve(ref, node);
          return { kind: 'element', ns, local: name, typeNs: '', typeName: '', min, max };
        }
        const local = node.attrs.get('name');
        const type = node.attrs.get('type');
        if (local === undefined || type === undefined) {
          fail(node, 'a local element declaration needs both @name and @type');
        }
        const resolved = resolve(type!, node);
        return {
          kind: 'element',
          ns: targetNamespace!,
          local: local!,
          typeNs: resolved.ns,
          typeName: resolved.name,
          min,
          max,
        };
      }
      default:
        return fail(node, `<${node.qname}> is outside the accepted profile`);
    }
  };

  /** The first compositor child of a complexType or group, ignoring attributes. */
  const contentOf = (node: RawNode): XsdParticle | undefined => {
    for (const child of node.children) {
      const tag = child.qname.startsWith(xsd) ? child.qname.slice(xsd.length) : child.qname;
      if (tag === 'attribute' || tag === 'attributeGroup' || tag === 'annotation') continue;
      return particle(child);
    }
    return undefined;
  };

  const complexTypes = new Map<string, XsdParticle | undefined>();
  const groups = new Map<string, XsdParticle>();
  const globalElements = new Map<string, { ns: string; name: string }>();

  for (const child of root.children) {
    assertNoDeclarations(child);
    const tag = child.qname.startsWith(xsd) ? child.qname.slice(xsd.length) : child.qname;
    switch (tag) {
      case 'complexType': {
        const name = child.attrs.get('name');
        if (name === undefined) fail(child, 'a top-level complexType needs @name');
        complexTypes.set(name!, contentOf(child));
        break;
      }
      case 'group': {
        const name = child.attrs.get('name');
        const content = contentOf(child);
        if (name === undefined || content === undefined)
          fail(child, 'a group needs @name and content');
        groups.set(name!, content!);
        break;
      }
      case 'element': {
        const name = child.attrs.get('name');
        const type = child.attrs.get('type');
        if (name === undefined || type === undefined) {
          fail(child, 'a global element needs @name and @type');
        }
        globalElements.set(name!, resolve(type!, child));
        break;
      }
      case 'simpleType':
      case 'attributeGroup':
      case 'import':
      case 'include':
      case 'annotation':
        break;
      default:
        fail(child, `top-level <${child.qname}> is outside the accepted profile`);
    }
  }

  return { file, targetNamespace: targetNamespace!, complexTypes, groups, globalElements };
}
