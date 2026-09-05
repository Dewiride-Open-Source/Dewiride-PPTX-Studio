/**
 * The node model: a tree over the source string, where every node remembers
 * where it came from.
 *
 * The rule that shapes everything: **a node whose `dirty` flag is false has not
 * been touched anywhere in its subtree, and re-emits by slicing
 * `source.slice(start, end)`.** Sub-phase 0.6's edits set `dirty` on the node
 * they change and on its ancestors, and on nothing else - which is why moving a
 * shape does not disturb the `p:timing` animation tree or the
 * `mc:AlternateContent` ink block sitting beside it in the same slide part.
 *
 * Offsets are flat numbers rather than a nested `raw: { start, end }` object.
 * The plan sketched the latter; an 8 MB part holds 262 429 elements, and giving
 * each one three span objects means three quarters of a million allocations to
 * express six integers. The information is identical.
 *
 * Namespaces are deliberately *not* resolved during the build. Prefixes are
 * kept exactly as written and bindings are read on demand from the tree,
 * because both of the shortcuts are wrong on real files:
 *
 *   - A package-wide prefix table is wrong. In our corpus `p14` is bound to
 *     `.../powerpoint/2010/main` in 46 parts and to
 *     `.../powerpoint/2007/7/12/main` in 8 others.
 *   - A per-part table built from the root element is wrong. 213 namespace
 *     declarations in the corpus sit on a non-root element, and a shipped
 *     Microsoft template contains
 *     `<p14:discardImageEditData xmlns="" xmlns:p14="...2007/7/12/main" val="0"/>`
 *     nested inside `<p:extLst>` - a mid-document prefix binding beside a
 *     default-namespace undeclaration.
 *
 * PowerPoint accepts both shapes, and accepts one prefix bound to two different
 * URIs in sibling subtrees of a single part - verified by building such files
 * and opening them. So scope resolution is a walk up the tree.
 */

import { XmlError } from '../errors.js';
import { decodeXmlSource, type XmlEncoding } from './source.js';
import {
  DEFAULT_TOKENIZER_LIMITS,
  XmlTokenizer,
  type XAttribute,
  type XmlElementToken,
  type XmlToken,
  type XmlTokenizerLimits,
} from './tokenizer.js';

export type { XAttribute };

interface NodeBase {
  /** Index of the node's first character in the source. */
  start: number;
  /** One past its last. */
  end: number;
  /**
   * True once this node or something beneath it has been edited.
   *
   * While it is false, `source.slice(start, end)` **is** this node, which is
   * the whole point of the package.
   */
  dirty: boolean;
  parent: XElement | undefined;
}

export interface XElement extends NodeBase {
  readonly type: 'element';
  qname: string;
  prefix: string;
  local: string;
  attributes: XAttribute[];
  children: XNode[];
  /**
   * `<x/>` rather than `<x></x>`.
   *
   * Both are legal and both survive PowerPoint's reader, so the distinction is
   * ours to preserve: 33 elements in the corpus - all in `docProps` - use the
   * long form, and PowerPoint's own writer collapses them on resave.
   */
  selfClosing: boolean;
  /** One past the `>` of the start tag. Children begin here. */
  openTagEnd: number;
  /** The `<` of the end tag. Equal to {@link NodeBase.end} when self-closing. */
  closeTagStart: number;
  /** One past the element name inside the start tag. */
  nameEnd: number;
  /**
   * Where the whitespace before `>` or `/>` begins.
   *
   * Not cosmetic: 70 822 of the 98 777 self-closing tags in the corpus are
   * written `<a:off x="0" y="0" />`, so a serializer that assumes `/>` follows
   * the last attribute directly loses a character on most of the corpus.
   */
  trailingSpaceStart: number;
  /**
   * True if this element carries any `xmlns` or `xmlns:*` attribute.
   *
   * Computed by the tokenizer, which has already read every attribute name.
   * {@link resolvePrefix} walks ancestors, and without this it also rescanned
   * each ancestor's whole attribute list: a 1.7 MB document 1000 deep with 200
   * attributes per element took 37 seconds in `undeclaredPrefixes`. Both
   * numbers are legal - depth 1001 is under `maxDepth`, 200 attributes under
   * `maxAttributes` - and nothing counted the product.
   */
  hasNamespaceDeclarations: boolean;
}

export interface XText extends NodeBase {
  readonly type: 'text';
  /** References expanded, line endings normalized. Not the source text. */
  value: string;
  /**
   * True if the run is `S` characters only - inter-element formatting.
   *
   * Derived from `value`, and therefore not `readonly`: sub-phase 0.6's
   * `setValue` maintains it. A derived field that an edit can leave stale is a
   * bug that shows up somewhere far away from the edit.
   */
  whitespaceOnly: boolean;
}

export interface XCData extends NodeBase {
  readonly type: 'cdata';
  value: string;
}

export interface XComment extends NodeBase {
  readonly type: 'comment';
  value: string;
}

export interface XProcessingInstruction extends NodeBase {
  readonly type: 'processingInstruction';
  target: string;
  data: string;
}

export interface XDeclaration extends NodeBase {
  readonly type: 'declaration';
  readonly version: string;
  readonly encoding: string | undefined;
  readonly standalone: string | undefined;
}

export type XNode = XElement | XText | XCData | XComment | XProcessingInstruction | XDeclaration;

/**
 * A parsed part.
 *
 * `children` covers the whole document - the prolog, the root, and anything
 * after it - so the source is reconstructible from the tree alone. A byte order
 * mark is the one thing that is not a node: it is markup-free, it may precede a
 * declaration, and modelling it as one would give it a node type that makes no
 * sense. It lives on {@link XDocument.bom} instead.
 */
export interface XDocument {
  readonly source: string;
  readonly encoding: XmlEncoding;
  readonly bom: boolean;
  readonly children: XNode[];
  readonly root: XElement;
  readonly declaration: XDeclaration | undefined;
}

/** Ceilings on the tree, on top of the tokenizer's lexical ones. */
export interface XmlParseLimits extends XmlTokenizerLimits {
  /**
   * Element nesting depth.
   *
   * The deepest part in our corpus reaches 23, inside a `p:timing` tree.
   * PowerPoint has no useful limit of its own - it opens a part nested 5000
   * deep - so this number protects us rather than matching it. The builder uses
   * an explicit stack and never recurses, so the limit bounds what a hostile
   * document can allocate; it is not standing between us and a stack overflow.
   */
  readonly maxDepth: number;
  /** Nodes in one document. */
  readonly maxNodes: number;
}

export const DEFAULT_PARSE_LIMITS: XmlParseLimits = {
  ...DEFAULT_TOKENIZER_LIMITS,
  maxDepth: 1024,
  maxNodes: 4_000_000,
};

function elementOf(token: XmlElementToken): XElement {
  const selfClosing = token.type === 'emptyElementTag';
  return {
    type: 'element',
    qname: token.qname,
    prefix: token.prefix,
    local: token.local,
    // Adopted, not copied: these are the objects the tokenizer built.
    attributes: token.attributes as XAttribute[],
    children: [],
    selfClosing,
    start: token.start,
    end: token.end,
    openTagEnd: token.end,
    closeTagStart: token.end,
    nameEnd: token.nameEnd,
    trailingSpaceStart: token.trailingSpaceStart,
    hasNamespaceDeclarations: token.hasNamespaceDeclarations,
    dirty: false,
    parent: undefined,
  };
}

function leafOf(token: XmlToken): XNode {
  const shared = { start: token.start, end: token.end, dirty: false, parent: undefined };
  switch (token.type) {
    case 'text':
      return { type: 'text', value: token.value, whitespaceOnly: token.whitespaceOnly, ...shared };
    case 'cdata':
      return { type: 'cdata', value: token.value, ...shared };
    case 'comment':
      return { type: 'comment', value: token.value, ...shared };
    case 'processingInstruction':
      return { type: 'processingInstruction', target: token.target, data: token.data, ...shared };
    case 'declaration':
      return {
        type: 'declaration',
        version: token.version,
        encoding: token.encoding,
        standalone: token.standalone,
        ...shared,
      };
    default:
      throw new XmlError('ERR_MALFORMED_XML', 'unexpected token', { offset: token.start });
  }
}

/**
 * Build a tree from a decoded source string.
 *
 * Iterative, with an explicit stack. Recursive descent is the obvious shape and
 * the wrong one: a part nested a few thousand elements deep is well-formed XML
 * that PowerPoint opens without complaint, and it would exhaust the JavaScript
 * stack with a `RangeError` - the one thing this package has promised never to
 * throw.
 */
export function parseXmlString(
  source: string,
  limits: XmlParseLimits = DEFAULT_PARSE_LIMITS,
): XDocument {
  const tokenizer = new XmlTokenizer(source, limits);
  const bom = source.charCodeAt(0) === 0xfeff;

  const top: XNode[] = [];
  const stack: XElement[] = [];
  let root: XElement | undefined;
  let declaration: XDeclaration | undefined;
  let nodes = 0;

  const append = (node: XNode): void => {
    if (++nodes > limits.maxNodes) {
      throw new XmlError(
        'ERR_LIMIT_EXCEEDED',
        'this part has more XML nodes than the limit allows',
        {
          limit: limits.maxNodes,
        },
      );
    }
    const open = stack[stack.length - 1];
    if (open === undefined) {
      top.push(node);
    } else {
      node.parent = open;
      open.children.push(node);
    }
  };

  for (let token = tokenizer.next(); token !== undefined; token = tokenizer.next()) {
    if (token.type === 'startTag' || token.type === 'emptyElementTag') {
      if (stack.length === 0 && root !== undefined) {
        throw new XmlError(
          'ERR_ROOT_ELEMENT',
          'a document may contain only one root element; <' + token.qname + '> is a second',
          { offset: token.start, name: token.qname },
        );
      }
      const element = elementOf(token);
      append(element);
      if (stack.length === 0) root = element;
      if (token.type === 'startTag') {
        if (stack.length >= limits.maxDepth) {
          throw new XmlError(
            'ERR_LIMIT_EXCEEDED',
            'elements are nested deeper than the limit allows',
            {
              offset: token.start,
              name: token.qname,
              limit: limits.maxDepth,
            },
          );
        }
        stack.push(element);
      }
      continue;
    }

    if (token.type === 'endTag') {
      const open = stack.pop();
      if (open === undefined) {
        throw new XmlError(
          'ERR_MISMATCHED_TAG',
          '</' + token.qname + '> closes an element that was never opened',
          { offset: token.start, name: token.qname },
        );
      }
      if (open.qname !== token.qname) {
        throw new XmlError(
          'ERR_MISMATCHED_TAG',
          '</' + token.qname + '> closes <' + open.qname + '>',
          {
            offset: token.start,
            name: token.qname,
          },
        );
      }
      open.closeTagStart = token.start;
      open.end = token.end;
      continue;
    }

    // Character data outside the root is forbidden, and a CDATA section is
    // character data - checking only `text` let `<r/><![CDATA[x]]>` through
    // while `<r/>x` was correctly refused.
    if (
      stack.length === 0 &&
      (token.type === 'cdata' || (token.type === 'text' && !token.whitespaceOnly))
    ) {
      throw new XmlError(
        'ERR_MALFORMED_XML',
        'character data is not allowed outside the root element',
        {
          offset: token.start,
        },
      );
    }

    const leaf = leafOf(token);
    if (leaf.type === 'declaration') declaration = leaf;
    append(leaf);
  }

  if (stack.length > 0) {
    const open = stack[stack.length - 1]!;
    throw new XmlError('ERR_UNCLOSED_TAG', '<' + open.qname + '> is never closed', {
      offset: open.start,
      name: open.qname,
    });
  }
  if (root === undefined) {
    throw new XmlError('ERR_ROOT_ELEMENT', 'this document has no root element', { offset: 0 });
  }

  return { source, encoding: 'utf-8', bom, children: top, root, declaration };
}

/** Decode a part and build its tree. The entry point callers actually want. */
export function parseXml(bytes: Uint8Array, limits?: XmlParseLimits): XDocument {
  const decoded = decodeXmlSource(bytes);
  return { ...parseXmlString(decoded.text, limits), encoding: decoded.encoding };
}

// ------------------------------------------------------------------ accessors

/**
 * The source a node came from.
 *
 * While `node.dirty` is false this *is* the node, byte for byte. Once it is
 * true the slice is stale; use `serializeNode`, which rebuilds instead.
 */
export function sourceOf(document: XDocument, node: XNode): string {
  return document.source.slice(node.start, node.end);
}

/** The start tag's source, which stays valid even when a child has been edited. */
export function startTagOf(document: XDocument, element: XElement): string {
  return document.source.slice(element.start, element.openTagEnd);
}

/** Child elements, in document order. */
export function childElements(element: XElement): XElement[] {
  return element.children.filter((child): child is XElement => child.type === 'element');
}

/** The first child element with this qualified name. */
export function firstChild(element: XElement, qname: string): XElement | undefined {
  for (const child of element.children) {
    if (child.type === 'element' && child.qname === qname) return child;
  }
  return undefined;
}

/** An attribute by qualified name. */
export function attribute(element: XElement, qname: string): XAttribute | undefined {
  for (const attr of element.attributes) {
    if (attr.qname === qname) return attr;
  }
  return undefined;
}

/** An attribute's normalized value. */
export function attributeValue(element: XElement, qname: string): string | undefined {
  return attribute(element, qname)?.value;
}

/** Every element in the subtree, in document order, starting with `element`. */
export function* descendantElements(element: XElement): IterableIterator<XElement> {
  const stack: XElement[] = [element];
  while (stack.length > 0) {
    const current = stack.pop()!;
    yield current;
    for (let i = current.children.length - 1; i >= 0; i--) {
      const child = current.children[i]!;
      if (child.type === 'element') stack.push(child);
    }
  }
}

/** Concatenated text of the subtree, CDATA included, in document order. */
export function textContent(element: XElement): string {
  let out = '';
  const stack: XNode[] = [];
  for (let i = element.children.length - 1; i >= 0; i--) stack.push(element.children[i]!);
  while (stack.length > 0) {
    const node = stack.pop()!;
    if (node.type === 'text' || node.type === 'cdata') out += node.value;
    else if (node.type === 'element') {
      for (let i = node.children.length - 1; i >= 0; i--) stack.push(node.children[i]!);
    }
  }
  return out;
}

// -------------------------------------------------------------------- editing

/**
 * Mark a node as edited, and its ancestors with it.
 *
 * This is the whole preservation guarantee in four lines. `dirty` means "the
 * source span no longer describes me", and it has to travel *up* - an element
 * whose child was rewritten can no longer be re-emitted as one slice either -
 * but it must never travel *down* or *sideways*. That asymmetry is what lets a
 * shape move without disturbing the `p:timing` animation tree or the
 * `mc:AlternateContent` ink block that shares the slide part with it.
 *
 * The walk stops at the first ancestor that is already dirty, which is what
 * keeps a 60-frame drag from being quadratic in tree depth: the second edit
 * inside a subtree pays only for the distance to the first one.
 *
 * `journal`, when given, collects exactly the nodes this call changed - not the
 * ones that were dirty already. An edit records that list so that its inverse
 * can put the flags back, which is what lets undo return a part to its original
 * bytes rather than to a rebuilt copy of them. See `edit.ts`.
 */
export function markDirty(node: XNode, journal?: XNode[]): void {
  if (!node.dirty) {
    node.dirty = true;
    journal?.push(node);
  }
  for (let parent = node.parent; parent !== undefined && !parent.dirty; parent = parent.parent) {
    parent.dirty = true;
    journal?.push(parent);
  }
}

/**
 * Mark one attribute as edited.
 *
 * Separate from {@link markDirty} because an attribute is not a node: it has no
 * `parent`, so it cannot propagate on its own. Setting `attribute.dirty` by hand
 * and forgetting the element is a silent no-op - the element would still be
 * clean, and the serializer would re-emit the whole start tag as a slice,
 * discarding the edit. {@link checkDirtyInvariant} exists to catch that.
 */
export function markAttributeDirty(
  element: XElement,
  attribute: XAttribute,
  journal?: { nodes: XNode[]; attributes: XAttribute[] },
): void {
  if (!attribute.dirty) {
    attribute.dirty = true;
    journal?.attributes.push(attribute);
  }
  markDirty(element, journal?.nodes);
}

/**
 * Places where `dirty` has been set without propagating, which the serializer
 * would silently ignore.
 *
 * Returns the same shape as {@link checkTreeCoverage} so a caller can check both
 * and concatenate. Both describe the same class of failure: a tree that looks
 * fine and serializes to the wrong bytes.
 */
export function checkDirtyInvariant(document: XDocument): CoverageGap[] {
  const gaps: CoverageGap[] = [];
  const stack: XNode[] = [...document.children];
  while (stack.length > 0) {
    const node = stack.pop()!;
    if (node.dirty && node.parent !== undefined && !node.parent.dirty) {
      gaps.push({
        kind: 'malformed',
        at: node.start,
        detail:
          'a dirty node sits under a clean <' +
          node.parent.qname +
          '>, which will re-emit as a slice and discard the edit',
      });
    }
    if (node.type !== 'element') continue;
    for (const attribute of node.attributes) {
      if (attribute.dirty && !node.dirty) {
        gaps.push({
          kind: 'malformed',
          at: attribute.start,
          detail:
            'attribute "' +
            attribute.qname +
            '" is dirty but <' +
            node.qname +
            '> is not; use markAttributeDirty',
        });
      }
    }
    for (const child of node.children) stack.push(child);
  }
  return gaps;
}

// ----------------------------------------------------------------- namespaces

/** The `xmlns` and `xmlns:*` bindings declared **on this element**. */
export function declaredNamespaces(element: XElement): Map<string, string> {
  const declared = new Map<string, string>();
  if (!element.hasNamespaceDeclarations) return declared;
  for (const attr of element.attributes) {
    if (attr.qname === 'xmlns') declared.set('', attr.value);
    else if (attr.prefix === 'xmlns') declared.set(attr.local, attr.value);
  }
  return declared;
}

/**
 * Resolve a prefix to a namespace URI by walking up the tree.
 *
 * Returns `''` for a prefix explicitly undeclared with `xmlns=""` - a real
 * construct that ships inside Microsoft's own templates - and `undefined` for
 * one that nothing binds. The two are different: the first says "in no
 * namespace", the second is an error.
 */
export function resolvePrefix(element: XElement, prefix: string): string | undefined {
  if (prefix === 'xml') return 'http://www.w3.org/XML/1998/namespace';
  if (prefix === 'xmlns') return 'http://www.w3.org/2000/xmlns/';
  for (let node: XElement | undefined = element; node !== undefined; node = node.parent) {
    // Step over ancestors that declare nothing without touching their
    // attribute lists. Without this the walk is O(depth x attributes-per-
    // ancestor) and a legal 1.7 MB document - depth 1001, 200 attributes an
    // element, both inside their limits - took 37 seconds in
    // `undeclaredPrefixes`, which calls this once per element and again per
    // prefixed attribute. Nothing counted the product.
    if (!node.hasNamespaceDeclarations) continue;
    for (const attr of node.attributes) {
      if (
        prefix === '' ? attr.qname === 'xmlns' : attr.prefix === 'xmlns' && attr.local === prefix
      ) {
        return attr.value;
      }
    }
  }
  return prefix === '' ? '' : undefined;
}

/** The namespace URI of an element, or `''` if it is in none. */
export function namespaceOf(element: XElement): string | undefined {
  return resolvePrefix(element, element.prefix);
}

/**
 * The namespace URI of an attribute.
 *
 * An unprefixed attribute is in **no namespace** - it does not pick up the
 * default. *Namespaces in XML* §6.2 is explicit about this, and getting it
 * wrong makes every unprefixed attribute in a slide part appear to be in the
 * PresentationML namespace.
 */
export function attributeNamespaceOf(owner: XElement, attr: XAttribute): string | undefined {
  return attr.prefix === '' ? '' : resolvePrefix(owner, attr.prefix);
}

/** Every prefix in scope at this element, the nearest binding winning. */
export function namespaceScope(element: XElement): Map<string, string> {
  const chain: XElement[] = [];
  for (let node: XElement | undefined = element; node !== undefined; node = node.parent) {
    chain.push(node);
  }
  const scope = new Map<string, string>();
  for (let i = chain.length - 1; i >= 0; i--) {
    for (const [prefix, uri] of declaredNamespaces(chain[i]!)) scope.set(prefix, uri);
  }
  return scope;
}

/**
 * A prefix already in scope here for this namespace, for writing a new name.
 *
 * Nearest binding first, and in document order among an element's own
 * declarations, so the answer is deterministic on a document that binds one URI
 * to two prefixes. A non-empty prefix wins over the default: `<p:extLst>` is
 * what PowerPoint writes and what a reader of the file expects, even where an
 * unprefixed `<extLst>` under a default-namespace declaration would mean the
 * same thing.
 *
 * `undefined` means nothing in scope binds it, and a caller that wants that
 * namespace has to declare it - which is an edit, and a decision this function
 * deliberately does not make on anyone's behalf.
 */
export function prefixFor(element: XElement, namespace: string): string | undefined {
  let fallback: string | undefined;
  for (let node: XElement | undefined = element; node !== undefined; node = node.parent) {
    if (!node.hasNamespaceDeclarations) continue;
    for (const attr of node.attributes) {
      if (attr.value !== namespace) continue;
      if (attr.prefix === 'xmlns') return attr.local;
      if (attr.qname === 'xmlns') fallback ??= '';
    }
    if (fallback !== undefined) return fallback;
  }
  return undefined;
}

/**
 * Every prefix the document binds, with the set of URIs each is bound to.
 *
 * A prefix mapping to more than one URI is legal and real, which is why this
 * returns a set rather than a string.
 *
 * `checkRoundTrip` compares this across a round trip - the check that catches a
 * prefix rewrite before it silently disables an `mc:AlternateContent` branch.
 * It is not the whole of that check, and on its own it would not be enough: the
 * declaration map cannot see a rewrite that renames a prefix consistently at
 * both its binding and its uses, which is exactly the rewrite `XMLSerializer`
 * is permitted to perform. So the qualified name of every element and attribute
 * is compared too.
 */
export function prefixMap(document: XDocument): Map<string, Set<string>> {
  const map = new Map<string, Set<string>>();
  for (const element of descendantElements(document.root)) {
    for (const [prefix, uri] of declaredNamespaces(element)) {
      const uris = map.get(prefix);
      if (uris === undefined) map.set(prefix, new Set([uri]));
      else uris.add(uri);
    }
  }
  return map;
}

/** Any prefix used in a name that nothing in scope binds. */
export function undeclaredPrefixes(document: XDocument): { prefix: string; at: number }[] {
  const missing: { prefix: string; at: number }[] = [];
  for (const element of descendantElements(document.root)) {
    if (element.prefix !== '' && resolvePrefix(element, element.prefix) === undefined) {
      missing.push({ prefix: element.prefix, at: element.start });
    }
    for (const attr of element.attributes) {
      if (attr.prefix === '' || attr.prefix === 'xmlns') continue;
      if (resolvePrefix(element, attr.prefix) === undefined) {
        missing.push({ prefix: attr.prefix, at: attr.start });
      }
    }
  }
  return missing;
}

// ------------------------------------------------------------------ xml:space

/**
 * The effective `xml:space` at an element, per XML 1.0 §2.10: the value is
 * inherited by descendants until one overrides it.
 *
 * Worth being clear about what this is *not* for. The plan's rule that an
 * `<a:t>` with leading or trailing whitespace needs `xml:space="preserve"` does
 * not hold for PresentationML: there is no `xml:space` anywhere in our
 * 37-package corpus, 120 `<a:t>` elements carry edge whitespace without it, and
 * a slide part built with trailing spaces and no `xml:space` round-trips
 * through PowerPoint with the spaces intact. So this reads the attribute. It
 * never synthesizes one.
 */
export function xmlSpace(element: XElement): 'default' | 'preserve' {
  for (let node: XElement | undefined = element; node !== undefined; node = node.parent) {
    const declared = attributeValue(node, 'xml:space');
    if (declared === 'preserve') return 'preserve';
    if (declared === 'default') return 'default';
  }
  return 'default';
}

// ------------------------------------------------------------------- coverage

/** A place where a tree's spans fail to account for the source exactly. */
export interface CoverageGap {
  readonly kind: 'gap' | 'overlap' | 'malformed';
  readonly at: number;
  readonly detail: string;
}

/**
 * Check that a start tag's own offsets tile it, exactly as the children tile
 * the element.
 *
 * This exists because of a hole found by mutation testing, and the hole is
 * instructive. `checkTreeCoverage` tiled document children and element
 * children; `checkSpanCoverage` tiled token `start`/`end`. Neither ever read
 * `nameEnd`, `trailingSpaceStart`, or any of an attribute's six offsets - so
 * the offsets that exist *solely* for sub-phase 0.6's rebuild path were the
 * only ones nothing verified.
 *
 * Introducing an off-by-one into `trailingSpaceStart` for tags with two or more
 * attributes passed all 123 unit tests, both coverage checks, the 4000-case
 * fuzz slice, and would have passed 0.5's byte-identical corpus gate too -
 * because in a round trip nothing is dirty and every element re-emits as
 * `slice(start, end)`. The first failure would have been in 0.6, on a user's
 * edited deck, rewriting `<p:cNvPr id="2" name="Title 1" />` as
 * `<p:cNvPr id="2" name="Title 1"/>`.
 *
 * The zero-attribute arm below is not redundant with the closing-delimiter
 * check: `<a:bodyPr  />` with `trailingSpaceStart` one too high still leaves a
 * slice of `" />"`, which matches the pattern.
 */
function checkTagInterior(source: string, element: XElement, gaps: CoverageGap[]): void {
  const where = '<' + element.qname + '>';
  const note = (at: number, detail: string): void => {
    gaps.push({ kind: 'malformed', at, detail: where + ' ' + detail });
  };

  if (source.slice(element.start + 1, element.nameEnd) !== element.qname) {
    note(element.start, 'nameEnd does not delimit the element name');
  }

  const attributes = element.attributes;
  if (attributes.length === 0) {
    if (element.nameEnd !== element.trailingSpaceStart) {
      note(element.nameEnd, 'nothing accounts for the gap between the name and the tag close');
    }
  } else {
    if (attributes[0]!.start !== element.nameEnd) {
      note(element.nameEnd, 'the first attribute does not begin where the name ends');
    }
    for (let i = 1; i < attributes.length; i++) {
      if (attributes[i]!.start !== attributes[i - 1]!.end) {
        note(attributes[i]!.start, 'attributes ' + (i - 1) + ' and ' + i + ' do not abut');
      }
    }
    if (attributes[attributes.length - 1]!.end !== element.trailingSpaceStart) {
      note(
        element.trailingSpaceStart,
        'the last attribute does not end where the tag close begins',
      );
    }
  }

  for (const attr of attributes) {
    if (source.slice(attr.nameStart, attr.nameEnd) !== attr.qname) {
      note(attr.nameStart, 'attribute "' + attr.qname + '" nameStart/nameEnd do not delimit it');
    }
    if (!(
      attr.start <= attr.nameStart &&
      attr.nameStart < attr.nameEnd &&
      attr.nameEnd < attr.valueStart
    )) {
      note(attr.start, 'attribute "' + attr.qname + '" offsets are out of order');
    }
    if (attr.valueStart > attr.valueEnd || attr.end !== attr.valueEnd + 1) {
      note(
        attr.valueStart,
        'attribute "' + attr.qname + '" value offsets do not close on the quote',
      );
    }
    if (source.charCodeAt(attr.valueStart - 1) !== attr.quote.charCodeAt(0)) {
      note(
        attr.valueStart,
        'attribute "' + attr.qname + '" does not open with the quote it records',
      );
    }
  }

  const closeFrom = element.trailingSpaceStart;
  const closeTo = element.selfClosing ? element.end : element.openTagEnd;
  if (!/^[ \t\r\n]*\/?>$/.test(source.slice(closeFrom, closeTo))) {
    note(closeFrom, 'the tail of the start tag is not whitespace followed by ">" or "/>"');
  }
}

/**
 * Check that the tree accounts for every character of the source.
 *
 * The document's children must tile `[bom ? 1 : 0, source.length)`, and every
 * element's children must tile the range between its tags. This is the real
 * gate for sub-phase 0.4, and it is stronger than "parse the corpus without
 * error" by exactly the margin that matters: an off-by-one in a span raises no
 * error at all, parses every part of every deck happily, and then silently
 * drops a character the first time 0.5 re-serializes from spans.
 *
 * **Meaningful only while the tree is clean.** Once anything has been edited,
 * the spans of the dirty nodes no longer describe the source and are not meant
 * to; a synthesized node's span is `[0, 0)`. Run this on a freshly parsed
 * document, or on one whose edits have all been undone - where it is a real
 * check, because an exact undo restores the spans along with everything else.
 * `checkRoundTrip` is the instrument for a document with edits standing.
 */
export function checkTreeCoverage(document: XDocument): CoverageGap[] {
  const gaps: CoverageGap[] = [];

  const tile = (nodes: readonly XNode[], from: number, to: number, where: string): void => {
    let cursor = from;
    for (const node of nodes) {
      if (node.start > cursor) {
        gaps.push({
          kind: 'gap',
          at: cursor,
          detail: where + ': ' + JSON.stringify(document.source.slice(cursor, node.start)),
        });
      } else if (node.start < cursor) {
        gaps.push({ kind: 'overlap', at: node.start, detail: where + ': spans overlap' });
      }
      cursor = node.end;
    }
    if (cursor !== to) {
      gaps.push({
        kind: 'gap',
        at: cursor,
        detail: where + ' trailing: ' + JSON.stringify(document.source.slice(cursor, to)),
      });
    }
  };

  tile(document.children, document.bom ? 1 : 0, document.source.length, 'document');
  for (const element of descendantElements(document.root)) {
    checkTagInterior(document.source, element, gaps);
    if (element.selfClosing) {
      if (element.children.length > 0) {
        gaps.push({
          kind: 'malformed',
          at: element.start,
          detail: '<' + element.qname + '/> has children',
        });
      }
      continue;
    }
    tile(element.children, element.openTagEnd, element.closeTagStart, '<' + element.qname + '>');
  }
  return gaps;
}
