import { NS, XmlTokenizer, type XAttribute, type XmlTokenizerLimits } from '@pptx-studio/xml';
import { FEATURE_RULES, type FeatureRule } from './features.js';
import { OOXML_NS } from './namespaces.js';

/**
 * The scan runs over **tokens, not trees**.
 *
 * `parseXml` would be the obvious thing to reach for and it is the wrong tool
 * here for two independent reasons.
 *
 * The first is memory. A census of a 200 MB deck is a few kilobytes of
 * counters; a tree of the same deck is every element as an object, each with a
 * parent pointer, plus the whole decoded source string held alive by every
 * node's byte offsets. The scanner below holds one part's source and a handful
 * of maps, so its working set does not grow with the deck.
 *
 * The second is that a tree is the wrong shape for the question. Counting how
 * many `a:tbl` a package contains does not need any node to know its children.
 * The tokenizer hands over a flat stream and the only state a census needs on
 * top of it is a namespace scope stack.
 *
 * What that costs: no structural queries. "Which shapes are inside a group that
 * is inside another group" is not answerable here, and should not be - that is
 * what the document model in Phase 2 is for. A census answers *what is in this
 * file*, never *how it is arranged*.
 *
 * ## Why the scope holds objects rather than URIs
 *
 * The obvious shape for a namespace scope is `Map<prefix, uri>`, and the first
 * version of this file used it. Per element that meant hashing a sixty-character
 * URI to reach the namespace's counters, and building a `uri + ' ' + local` key
 * to look up the feature table - one string allocation for every element in the
 * package. On a deck with two million elements that was the dominant cost: the
 * scan ran at 10 MiB/s, slow enough to be mistaken for the tokenizer being slow,
 * which would have been the wrong thing to go and optimise.
 *
 * So a prefix binds to a `Binding`, which already holds the namespace's
 * counters and its slice of the feature table. Per element the work is one map
 * lookup on a short prefix and one on a short local name. Nothing is
 * concatenated, and nothing sixty characters long is hashed.
 */

interface NamespaceUse {
  count: number;
  readonly prefixes: Set<string>;
  required: boolean;
  ignorable: boolean;
}

/**
 * A prefix, resolved once.
 *
 * Built when a namespace is declared rather than when it is used, so the cost
 * is per `xmlns` attribute in the package instead of per element.
 */
interface Binding {
  readonly uri: string;
  readonly use: NamespaceUse;
  /** This namespace's feature rules by local name, or `undefined` for none. */
  readonly features: ReadonlyMap<string, FeatureRule> | undefined;
  readonly isDrawingMl: boolean;
  readonly isMc: boolean;
}

/** Namespace bindings in scope, keyed by prefix. The default namespace is `''`. */
type Scope = ReadonlyMap<string, Binding>;

const A = OOXML_NS.a;
const MC = OOXML_NS.mc;

/** `uri -> local -> rule`, so the per-element lookup never has to build a key. */
const FEATURES_BY_URI: ReadonlyMap<string, ReadonlyMap<string, FeatureRule>> = (() => {
  const byUri = new Map<string, Map<string, FeatureRule>>();
  for (const rule of FEATURE_RULES) {
    let byLocal = byUri.get(rule.uri);
    if (byLocal === undefined) {
      byLocal = new Map();
      byUri.set(rule.uri, byLocal);
    }
    byLocal.set(rule.local, rule);
  }
  return byUri;
})();

/**
 * Everything a census accumulates, across every part.
 *
 * Mutable and shared: one of these is threaded through every part in the
 * package so the histograms are package-wide. `census.ts` freezes it into the
 * plain-JSON report at the end.
 */
export interface ScanTotals {
  readonly elements: Map<string, number>;
  readonly namespaces: Map<string, NamespaceUse>;
  readonly features: Map<string, number>;
  readonly presetGeometry: Map<string, number>;
  readonly typefaces: Map<string, number>;
  readonly languages: Map<string, number>;
  paragraphs: number;
  runs: number;
  characters: number;
  fields: number;
  hardBreaks: number;
  bulletChar: number;
  bulletAutoNum: number;
  bulletBlip: number;
  normAutofit: number;
  shapeAutofit: number;
  /** Total elements across every part scanned. */
  totalElements: number;
  /** Bindings already built, so every part in the package shares the objects. */
  readonly bindings: Map<string, Binding>;
}

export function createTotals(): ScanTotals {
  return {
    elements: new Map(),
    namespaces: new Map(),
    features: new Map(),
    presetGeometry: new Map(),
    typefaces: new Map(),
    languages: new Map(),
    paragraphs: 0,
    runs: 0,
    characters: 0,
    fields: 0,
    hardBreaks: 0,
    bulletChar: 0,
    bulletAutoNum: 0,
    bulletBlip: 0,
    normAutofit: 0,
    shapeAutofit: 0,
    totalElements: 0,
    bindings: new Map(),
  };
}

/** What one part contributed, for the per-part table. */
export interface PartScan {
  readonly elements: number;
  readonly maxDepth: number;
  /**
   * Why the part is not well formed, or `null`.
   *
   * A token stream is not a tree, so nothing in the tokenizer objects to
   * `<a><b></a>` or to a document that simply stops - it hands over the tokens
   * it found and falls silent. The first version of this scanner therefore
   * reported a truncated slide as perfectly healthy, which is precisely the
   * wrong answer from a tool whose whole job is to describe a file that is
   * already suspect. Balance is checked here, against a stack of open names,
   * because this is the only place that has one.
   */
  readonly imbalance: string | null;
}

function bump(map: Map<string, number>, key: string): void {
  map.set(key, (map.get(key) ?? 0) + 1);
}

function bindingFor(totals: ScanTotals, uri: string): Binding {
  const existing = totals.bindings.get(uri);
  if (existing !== undefined) return existing;

  let use = totals.namespaces.get(uri);
  if (use === undefined) {
    use = { count: 0, prefixes: new Set(), required: false, ignorable: false };
    totals.namespaces.set(uri, use);
  }
  const binding: Binding = {
    uri,
    use,
    features: FEATURES_BY_URI.get(uri),
    isDrawingMl: uri === A,
    isMc: uri === MC,
  };
  totals.bindings.set(uri, binding);
  return binding;
}

/** `xml:` is bound implicitly and may never be redeclared. */
function rootScope(totals: ScanTotals): Scope {
  return new Map([['xml', bindingFor(totals, NS.xml)]]);
}

/** A scope with this tag's `xmlns` declarations layered over its parent's. */
function derive(parent: Scope, attributes: readonly XAttribute[], totals: ScanTotals): Scope {
  const next = new Map(parent);
  for (const attribute of attributes) {
    // `xmlns="..."` is unprefixed with local `xmlns`; `xmlns:a="..."` is
    // prefixed `xmlns` with local `a`. The two spellings mean different things
    // and both have to be read.
    const isDefault = attribute.prefix === '' && attribute.local === 'xmlns';
    if (!isDefault && attribute.prefix !== 'xmlns') continue;
    const prefix = isDefault ? '' : attribute.local;
    const uri = attribute.value;
    if (uri === '') {
      // `xmlns=""` undeclares the default namespace. Legal, and it happens.
      next.delete(prefix);
      continue;
    }
    const binding = bindingFor(totals, uri);
    next.set(prefix, binding);
    // Recorded on declaration as well as on use, so a namespace a deck declares
    // and never uses still shows up. That gap is informative: it is usually the
    // fingerprint of the producer rather than of the content.
    binding.use.prefixes.add(prefix);
  }
  return next;
}

/**
 * Prefixes named in `mc:Ignorable` / `@Requires`, resolved against the scope
 * they were written in.
 *
 * These attributes hold prefixes, not URIs, and the same prefix resolves to
 * different namespaces in different parts of one corpus. Resolving them here -
 * rather than storing the spelling - is the whole reason the scanner tracks
 * scope at all.
 */
function markNamedNamespaces(scope: Scope, value: string, field: 'required' | 'ignorable'): void {
  for (const prefix of value.split(/\s+/)) {
    if (prefix === '') continue;
    const binding = scope.get(prefix);
    if (binding === undefined) continue;
    binding.use[field] = true;
  }
}

/** Typeface-bearing DrawingML elements, all in the `a` namespace. */
const TYPEFACE_ELEMENTS = new Set(['latin', 'ea', 'cs', 'sym', 'font', 'buFont']);

/**
 * Scan one XML part, folding what it holds into `totals`.
 *
 * Throws whatever the tokenizer throws - always an `XmlError`, never a
 * `RangeError`. The caller decides whether a part that will not tokenize ends
 * the census or becomes a problem entry; for `pptx-studio inspect` it is the
 * second, because a report that stops at the first broken part is exactly the
 * report you cannot use when a part is broken.
 */
export function scanXmlPart(
  source: string,
  totals: ScanTotals,
  limits?: XmlTokenizerLimits,
): PartScan {
  const tokenizer = new XmlTokenizer(source, limits);
  const noNamespace = bindingFor(totals, '');

  /** Saved parent scopes, one slot per open element. `null` means unchanged. */
  const saved: (Scope | null)[] = [];
  /** Open element names, innermost last. Only for the balance check. */
  const open: string[] = [];
  let imbalance: string | null = null;
  let scope: Scope = rootScope(totals);
  let depth = 0;
  let maxDepth = 0;
  let elements = 0;
  /** Depth of the innermost open `a:t`, or 0. Text elements do not nest. */
  let inText = 0;

  for (let token = tokenizer.next(); token !== undefined; token = tokenizer.next()) {
    if (token.type === 'endTag') {
      const expected = open.pop();
      if (imbalance === null && expected !== token.qname) {
        imbalance =
          expected === undefined
            ? '</' + token.qname + '> closes an element that was never opened'
            : '</' + token.qname + '> closes <' + expected + '>';
      }
      const restore = saved.pop();
      if (restore != null) scope = restore;
      if (inText === depth) inText = 0;
      depth -= 1;
      continue;
    }

    if (token.type === 'text' || token.type === 'cdata') {
      // Counted in UTF-16 code units, which is what every downstream measure -
      // canvas advance, line breaking, ProseMirror position - also counts in.
      if (inText > 0) totals.characters += token.value.length;
      continue;
    }

    if (token.type !== 'startTag' && token.type !== 'emptyElementTag') continue;

    const selfClosing = token.type === 'emptyElementTag';
    const parent = scope;
    const derived = token.hasNamespaceDeclarations
      ? derive(scope, token.attributes, totals)
      : undefined;
    if (derived !== undefined) scope = derived;

    elements += 1;
    depth += 1;
    if (depth > maxDepth) maxDepth = depth;

    const binding = scope.get(token.prefix) ?? noNamespace;
    binding.use.count += 1;
    bump(totals.elements, token.qname);

    const rule = binding.features?.get(token.local);
    if (rule !== undefined) bump(totals.features, rule.key);

    if (binding.isDrawingMl) countDrawingMl(totals, token.local, token.attributes);
    if (token.attributes.length > 0) countAttributes(totals, scope, binding, token.attributes);

    if (binding.isDrawingMl && token.local === 't' && !selfClosing) inText = depth;

    if (selfClosing) {
      // Nothing was pushed for a self-closing tag, so nothing is popped: it
      // opened and closed inside this one token.
      depth -= 1;
      if (derived !== undefined) scope = parent;
    } else {
      saved.push(derived === undefined ? null : parent);
      open.push(token.qname);
    }
  }

  if (imbalance === null && open.length > 0) {
    imbalance = 'the part ends with <' + (open[open.length - 1] ?? '?') + '> still open';
  }

  totals.totalElements += elements;
  return { elements, maxDepth, imbalance };
}

function countAttributes(
  totals: ScanTotals,
  scope: Scope,
  owner: Binding,
  attributes: readonly XAttribute[],
): void {
  for (const attribute of attributes) {
    if (attribute.prefix === '') {
      // Unprefixed attributes are in **no** namespace: they do not pick up the
      // default. That is why `@Requires` is looked for here rather than among
      // the prefixed names below - it is written `Requires`, never
      // `mc:Requires`, and looking for the prefixed spelling finds nothing.
      if (attribute.local === 'lang') {
        if (attribute.value !== '') bump(totals.languages, attribute.value);
      } else if (owner.isMc && attribute.local === 'Requires') {
        markNamedNamespaces(scope, attribute.value, 'required');
      }
      continue;
    }
    if (attribute.prefix === 'xmlns') continue;
    const binding = scope.get(attribute.prefix);
    if (binding === undefined) continue;
    binding.use.prefixes.add(attribute.prefix);
    if (binding.isMc && attribute.local === 'Ignorable') {
      markNamedNamespaces(scope, attribute.value, 'ignorable');
    }
  }
}

/** The DrawingML elements a census counts beyond the feature table. */
function countDrawingMl(
  totals: ScanTotals,
  local: string,
  attributes: readonly XAttribute[],
): void {
  switch (local) {
    case 'p':
      totals.paragraphs += 1;
      return;
    case 'r':
      totals.runs += 1;
      return;
    case 'br':
      totals.hardBreaks += 1;
      return;
    case 'fld':
      totals.fields += 1;
      return;
    case 'buChar':
      totals.bulletChar += 1;
      return;
    case 'buAutoNum':
      totals.bulletAutoNum += 1;
      return;
    case 'buBlip':
      totals.bulletBlip += 1;
      return;
    case 'normAutofit':
      totals.normAutofit += 1;
      return;
    case 'spAutoFit':
      totals.shapeAutofit += 1;
      return;
    case 'prstGeom': {
      const prst = attributeValue(attributes, 'prst');
      if (prst !== undefined) bump(totals.presetGeometry, prst);
      return;
    }
    default: {
      if (!TYPEFACE_ELEMENTS.has(local)) return;
      const typeface = attributeValue(attributes, 'typeface');
      // `+mj-lt` and `+mn-lt` are theme indirection, not typefaces. They are
      // counted under their own names on purpose: a deck whose runs all say
      // `+mn-lt` has no hardcoded fonts at all, which is worth being able to
      // see at a glance.
      if (typeface !== undefined && typeface !== '') bump(totals.typefaces, typeface);
      return;
    }
  }
}

/** An unprefixed attribute's value. Unprefixed attributes are in no namespace. */
function attributeValue(attributes: readonly XAttribute[], local: string): string | undefined {
  for (const attribute of attributes) {
    if (attribute.prefix === '' && attribute.local === local) return attribute.value;
  }
  return undefined;
}
