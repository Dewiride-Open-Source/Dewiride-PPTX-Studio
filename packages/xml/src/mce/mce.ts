/**
 * Markup Compatibility and Extensibility - ISO/IEC 29500-3.
 *
 * MCE is how a `.pptx` written by a newer PowerPoint stays readable by an older
 * one. `mc:AlternateContent` is a switch statement over namespaces:
 *
 * ```xml
 * <mc:AlternateContent xmlns:mc="…/markup-compatibility/2006">
 *   <mc:Choice Requires="a14"><!-- ink, if you understand a14 --></mc:Choice>
 *   <mc:Fallback><p:pic><!-- a picture of the ink, if you do not --></p:pic></mc:Fallback>
 * </mc:AlternateContent>
 * ```
 *
 * ## The reason this package exists
 *
 * **`@Requires` holds prefixes, not URIs.** So does `mc:Ignorable`, and so does
 * `mc:MustUnderstand`. A serializer that rewrites `a14` to `ns0` - which the DOM
 * specification explicitly permits, and which `XMLSerializer` does - leaves a
 * `Requires` naming a prefix that no longer resolves, and turns a document that
 * degrades gracefully into one that is an error to read. Everything about the
 * tokenizer and the serializer follows from not being allowed to do that.
 *
 * Our corpus has four `mc:AlternateContent` elements, three requiring `c14` and
 * one requiring `p14` - and `p14` is bound to `…/powerpoint/2010/main` in 46
 * parts of that same corpus and to `…/powerpoint/2007/7/12/main` in 8 others.
 * There is no document-wide prefix table that gets that right; the binding has
 * to be read where the `Requires` was written.
 *
 * ## This module never edits
 *
 * It answers "what would a consumer that understands *these* namespaces see?",
 * for the read side - the semantic projection sub-phase 2.9 builds. It does not
 * remove a branch, rewrite one, or mark anything dirty. The plan's rule is
 * absolute: *never rewrite an `mc:AlternateContent` branch you don't
 * understand*, and the way to keep it is to have no code here that could.
 */

import { XmlError } from '../errors.js';
import { NS } from './namespaces.js';
import { qualifiedKey } from '../edit/schema-order.js';
import {
  declaredNamespaces,
  resolvePrefix,
  namespaceOf,
  type XElement,
  type XNode,
} from '../parse/xnode.js';
import type { XAttribute } from '../parse/tokenizer.js';

/** The three MCE element names, unprefixed. */
export const MC_ALTERNATE_CONTENT = 'AlternateContent';
export const MC_CHOICE = 'Choice';
export const MC_FALLBACK = 'Fallback';

/** True if this element is `mc:AlternateContent`, whatever prefix it was written with. */
export function isAlternateContent(element: XElement): boolean {
  return element.local === MC_ALTERNATE_CONTENT && namespaceOf(element) === NS.mc;
}

function isMc(element: XElement, local: string): boolean {
  return element.local === local && namespaceOf(element) === NS.mc;
}

function fail(message: string, at: number, name?: string): never {
  throw new XmlError('ERR_INVALID_MCE', message, { offset: at, name });
}

/** Whitespace-delimited, per §A.1. Empty entries are dropped rather than being an error. */
function tokens(value: string): string[] {
  return value.split(/\s+/u).filter((t) => t !== '');
}

/**
 * Resolve a whitespace-delimited prefix list to namespace URIs, at the element
 * that wrote it.
 *
 * A prefix nothing binds is a hard failure, not a namespace we happen not to
 * support. The difference matters: treating it as unsupported would silently
 * pick a different `mc:Choice`, which is exactly the corruption this package
 * was built to prevent, arrived at from the other direction.
 */
function resolveAll(element: XElement, attribute: XAttribute, what: string): string[] {
  const names = tokens(attribute.value);
  if (names.length === 0) {
    fail('mc:' + what + ' names no prefix', attribute.valueStart, attribute.qname);
  }
  return names.map((prefix) => {
    const uri = resolvePrefix(element, prefix);
    if (uri === undefined || uri === '') {
      fail(
        'mc:' + what + ' names the prefix "' + prefix + '", which nothing in scope binds',
        attribute.valueStart,
        prefix,
      );
    }
    return uri;
  });
}

/**
 * A compatibility-rule attribute: `mc:Ignorable`, `mc:MustUnderstand`,
 * `mc:ProcessContent`, `mc:PreserveElements`, `mc:PreserveAttributes`.
 *
 * These sit on arbitrary elements, so they carry a prefix bound to the MC
 * namespace. A *bare* `Ignorable` is not one of them: *Namespaces in XML* §6.2
 * puts an unprefixed attribute in **no** namespace, never in the element's
 * default one, so a bare name here would be an ordinary attribute that happens
 * to share a spelling.
 */
function mcAttribute(element: XElement, local: string): XAttribute | undefined {
  for (const attribute of element.attributes) {
    if (attribute.local !== local || attribute.prefix === '') continue;
    if (resolvePrefix(element, attribute.prefix) === NS.mc) return attribute;
  }
  return undefined;
}

/**
 * `Requires`, which is the exception, and the exception is the whole point.
 *
 * It is an attribute *of* `mc:Choice` rather than an annotation placed on
 * foreign markup, so it is written **unprefixed** - `<mc:Choice Requires="a14">`
 * - and by §6.2 that puts it in no namespace. Looking for `mc:Requires` finds
 * nothing in any real file, and a walker that then treats the Choice as
 * requiring nothing selects the *first* branch of every `mc:AlternateContent` in
 * the document.
 */
function requiresAttribute(choice: XElement): XAttribute | undefined {
  for (const attribute of choice.attributes) {
    if (attribute.prefix === '' && attribute.qname === 'Requires') return attribute;
  }
  return undefined;
}

/**
 * Namespaces an ancestor-or-self declared ignorable, as URIs.
 *
 * The compatibility-rule attributes apply to the element carrying them and to
 * everything beneath, and a descendant may add more. There is no way to
 * un-ignore, so this is a union up the tree and never a replacement. Each
 * prefix is resolved at the element that named it, because that is the only
 * scope in which it is guaranteed to mean anything.
 */
export function ignorableNamespaces(element: XElement): Set<string> {
  const found = new Set<string>();
  for (let node: XElement | undefined = element; node !== undefined; node = node.parent) {
    const attribute = mcAttribute(node, 'Ignorable');
    if (attribute !== undefined)
      for (const uri of resolveAll(node, attribute, 'Ignorable')) found.add(uri);
  }
  return found;
}

/** Namespaces an ancestor-or-self declared must-understand, as URIs. */
export function mustUnderstandNamespaces(element: XElement): Set<string> {
  const found = new Set<string>();
  for (let node: XElement | undefined = element; node !== undefined; node = node.parent) {
    const attribute = mcAttribute(node, 'MustUnderstand');
    if (attribute !== undefined) {
      for (const uri of resolveAll(node, attribute, 'MustUnderstand')) found.add(uri);
    }
  }
  return found;
}

/**
 * Elements whose *content* survives when the element itself is ignored, as
 * `{namespace}local` keys.
 *
 * `mc:ProcessContent` holds qualified names rather than bare prefixes, which is
 * the one place in MCE where that is true, and getting it wrong yields a set
 * that silently matches nothing.
 */
export function processContentNames(element: XElement): Set<string> {
  const found = new Set<string>();
  for (let node: XElement | undefined = element; node !== undefined; node = node.parent) {
    const attribute = mcAttribute(node, 'ProcessContent');
    if (attribute === undefined) continue;
    for (const qname of tokens(attribute.value)) {
      const colon = qname.indexOf(':');
      if (colon <= 0) {
        fail(
          'mc:ProcessContent names "' + qname + '", which is not a qualified name',
          attribute.valueStart,
          qname,
        );
      }
      const uri = resolvePrefix(node, qname.slice(0, colon));
      if (uri === undefined || uri === '') {
        fail(
          'mc:ProcessContent names the prefix "' +
            qname.slice(0, colon) +
            '", which nothing in scope binds',
          attribute.valueStart,
          qname,
        );
      }
      found.add(qualifiedKey(uri, qname.slice(colon + 1)));
    }
  }
  return found;
}

/**
 * The branch of an `mc:AlternateContent` a consumer supporting `supported`
 * would take, or `undefined` if it would take none.
 *
 * The first `mc:Choice` **all** of whose required namespaces are supported
 * wins; failing that, `mc:Fallback`; failing that, nothing, and the element
 * contributes no content at all. That last case is legal and does happen.
 *
 * The whole element is checked before anything is returned, rather than
 * stopping at the winning branch. Short-circuiting would be faster on an
 * element with at most three children, and it would mean that a document whose
 * *second* `mc:Choice` names a prefix nothing binds reads perfectly on the
 * machine where the first Choice happens to be supported, and is an error on
 * the machine where it is not. Bugs that appear only on someone else's build of
 * Office are the expensive kind.
 */
export function selectAlternateContent(
  alternateContent: XElement,
  supported: ReadonlySet<string>,
): XElement | undefined {
  let fallback: XElement | undefined;
  let selected: XElement | undefined;
  let choices = 0;

  for (const child of alternateContent.children) {
    if (child.type !== 'element') continue;
    if (isMc(child, MC_CHOICE)) {
      if (fallback !== undefined) {
        fail('an mc:Choice follows the mc:Fallback', child.start, child.qname);
      }
      choices++;
      const requires = requiresAttribute(child);
      if (requires === undefined) {
        fail('an mc:Choice has no Requires attribute', child.start, child.qname);
      }
      const needed = resolveAll(child, requires, 'Choice/@Requires');
      if (selected === undefined && needed.every((uri) => supported.has(uri))) selected = child;
      continue;
    }
    if (isMc(child, MC_FALLBACK)) {
      if (fallback !== undefined) {
        fail('an mc:AlternateContent has two mc:Fallback children', child.start, child.qname);
      }
      fallback = child;
      continue;
    }
    fail(
      '<' + child.qname + '> is not allowed inside mc:AlternateContent',
      child.start,
      child.qname,
    );
  }

  if (choices === 0) {
    fail('an mc:AlternateContent has no mc:Choice', alternateContent.start, alternateContent.qname);
  }
  return selected ?? fallback;
}

/**
 * The children a consumer supporting `supported` would process, in order.
 *
 * `mc:AlternateContent` is transparent: the selected branch's own children take
 * its place, and they are walked too, so a nested `AlternateContent` or an
 * ignorable wrapper inside a branch resolves as well. An element in an ignorable
 * namespace we do not support disappears - unless `mc:ProcessContent` names it,
 * in which case its children are processed where it stood.
 *
 * An element in a namespace that is neither supported nor declared ignorable is
 * **yielded**, not dropped. MCE says a producer should have marked it, but a
 * consumer that deletes what a producer forgot to annotate loses content, and
 * for an editor whose thesis is preservation that is the worse error. The
 * caller can ignore what it does not recognise; it cannot recover what this
 * function swallowed.
 */
export function* effectiveChildren(
  element: XElement,
  supported: ReadonlySet<string>,
): IterableIterator<XNode> {
  const ignorable = ignorableNamespaces(element);
  const processContent = processContentNames(element);

  for (const child of element.children) {
    if (child.type !== 'element') {
      yield child;
      continue;
    }
    if (isAlternateContent(child)) {
      const branch = selectAlternateContent(child, supported);
      if (branch !== undefined) yield* effectiveChildren(branch, supported);
      continue;
    }
    const namespace = namespaceOf(child);
    if (namespace !== undefined && ignorable.has(namespace) && !supported.has(namespace)) {
      if (processContent.has(qualifiedKey(namespace, child.local))) {
        yield* effectiveChildren(child, supported);
      }
      continue;
    }
    yield child;
  }
}

/**
 * The attributes a consumer supporting `supported` would read.
 *
 * The MCE attributes themselves are never among them - they are instructions to
 * the reader, not content - and neither is an attribute in an ignorable
 * namespace we do not support. An unprefixed attribute is in no namespace and
 * always survives.
 */
export function* effectiveAttributes(
  element: XElement,
  supported: ReadonlySet<string>,
): IterableIterator<XAttribute> {
  const ignorable = ignorableNamespaces(element);
  for (const attribute of element.attributes) {
    if (attribute.prefix === '') {
      yield attribute;
      continue;
    }
    if (attribute.prefix === 'xmlns') continue;
    const namespace = resolvePrefix(element, attribute.prefix);
    if (namespace === NS.mc) continue;
    if (namespace !== undefined && ignorable.has(namespace) && !supported.has(namespace)) continue;
    yield attribute;
  }
}

/** One way a document's markup-compatibility annotations do not hold together. */
export interface McProblem {
  readonly at: number;
  readonly detail: string;
}

/**
 * Check every MCE construct in a subtree without throwing.
 *
 * The same shape as `checkTreeCoverage` and `checkRoundTrip`, and for the same
 * reason: the corpus gate wants to count problems across thousands of parts
 * rather than stop at the first. It also reports `mc:MustUnderstand` naming a
 * namespace outside `supported`, which is the one MCE condition that says a
 * consumer should refuse the document outright rather than degrade.
 */
export function checkMarkupCompatibility(
  root: XElement,
  supported: ReadonlySet<string>,
): McProblem[] {
  const problems: McProblem[] = [];
  const stack: XElement[] = [root];
  while (stack.length > 0) {
    const element = stack.pop()!;
    try {
      for (const uri of mustUnderstandNamespaces(element)) {
        if (!supported.has(uri)) {
          problems.push({
            at: element.start,
            detail: 'mc:MustUnderstand requires "' + uri + '", which this consumer does not know',
          });
        }
      }
      ignorableNamespaces(element);
      processContentNames(element);
      if (isAlternateContent(element)) selectAlternateContent(element, supported);
    } catch (error) {
      problems.push({
        at: element.start,
        detail: error instanceof XmlError ? error.message : String(error),
      });
    }
    for (const child of element.children) if (child.type === 'element') stack.push(child);
  }
  return problems;
}

/**
 * Every namespace the document binds anywhere, as a set of URIs.
 *
 * Useful for answering "what would this file need me to understand?" before
 * deciding what to pass as `supported`. Deliberately built by walking, not from
 * the root element's declarations, because 213 declarations in the corpus sit
 * below the root.
 */
export function boundNamespaces(root: XElement): Set<string> {
  const found = new Set<string>();
  const stack: XElement[] = [root];
  while (stack.length > 0) {
    const element = stack.pop()!;
    if (element.hasNamespaceDeclarations) {
      for (const uri of declaredNamespaces(element).values()) if (uri !== '') found.add(uri);
    }
    for (const child of element.children) if (child.type === 'element') stack.push(child);
  }
  return found;
}
