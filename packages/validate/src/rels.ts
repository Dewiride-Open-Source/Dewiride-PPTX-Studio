import {
  isRelationshipPartName,
  normalizePartName,
  resolveRelativeTarget,
  sourcePartNameForRels,
} from '@pptx-studio/opc';
import { attributeValue, childElements, type XElement } from '@pptx-studio/xml';
import type { Context } from './context.js';

/**
 * Relationship parts, read as XML rather than through `Relationships`.
 *
 * This is not duplication for its own sake, and the reason is worth stating
 * because the obvious code - call `store.relationships(part)` - is wrong here in
 * a way that is invisible until a rule fails to fire.
 *
 * `Relationships.parse` **refuses** a duplicate `Id`, an `Id` that is not an
 * `xsd:ID`, and an empty `Target`. It has to: it is the type the rest of the
 * codebase uses to look relationships up, and a collection with two `rId3`s
 * cannot answer the only question anyone asks it. But three of the twenty-nine
 * rules are *about* exactly those defects, and a rule that can only see files
 * its own parser already accepted can never report them. Going through that
 * parser would turn every one of those findings into "this part could not be
 * read", which is true and useless.
 *
 * So the rules read the markup. That also buys the thing the plan asks for: a
 * finding gets an XPath into the `.rels` part - `/Relationships/Relationship[4]`
 * - instead of a part name and an apology.
 */

const REL_NS = 'http://schemas.openxmlformats.org/package/2006/relationships';

export interface RawRelationship {
  readonly id: string;
  readonly type: string;
  readonly target: string;
  /** `Internal` unless the attribute says otherwise; the attribute is optional. */
  readonly targetMode: string;
  /** The `<Relationship>` element, for the finding's location. */
  readonly element: XElement;
}

export interface RelsPart {
  /** The `.rels` part's own name, e.g. `/ppt/slides/_rels/slide1.xml.rels`. */
  readonly partName: string;
  /** The part it describes, e.g. `/ppt/slides/slide1.xml`. `/` for the root rels. */
  readonly source: string;
  readonly relationships: readonly RawRelationship[];
  readonly root: XElement;
}

/**
 * Every relationship part in the package, parsed.
 *
 * A `.rels` whose root is not `Relationships` in the OPC namespace is skipped
 * and recorded as a problem, because there is nothing sensible to say about it
 * rule by rule and saying it once is better than saying it four times.
 */
const CACHE = new WeakMap<Context, readonly RelsPart[]>();

export function readRelsParts(ctx: Context): readonly RelsPart[] {
  const cached = CACHE.get(ctx);
  if (cached !== undefined) return cached;
  const parts = scanRelsParts(ctx);
  CACHE.set(ctx, parts);
  return parts;
}

/**
 * Memoised above, because five of the twenty-nine rules want this list and each
 * would otherwise re-walk every `.rels` part in the package. The parsed trees
 * are already cached on the context; what is saved here is the element scan and
 * the allocation, which on a three-hundred-slide deck is five passes over a
 * thousand relationship parts for one answer that cannot have changed.
 *
 * Keyed by context rather than by store, so nothing survives one validation run
 * into the next.
 */
function scanRelsParts(ctx: Context): readonly RelsPart[] {
  const out: RelsPart[] = [];
  for (const partName of ctx.parts()) {
    if (!isRelationshipPartName(partName)) continue;
    const document = ctx.document(partName);
    if (document === null) continue;

    const root = document.root;
    if (root.local !== 'Relationships') {
      ctx.problem(partName, 'root element is <' + root.qname + '>, not <Relationships>');
      continue;
    }

    let source: string;
    try {
      source = partName === '/_rels/.rels' ? '/' : sourcePartNameForRels(partName);
    } catch {
      ctx.problem(partName, 'is not a well-formed relationship part name');
      continue;
    }

    const relationships: RawRelationship[] = [];
    for (const child of childElements(root)) {
      if (child.local !== 'Relationship') continue;
      relationships.push({
        id: attributeValue(child, 'Id') ?? '',
        type: attributeValue(child, 'Type') ?? '',
        target: attributeValue(child, 'Target') ?? '',
        targetMode: attributeValue(child, 'TargetMode') ?? 'Internal',
        element: child,
      });
    }
    out.push({ partName, source, relationships, root });
  }
  return out;
}

/** The OPC relationships namespace, for the rules that check the root element. */
export { REL_NS };

export interface ResolvedTarget {
  /** The absolute part name the target names, or `null` when it will not resolve. */
  readonly part: string | null;
  /** Why it would not resolve. `null` when it did. */
  readonly failure: string | null;
}

/**
 * Resolve an internal target against the folder of the part that owns the
 * `.rels` - never against the package root, and never against the `.rels`
 * part's own folder.
 *
 * The distinction is the whole of rule `V008`. `_rels/` is a directory in the
 * archive but not in the resolution model: a `Target="../media/image1.png"` in
 * `/ppt/slides/_rels/slide1.xml.rels` resolves against `/ppt/slides/`, giving
 * `/ppt/media/image1.png`. Resolving it against `/ppt/slides/_rels/` instead
 * yields `/ppt/slides/media/image1.png`, which is one directory too deep and is
 * why a rebuilt deck loses every image at once rather than one of them.
 */
export function resolveTarget(source: string, target: string): ResolvedTarget {
  try {
    return { part: resolveRelativeTarget(source, target), failure: null };
  } catch (error) {
    return { part: null, failure: error instanceof Error ? error.message : String(error) };
  }
}

/** Relationship ids declared in a part's `.rels`, for reference checking. */
export function idsOf(rels: RelsPart | undefined): ReadonlySet<string> {
  const ids = new Set<string>();
  for (const rel of rels?.relationships ?? []) ids.add(rel.id);
  return ids;
}

/** `.rels` parts keyed by the **source** part they describe, normalised. */
export function bySource(parts: readonly RelsPart[]): ReadonlyMap<string, RelsPart> {
  const map = new Map<string, RelsPart>();
  for (const part of parts) {
    map.set(part.source === '/' ? '/' : normalizePartName(part.source), part);
  }
  return map;
}
