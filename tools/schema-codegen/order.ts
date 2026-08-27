/**
 * Turning `xsd:sequence` into something an editor can insert into.
 *
 * ## What the table has to answer
 *
 * Not "is this document valid" - that is sub-phase 1.2's job. The single
 * question here is: *given a parent element and a child element name, where in
 * the parent's existing children does the new one go?*
 *
 * The answer is a **rank**. Every child name a parent admits gets an integer,
 * and in any valid instance the children appear in non-decreasing rank order.
 * Two names sharing a rank may appear in either order - which is not a rounding
 * of the schema but a fact about it: `p:spTree` admits `sp`, `grpSp`,
 * `graphicFrame`, `cxnSp`, `pic` and `contentPart` under one
 * `maxOccurs="unbounded"` choice, and *that freedom is z-order*. Collapsing it
 * to a strict list would make inserting a shape reorder the slide.
 *
 * ## The three rules that produce a rank
 *
 * 1. A `sequence` advances the rank once per particle.
 * 2. A `choice` gives every branch the same starting rank, because at most one
 *    branch is ever instantiated, and takes the highest rank any branch
 *    reached. Ranks within a branch still order that branch's own elements.
 * 3. **A compositor that repeats freezes everything beneath it to one rank.**
 *    If `maxOccurs > 1`, a second repetition can put a first-particle element
 *    after a last-particle one, so no order between them survives. This is the
 *    rule that gets `a:r`, `a:br` and `a:fld` right inside `a:p` while keeping
 *    `a:pPr` before them and `a:endParaRPr` after.
 *
 * ## And the fourth thing, which is not a rank
 *
 * `xsd:any` marks a type **open**. There are exactly three in PresentationML
 * and DrawingML - `CT_Extension`, `CT_OfficeArtExtension` and
 * `CT_GraphicalObjectData` - and all three are places the plan forbids us to
 * rebuild: the inside of an `extLst` entry, and the inside of a
 * `graphicFrame`'s payload where a chart or a diagram lives. An open type gets
 * no rank table at all, so `insertInOrder` refuses it and the caller has to say
 * explicitly that it means to append.
 */

import type { XsdParticle, XsdSchema } from './read-xsd.ts';

/** `{namespace}local`, the only key that is prefix-independent. */
export function key(ns: string, local: string): string {
  return `{${ns}}${local}`;
}

export interface Universe {
  readonly complexTypes: ReadonlyMap<string, XsdParticle | undefined>;
  readonly groups: ReadonlyMap<string, XsdParticle>;
  /** Global element declarations: element key -> type key. */
  readonly globalElements: ReadonlyMap<string, string>;
}

export function universeOf(schemas: readonly XsdSchema[]): Universe {
  const complexTypes = new Map<string, XsdParticle | undefined>();
  const groups = new Map<string, XsdParticle>();
  const globalElements = new Map<string, string>();
  for (const schema of schemas) {
    for (const [name, content] of schema.complexTypes) {
      complexTypes.set(key(schema.targetNamespace, name), content);
    }
    for (const [name, content] of schema.groups) {
      groups.set(key(schema.targetNamespace, name), content);
    }
    for (const [local, type] of schema.globalElements) {
      globalElements.set(key(schema.targetNamespace, local), key(type.ns, type.name));
    }
  }
  return { complexTypes, groups, globalElements };
}

/** The content model of one complex type, flattened. */
export interface TypeOrder {
  /** Child element key -> rank. Empty for a type with no element children. */
  readonly ranks: ReadonlyMap<string, number>;
  /** The type admits `xsd:any`, so its content is not ours to order. */
  readonly open: boolean;
  /** Element keys the walk assigned two different ranks. Expected to be empty. */
  readonly collisions: readonly string[];
}

/** Every (element key -> type key) pair seen at any use site, for ambiguity analysis. */
export interface Binding {
  readonly elementKey: string;
  readonly typeKey: string;
  /** The type whose content model declared it, for the diagnostic. */
  readonly declaredIn: string;
}

export interface FlattenResult {
  readonly orders: ReadonlyMap<string, TypeOrder>;
  readonly bindings: readonly Binding[];
}

/**
 * Flatten every complex type in the universe, and collect the element-to-type
 * bindings on the way.
 */
export function flatten(universe: Universe): FlattenResult {
  const orders = new Map<string, TypeOrder>();
  const bindings: Binding[] = [];

  for (const [typeKey, content] of universe.complexTypes) {
    const ranks = new Map<string, number>();
    const collisions: string[] = [];
    let open = false;
    const visiting = new Set<string>();

    const elementKeyOf = (p: Extract<XsdParticle, { kind: 'element' }>): string =>
      key(p.ns, p.local);

    const typeKeyOf = (p: Extract<XsdParticle, { kind: 'element' }>): string => {
      // A `ref=` to a global declaration carries no type of its own.
      if (p.typeName === '') {
        const resolved = universe.globalElements.get(elementKeyOf(p));
        if (resolved === undefined) {
          throw new Error(
            `${typeKey}: element ref ${elementKeyOf(p)} resolves to no global element`,
          );
        }
        return resolved;
      }
      return key(p.typeNs, p.typeName);
    };

    const place = (p: Extract<XsdParticle, { kind: 'element' }>, rank: number): void => {
      const k = elementKeyOf(p);
      const seen = ranks.get(k);
      if (seen !== undefined && seen !== rank) collisions.push(k);
      else ranks.set(k, rank);
      bindings.push({ elementKey: k, typeKey: typeKeyOf(p), declaredIn: typeKey });
    };

    const groupOf = (p: Extract<XsdParticle, { kind: 'group' }>): XsdParticle => {
      const k = key(p.ns, p.name);
      const g = universe.groups.get(k);
      if (g === undefined) throw new Error(`${typeKey}: group ${k} is not defined`);
      if (visiting.has(k)) throw new Error(`${typeKey}: group ${k} is recursive`);
      return g;
    };

    /** Everything beneath a repeating compositor shares one rank. */
    const frozen = (p: XsdParticle, rank: number): void => {
      switch (p.kind) {
        case 'any':
          open = true;
          return;
        case 'element':
          place(p, rank);
          return;
        case 'group': {
          const k = key(p.ns, p.name);
          visiting.add(k);
          frozen(groupOf(p), rank);
          visiting.delete(k);
          return;
        }
        case 'sequence':
        case 'choice':
          for (const child of p.children) frozen(child, rank);
          return;
      }
    };

    /** Returns the rank the next sibling particle starts at. */
    const walk = (p: XsdParticle, base: number): number => {
      switch (p.kind) {
        case 'any':
          open = true;
          return base;
        case 'element':
          place(p, base);
          return base + 1;
        case 'group': {
          const k = key(p.ns, p.name);
          const g = groupOf(p);
          if (p.max > 1) {
            frozen(g, base);
            return base + 1;
          }
          visiting.add(k);
          const next = walk(g, base);
          visiting.delete(k);
          return next;
        }
        case 'sequence': {
          if (p.max > 1) {
            frozen(p, base);
            return base + 1;
          }
          let rank = base;
          for (const child of p.children) rank = walk(child, rank);
          return rank;
        }
        case 'choice': {
          if (p.max > 1) {
            frozen(p, base);
            return base + 1;
          }
          let highest = base;
          for (const child of p.children) highest = Math.max(highest, walk(child, base));
          return highest;
        }
      }
    };

    if (content !== undefined) walk(content, 0);
    orders.set(typeKey, { ranks, open, collisions });
  }

  return { orders, bindings };
}

/** A stable string for one type's content model, so two can be compared. */
export function orderSignature(order: TypeOrder): string {
  const parts = [...order.ranks].sort((a, b) => a[1] - b[1] || a[0].localeCompare(b[0]));
  return (order.open ? 'open ' : '') + parts.map(([k, r]) => `${String(r)}:${k}`).join(' ');
}

export interface Ambiguity {
  readonly elementKey: string;
  /** Type key -> the content-model signature it produces. More than one entry is the problem. */
  readonly variants: ReadonlyMap<string, string>;
  /** Which parent types declared each variant, for the diagnostic. */
  readonly declaredIn: ReadonlyMap<string, readonly string[]>;
}

/**
 * Element names used with more than one *distinguishable* content model.
 *
 * The table is keyed by parent element name, because that is what a caller
 * holding an `XElement` has. The schema keys by complex type, and the two are
 * not the same map: `xfrm` is `CT_Transform2D` under `a:spPr` and
 * `CT_GroupTransform2D` under `a:grpSpPr`. Most such splits are invisible here
 * because the two types produce identical child orders, or because neither has
 * element children at all - `a:ext` is `CT_PositiveSize2D` in one place and
 * `CT_OfficeArtExtension` in another, and only the second has any content.
 *
 * What this function finds is the residue: names where the choice of type
 * genuinely changes where a child belongs. Those cannot be keyed by name alone.
 */
export function ambiguities(universe: Universe, flattened: FlattenResult): Ambiguity[] {
  const byElement = new Map<string, Map<string, string>>();
  const declaredIn = new Map<string, Map<string, string[]>>();

  const record = (elementKey: string, typeKey: string, from: string): void => {
    const order = flattened.orders.get(typeKey);
    // A simple type, or one from a schema we did not read. Neither can hold
    // element children, so neither can disagree about their order.
    const signature = order === undefined ? '' : orderSignature(order);
    let variants = byElement.get(elementKey);
    if (variants === undefined) {
      variants = new Map();
      byElement.set(elementKey, variants);
    }
    variants.set(typeKey, signature);
    let sites = declaredIn.get(elementKey);
    if (sites === undefined) {
      sites = new Map();
      declaredIn.set(elementKey, sites);
    }
    const list = sites.get(typeKey) ?? [];
    if (!list.includes(from)) list.push(from);
    sites.set(typeKey, list);
  };

  for (const binding of flattened.bindings)
    record(binding.elementKey, binding.typeKey, binding.declaredIn);
  for (const [elementKey, typeKey] of universe.globalElements) {
    record(elementKey, typeKey, '(global)');
  }

  const found: Ambiguity[] = [];
  for (const [elementKey, variants] of byElement) {
    const signatures = new Set([...variants.values()].filter((s) => s !== ''));
    if (signatures.size > 1) {
      found.push({ elementKey, variants, declaredIn: declaredIn.get(elementKey)! });
    }
  }
  return found.sort((a, b) => a.elementKey.localeCompare(b.elementKey));
}

/**
 * The table as it will be generated: parent element key -> child ranks.
 *
 * Built only from names whose content model is unambiguous. An open type
 * contributes nothing, which is the point - a caller must not be able to insert
 * into an `extLst` entry by accident.
 */
export function tableOf(
  universe: Universe,
  flattened: FlattenResult,
  ambiguous: ReadonlySet<string>,
): Map<string, ReadonlyMap<string, number>> {
  const table = new Map<string, ReadonlyMap<string, number>>();
  const seen = new Set<string>();

  const consider = (elementKey: string, typeKey: string): void => {
    if (ambiguous.has(elementKey) || seen.has(elementKey)) return;
    const order = flattened.orders.get(typeKey);
    if (order === undefined || order.open || order.ranks.size === 0) return;
    seen.add(elementKey);
    table.set(elementKey, order.ranks);
  };

  for (const binding of flattened.bindings) consider(binding.elementKey, binding.typeKey);
  for (const [elementKey, typeKey] of universe.globalElements) consider(elementKey, typeKey);
  return table;
}

export interface ContextTable {
  /** `<grandparent key>|<parent key>` -> the parent's child ranks. */
  readonly entries: ReadonlyMap<string, ReadonlyMap<string, number>>;
  /** Pairs the grandparent still fails to separate. Expected to be empty. */
  readonly unresolved: readonly string[];
}

/**
 * The three names the parent alone cannot place, resolved by one more level.
 *
 * `a:xfrm` is the one that matters. Under `p:spPr` it is `CT_Transform2D` -
 * `a:off` then `a:ext` - and under `p:grpSpPr` it is `CT_GroupTransform2D`,
 * which adds `a:chOff` and `a:chExt` after them. Moving a shape writes into the
 * first of those on every drag, so dropping the name for being ambiguous would
 * have cost us the single most-used insertion point in the project. The other
 * two are `a:path`, which is a gradient's shade path under `a:gradFill` and a
 * geometry subpath under `a:pathLst`, and `p:to`, which is a colour under
 * `p:animClr` and an animation variant under `p:set`.
 *
 * One extra level is enough for all three, and this returns the pairs where it
 * would not be, so that "enough" is a measurement rather than a hope.
 */
export function contextTableOf(
  universe: Universe,
  flattened: FlattenResult,
  ambiguous: ReadonlySet<string>,
): ContextTable {
  // Which element names carry which complex type - the reverse of `bindings`.
  const carriers = new Map<string, Set<string>>();
  const carry = (typeKey: string, elementKey: string): void => {
    const set = carriers.get(typeKey) ?? new Set<string>();
    set.add(elementKey);
    carriers.set(typeKey, set);
  };
  for (const binding of flattened.bindings) carry(binding.typeKey, binding.elementKey);
  for (const [elementKey, typeKey] of universe.globalElements) carry(typeKey, elementKey);

  const entries = new Map<string, ReadonlyMap<string, number>>();
  const chosen = new Map<string, string>();
  const unresolved: string[] = [];

  for (const binding of flattened.bindings) {
    if (!ambiguous.has(binding.elementKey)) continue;
    const order = flattened.orders.get(binding.typeKey);
    if (order === undefined || order.open || order.ranks.size === 0) continue;
    for (const grandparent of carriers.get(binding.declaredIn) ?? []) {
      const pair = `${grandparent}|${binding.elementKey}`;
      const already = chosen.get(pair);
      if (already !== undefined && already !== binding.typeKey) {
        unresolved.push(`${pair}: both ${already} and ${binding.typeKey}`);
        continue;
      }
      chosen.set(pair, binding.typeKey);
      entries.set(pair, order.ranks);
    }
  }

  return { entries, unresolved };
}
