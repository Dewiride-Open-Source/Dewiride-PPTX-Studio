import { PACKAGES, REACT_OWNER, REQUIRED_FIELDS, SCOPE, type PackageSpec } from './layers.ts';

export interface PackageJsonLike {
  name?: string;
  private?: boolean;
  type?: string;
  license?: string;
  sideEffects?: unknown;
  dependencies?: Record<string, string>;
  devDependencies?: Record<string, string>;
  peerDependencies?: Record<string, string>;
  optionalDependencies?: Record<string, string>;
  exports?: unknown;
  publishConfig?: { access?: string };
  repository?: unknown;
}

export interface Manifest {
  /** Workspace-relative directory, e.g. `packages/opc`. */
  readonly dir: string;
  readonly json: PackageJsonLike;
}

export type RuleId =
  | 'L001-unknown-package'
  | 'L002-upward-dependency'
  | 'L003-dependency-cycle'
  | 'L004-react-leak'
  | 'L005-node-condition'
  | 'L006-workspace-protocol'
  | 'L007-manifest-field'
  | 'L008-uncatalogued-version'
  | 'L009-package-name';

export interface Violation {
  readonly rule: RuleId;
  /** The offending package's directory, so the message is clickable. */
  readonly where: string;
  readonly message: string;
}

/** Runtime dependency fields. Kept separate from devDependencies deliberately. */
const RUNTIME_DEP_FIELDS = ['dependencies', 'peerDependencies', 'optionalDependencies'] as const;

function runtimeDeps(json: PackageJsonLike): Record<string, string> {
  const out: Record<string, string> = {};
  for (const field of RUNTIME_DEP_FIELDS) {
    Object.assign(out, json[field] ?? {});
  }
  return out;
}

function allDeps(json: PackageJsonLike): Record<string, string> {
  return { ...runtimeDeps(json), ...(json.devDependencies ?? {}) };
}

/** `@pptx-studio/render-svg` -> `render-svg`; anything else -> undefined. */
function shortName(dep: string): string | undefined {
  const prefix = SCOPE + '/';
  return dep.startsWith(prefix) ? dep.slice(prefix.length) : undefined;
}

/** Depth-first search for a `"node"` key anywhere inside an exports map. */
function hasNodeCondition(exports: unknown): boolean {
  if (exports === null || typeof exports !== 'object') return false;
  if (Array.isArray(exports)) return exports.some(hasNodeCondition);
  for (const [key, value] of Object.entries(exports)) {
    if (key === 'node' || key === 'node-addons') return true;
    if (hasNodeCondition(value)) return true;
  }
  return false;
}

/** Iterative-colour DFS. Returns one representative path per cycle found. */
function findCycles(graph: ReadonlyMap<string, readonly string[]>): string[][] {
  const WHITE = 0;
  const GREY = 1;
  const BLACK = 2;
  const colour = new Map<string, number>();
  const cycles: string[][] = [];

  const visit = (node: string, path: string[]): void => {
    colour.set(node, GREY);
    path.push(node);
    for (const next of graph.get(node) ?? []) {
      if (!graph.has(next)) continue;
      const state = colour.get(next) ?? WHITE;
      if (state === GREY) {
        cycles.push([...path.slice(path.indexOf(next)), next]);
      } else if (state === WHITE) {
        visit(next, path);
      }
    }
    path.pop();
    colour.set(node, BLACK);
  };

  for (const node of graph.keys()) {
    if ((colour.get(node) ?? WHITE) === WHITE) visit(node, []);
  }
  return cycles;
}

export interface CheckInput {
  readonly manifests: readonly Manifest[];
  /** The `catalog:` block from pnpm-workspace.yaml. */
  readonly catalog: Readonly<Record<string, string>>;
}

/**
 * Pure. Given every workspace manifest plus the catalog, return every layering
 * and manifest-hygiene violation. No filesystem access and no process exit, so
 * it can be exercised from tests with synthetic manifests.
 */
export function checkLayering({ manifests, catalog }: CheckInput): Violation[] {
  const violations: Violation[] = [];
  const report = (rule: RuleId, where: string, message: string): void => {
    violations.push({ rule, where, message });
  };

  const packageManifests = manifests.filter((m) => m.dir.startsWith('packages/'));
  const graph = new Map<string, readonly string[]>();

  for (const manifest of packageManifests) {
    const dir = manifest.dir;
    const short = dir.slice('packages/'.length);
    const json = manifest.json;
    const spec: PackageSpec | undefined = PACKAGES[short];

    // L001 - every directory under packages/ is in the layer table.
    if (!spec) {
      report(
        'L001-unknown-package',
        dir,
        'packages/' +
          short +
          ' has no entry in tools/repo/layering/layers.ts. Add one (with its layer and runtime) before ' +
          'adding code, so the dependency direction is declared rather than discovered.',
      );
      continue;
    }

    // L009 - directory name and package name agree.
    const expectedName = SCOPE + '/' + short;
    if (json.name !== expectedName) {
      report(
        'L009-package-name',
        dir,
        'package name is ' +
          JSON.stringify(json.name) +
          ', expected ' +
          JSON.stringify(expectedName) +
          '. The layer table is keyed by directory; a mismatch means the guard silently checks the ' +
          'wrong package.',
      );
    }

    // L007 - manifest fields the publish pipeline depends on.
    for (const [field, expected] of Object.entries(REQUIRED_FIELDS)) {
      const actual = json[field as keyof PackageJsonLike];
      if (actual !== expected) {
        report(
          'L007-manifest-field',
          dir,
          JSON.stringify(field) +
            ' is ' +
            JSON.stringify(actual) +
            ', expected ' +
            JSON.stringify(expected) +
            '.',
        );
      }
    }
    if (json.publishConfig?.access !== 'public') {
      report(
        'L007-manifest-field',
        dir,
        'scoped packages default to restricted; set "publishConfig": { "access": "public" } or the ' +
          'first publish fails.',
      );
    }
    if (json.repository === undefined) {
      report(
        'L007-manifest-field',
        dir,
        '"repository" is required for npm provenance attestation.',
      );
    }

    // L005 - a browser package must not advertise a Node entry point.
    if (spec.runtime === 'browser' && hasNodeCondition(json.exports)) {
      report(
        'L005-node-condition',
        dir,
        'exports map declares a "node" condition. ' +
          expectedName +
          ' is browser-only; a node condition lets a Node-specific build exist, which is how core ' +
          'packages quietly grow a server half.',
      );
    }

    const deps = allDeps(json);
    const workspaceDeps: string[] = [];

    for (const [dep, range] of Object.entries(deps)) {
      const depShort = shortName(dep);

      if (depShort !== undefined) {
        workspaceDeps.push(depShort);

        // L006 - internal deps go through the workspace protocol.
        if (!range.startsWith('workspace:')) {
          report(
            'L006-workspace-protocol',
            dir,
            'depends on ' +
              dep +
              ' with ' +
              JSON.stringify(range) +
              '; use "workspace:^" so local development never resolves a published copy.',
          );
        }

        const depSpec = PACKAGES[depShort];
        if (!depSpec) {
          report(
            'L001-unknown-package',
            dir,
            'depends on ' + dep + ', which is not in the layer table.',
          );
          continue;
        }

        // L002 - no upward dependencies.
        if (depSpec.layer > spec.layer) {
          report(
            'L002-upward-dependency',
            dir,
            expectedName +
              ' (layer ' +
              spec.layer +
              ') depends on ' +
              dep +
              ' (layer ' +
              depSpec.layer +
              '). Dependencies run downward only.',
          );
        }
        continue;
      }

      // L004 - React is confined to one package.
      const touchesReact = dep === 'react' || dep === 'react-dom' || dep.startsWith('@types/react');
      if (touchesReact && short !== REACT_OWNER) {
        report(
          'L004-react-leak',
          dir,
          'depends on ' +
            dep +
            '. Only ' +
            SCOPE +
            '/' +
            REACT_OWNER +
            ' and apps/ may reference React - the core is framework-free by contract, and that ' +
            'contract is what makes a Vue or Svelte binding possible later.',
        );
      }

      // L008 - shared external versions come from the catalog.
      if (catalog[dep] !== undefined && !range.startsWith('catalog:')) {
        report(
          'L008-uncatalogued-version',
          dir,
          'pins ' +
            dep +
            '@' +
            range +
            ' but the workspace catalog already defines it. Use "catalog:" so one bump moves every ' +
            'package at once.',
        );
      }
    }

    graph.set(short, workspaceDeps);
  }

  // L003 - the graph must stay acyclic. Same-layer edges are legal; cycles are not.
  for (const cycle of findCycles(graph)) {
    report(
      'L003-dependency-cycle',
      'packages/' + cycle[0],
      'dependency cycle: ' + cycle.map((c) => SCOPE + '/' + c).join(' -> '),
    );
  }

  return violations;
}
