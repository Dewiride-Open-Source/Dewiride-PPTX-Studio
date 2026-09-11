/**
 * The scratch consumer a release candidate is proved against, as pure functions.
 *
 * `publint` and `attw` read a tarball's manifest; they never resolve a specifier
 * at run time and never execute a line of the package. Nothing else installed
 * the packages either - `apps/studio` is private and resolves `workspace:^`
 * through the link farm - so 0.1.0 reached npm having never been installed
 * anywhere. ADR 0046.
 */

/** A tarball `pnpm pack --json` produced, named by the package inside it. */
export interface Packed {
  readonly name: string;
  readonly version: string;
  /** Absolute path to the `.tgz`. */
  readonly filename: string;
}

/**
 * The tarballs out of `pnpm -r pack --json`.
 *
 * Each entry carries a nested `files` array, so this reads brace depth rather
 * than matching a flat object, and takes the filename from the report instead
 * of rebuilding it - the default name flattens the scope.
 */
function jsonValues(raw: string): unknown[] {
  const values: unknown[] = [];
  let depth = 0;
  let start = -1;
  let inString = false;
  let escaped = false;
  for (let i = 0; i < raw.length; i++) {
    const c = raw[i];
    if (inString) {
      if (escaped) escaped = false;
      else if (c === '\\') escaped = true;
      else if (c === '"') inString = false;
    } else if (c === '"') {
      inString = true;
    } else if (c === '{' || c === '[') {
      if (depth === 0) start = i;
      depth += 1;
    } else if ((c === '}' || c === ']') && --depth === 0 && start >= 0) {
      values.push(JSON.parse(raw.slice(start, i + 1)));
      start = -1;
    }
  }
  return values;
}

export function packedFrom(raw: string): readonly Packed[] {
  const values = jsonValues(raw);
  const out: Packed[] = [];
  for (const value of values.flatMap((v) => (Array.isArray(v) ? (v as unknown[]) : [v]))) {
    if (typeof value !== 'object' || value === null) continue;
    const { name, version, filename } = value as Record<string, unknown>;
    if (typeof name !== 'string' || typeof version !== 'string' || typeof filename !== 'string') {
      continue;
    }
    out.push({ name, version, filename });
  }
  return out;
}

export interface ScratchManifest {
  readonly name: string;
  readonly version: string;
  readonly private: true;
  readonly type: 'module';
  /** The consumer's own, so the gate runs the commands the example declares. */
  readonly scripts: Readonly<Record<string, string>>;
  readonly dependencies: Readonly<Record<string, string>>;
  readonly overrides: Readonly<Record<string, string>>;
}

/**
 * `file:` needs forward slashes on every platform, including Windows.
 *
 * A backslash in a `file:` specifier is read as an escape by npm's parser.
 */
export function fileSpec(filename: string): string {
  return `file:${filename.replace(/\\/g, '/')}`;
}

/**
 * The scratch project that installs the candidates and nothing else.
 *
 * `dependencies` and `overrides` carry the same specifier for every package:
 * a direct dependency pins only itself, and the edges between the packages are
 * plain semver ranges once `pnpm pack` has rewritten `workspace:^`. Without the
 * overrides, an unbumped package's `^0.1.0` is satisfied by the copy already on
 * the registry and the gate silently tests what is published. npm refuses an
 * override that differs from the direct dependency, so the two are one string.
 */
export function scratchManifest(
  packed: readonly Packed[],
  extra: Readonly<Record<string, string>>,
  scripts: Readonly<Record<string, string>> = {},
): ScratchManifest {
  const pinned: Record<string, string> = {};
  for (const entry of [...packed].sort((a, b) => (a.name < b.name ? -1 : 1))) {
    pinned[entry.name] = fileSpec(entry.filename);
  }
  return {
    name: 'pptx-studio-release-candidate',
    version: '0.0.0',
    private: true,
    type: 'module',
    scripts,
    dependencies: { ...pinned, ...extra },
    overrides: pinned,
  };
}

/** One `package-lock.json` entry, of the fields provenance is read from. */
export interface LockEntry {
  readonly name?: string;
  readonly resolved?: string;
  readonly integrity?: string;
  readonly link?: boolean;
}

export interface Lockfile {
  readonly packages?: Readonly<Record<string, LockEntry>>;
}

/**
 * Every way the install could have reached something other than a candidate.
 *
 * A registry fallback is the failure this exists to catch, and it is silent: it
 * resolves, it installs, and the suite passes against the previous release.
 */
export function provenanceFailures(
  lock: Lockfile,
  packed: readonly Packed[],
  scope: string,
  integrityOf: (filename: string) => string,
): readonly string[] {
  const failures: string[] = [];
  const wanted = new Map(packed.map((entry) => [entry.name, entry]));
  const seen = new Set<string>();

  for (const [path, entry] of Object.entries(lock.packages ?? {})) {
    const name = entry.name ?? path.split('node_modules/').pop() ?? '';
    if (!name.startsWith(scope)) continue;
    seen.add(name);
    if (entry.link === true) {
      failures.push(`${path}: a symlink, so nothing was installed from a tarball`);
      continue;
    }
    const resolved = entry.resolved ?? '';
    if (!resolved.startsWith('file:')) {
      failures.push(`${path}: resolved from ${resolved === '' ? 'nowhere recorded' : resolved}`);
      continue;
    }
    const candidate = wanted.get(name);
    if (candidate === undefined) {
      failures.push(`${path}: ${name} is not a package this run packed`);
      continue;
    }
    const expected = integrityOf(candidate.filename);
    if (entry.integrity !== expected) {
      failures.push(
        `${path}: integrity ${String(entry.integrity)} is not the tarball's ${expected}`,
      );
    }
  }

  for (const name of wanted.keys()) {
    if (!seen.has(name)) failures.push(`${name} was packed but never installed`);
  }
  return failures;
}
