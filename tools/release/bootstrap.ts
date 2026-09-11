/**
 * Which packages a release would be creating on npm rather than updating.
 *
 * A trusted publisher is a setting on an npm package, and npm has no way to
 * attach one to a name that does not exist yet, so the first publish of a new
 * name cannot go out over OIDC. `changeset publish` publishes one package at a
 * time, so discovering that halfway through leaves some published and some not.
 * This asks first. ADR 0046.
 */

export interface Publishable {
  readonly name: string;
  readonly version: string;
}

/**
 * The packages a release would publish: public, and not held back by changesets.
 *
 * `ignore` is the only thing that keeps a public package out of a release, so
 * a name leaving that list is exactly the moment this check starts to matter.
 */
export function publishableIn(
  manifests: readonly { name: string; version: string; private?: boolean }[],
  ignored: readonly string[],
): readonly Publishable[] {
  const held = new Set(ignored);
  return manifests
    .filter((m) => m.private !== true && !held.has(m.name))
    .map((m) => ({ name: m.name, version: m.version }))
    .sort((a, b) => (a.name < b.name ? -1 : 1));
}

/** The names the registry has never heard of, in the order they were given. */
export function newcomers(
  packages: readonly Publishable[],
  published: ReadonlySet<string>,
): readonly string[] {
  return packages.filter((p) => !published.has(p.name)).map((p) => p.name);
}

/**
 * What to do about them, as the message the release fails with.
 *
 * Written out in full because the operator reading it is mid-release and the
 * steps are not guessable: npm/cli#8544 is open, so there is no way to
 * pre-register, and the token has to be granular and scope-wide.
 */
export function bootstrapInstructions(names: readonly string[]): string {
  const list = names.map((name) => `  - ${name}`).join('\n');
  return [
    `${String(names.length)} package(s) have never been published, and a trusted publisher`,
    'cannot be attached to a name npm does not have yet (npm/cli#8544):',
    list,
    '',
    'This release would have created them over OIDC, which fails, and changeset',
    'publish goes one package at a time - so some would have shipped and some',
    'would not. Nothing was published.',
    '',
    'To create them, once, by hand:',
    '  1. npmjs.com -> Access Tokens -> Generate -> Granular access token.',
    '     Scope @pptx-studio, "Read and write", expiry 1 day, bypass 2FA.',
    '  2. npm publish --provenance --access public, from each package directory,',
    '     with NODE_AUTH_TOKEN set to that token. Provenance with a token needs',
    '     only id-token: write; it does not need trusted publishing.',
    '  3. npmjs.com -> each new package -> Settings -> Trusted publisher ->',
    '     this repository, workflow release.yml.',
    '  4. Revoke the token.',
    '  5. Remove the name from "ignore" in .changeset/config.json.',
  ].join('\n');
}
