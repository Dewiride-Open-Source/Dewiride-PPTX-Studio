import { describe, expect, it } from 'vitest';
import { checkLayering, type Manifest, type PackageJsonLike, type RuleId } from './check.ts';
import { PACKAGES } from './layers.ts';

const CATALOG = { fflate: '^0.8.3', typescript: '6.0.3' };

/** A manifest that satisfies every hygiene rule, so tests vary one thing at a time. */
function manifest(short: string, overrides: Partial<PackageJsonLike> = {}): Manifest {
  return {
    dir: `packages/${short}`,
    json: {
      name: `@pptx-studio/${short}`,
      type: 'module',
      license: 'Apache-2.0',
      sideEffects: false,
      repository: { type: 'git', url: 'git+https://example.invalid/repo.git' },
      publishConfig: { access: 'public' },
      exports: { '.': { types: './dist/index.d.ts', default: './dist/index.js' } },
      ...overrides,
    },
  };
}

function rules(manifests: Manifest[]): RuleId[] {
  return checkLayering({ manifests, catalog: CATALOG }).map((v) => v.rule);
}

describe('layer table', () => {
  it('places every package the plan names, with react above the renderers', () => {
    expect(PACKAGES['opc']?.layer).toBe(0);
    expect(PACKAGES['xml']?.layer).toBe(0);
    expect(PACKAGES['model']!.layer).toBeGreaterThan(PACKAGES['geometry']!.layer);
    expect(PACKAGES['render-dom']!.layer).toBeGreaterThan(PACKAGES['model']!.layer);
    expect(PACKAGES['react']!.layer).toBeGreaterThan(PACKAGES['render-dom']!.layer);
    expect(PACKAGES['cli']!.layer).toBeGreaterThan(PACKAGES['react']!.layer);
  });

  it('marks cli as the only Node-runtime package', () => {
    const nodePackages = Object.entries(PACKAGES)
      .filter(([, spec]) => spec.runtime === 'node')
      .map(([name]) => name);
    expect(nodePackages).toEqual(['cli']);
  });
});

describe('checkLayering', () => {
  it('accepts a clean workspace', () => {
    expect(
      rules([
        manifest('xml'),
        manifest('opc', { dependencies: { fflate: 'catalog:' } }),
        manifest('model', {
          dependencies: {
            '@pptx-studio/opc': 'workspace:^',
            '@pptx-studio/xml': 'workspace:^',
          },
        }),
      ]),
    ).toEqual([]);
  });

  it('L001 flags a package directory missing from the layer table', () => {
    expect(rules([manifest('mystery-box')])).toContain('L001-unknown-package');
  });

  it('L002 flags a dependency on a higher layer', () => {
    // opc is layer 0; model is layer 2. This is the violation the whole file exists for.
    const found = rules([
      manifest('opc', { dependencies: { '@pptx-studio/model': 'workspace:^' } }),
    ]);
    expect(found).toContain('L002-upward-dependency');
  });

  it('L002 allows a dependency on a strictly lower layer', () => {
    expect(
      rules([manifest('model', { dependencies: { '@pptx-studio/xml': 'workspace:^' } })]),
    ).not.toContain('L002-upward-dependency');
  });

  it('L002 allows a same-layer dependency', () => {
    // render-dom building on render-svg is legitimate; only cycles are not.
    expect(
      rules([
        manifest('render-dom', { dependencies: { '@pptx-studio/render-svg': 'workspace:^' } }),
        manifest('render-svg'),
      ]),
    ).not.toContain('L002-upward-dependency');
  });

  it('L003 flags a cycle between same-layer packages', () => {
    const found = rules([
      manifest('render-dom', { dependencies: { '@pptx-studio/render-svg': 'workspace:^' } }),
      manifest('render-svg', { dependencies: { '@pptx-studio/render-dom': 'workspace:^' } }),
    ]);
    expect(found).toContain('L003-dependency-cycle');
  });

  it('L004 flags React anywhere but the react package', () => {
    expect(rules([manifest('model', { dependencies: { react: '^19.2.8' } })])).toContain(
      'L004-react-leak',
    );
    expect(
      rules([manifest('render-dom', { devDependencies: { '@types/react': '^19.0.0' } })]),
    ).toContain('L004-react-leak');
  });

  it('L004 permits React inside the react package', () => {
    expect(rules([manifest('react', { peerDependencies: { react: '^19.0.0' } })])).not.toContain(
      'L004-react-leak',
    );
  });

  it('L005 flags a node condition in a browser package exports map', () => {
    const found = rules([
      manifest('opc', {
        exports: {
          '.': {
            types: './dist/index.d.ts',
            node: './dist/index.node.js',
            default: './dist/index.js',
          },
        },
      }),
    ]);
    expect(found).toContain('L005-node-condition');
  });

  it('L005 permits a node condition in the cli package', () => {
    expect(rules([manifest('cli', { exports: { '.': { node: './dist/cli.js' } } })])).not.toContain(
      'L005-node-condition',
    );
  });

  it('L006 flags an internal dependency that is not on the workspace protocol', () => {
    expect(
      rules([manifest('model', { dependencies: { '@pptx-studio/xml': '^0.1.0' } })]),
    ).toContain('L006-workspace-protocol');
  });

  it('L007 flags manifest fields the publish pipeline depends on', () => {
    expect(rules([manifest('xml', { sideEffects: true })])).toContain('L007-manifest-field');
    expect(rules([manifest('xml', { license: 'MIT' })])).toContain('L007-manifest-field');
    expect(rules([manifest('xml', { publishConfig: {} })])).toContain('L007-manifest-field');
    expect(rules([manifest('xml', { repository: undefined })])).toContain('L007-manifest-field');
  });

  it('L008 flags a literal version for a dependency the catalog already defines', () => {
    expect(rules([manifest('opc', { dependencies: { fflate: '^0.8.3' } })])).toContain(
      'L008-uncatalogued-version',
    );
    expect(rules([manifest('opc', { dependencies: { fflate: 'catalog:' } })])).not.toContain(
      'L008-uncatalogued-version',
    );
  });

  it('L009 flags a package whose name disagrees with its directory', () => {
    expect(rules([manifest('xml', { name: '@pptx-studio/xmlz' })])).toContain('L009-package-name');
  });

  it('ignores apps/, which are private and not layered', () => {
    expect(
      checkLayering({
        manifests: [
          { dir: 'apps/studio', json: { name: 'studio', dependencies: { react: '^19' } } },
        ],
        catalog: CATALOG,
      }),
    ).toEqual([]);
  });
});
