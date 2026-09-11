/**
 * The preflight refuses on npm's answer, not on a guess about the answer.
 *
 * Its network half runs only inside a workflow that can mint an ID token, so
 * what is provable here is the addressing and the decision. ADR 0048.
 */

import { describe, expect, it } from 'vitest';

import type { Publishable } from './bootstrap.ts';
import {
  type Attempt,
  audienceFor,
  escapedName,
  exchangeUrl,
  idTokenUrl,
  missingPublisherInstructions,
  pending,
  refusals,
} from './oidc.ts';

const REGISTRY = 'https://registry.npmjs.org';

const at = (name: string, version: string): Publishable => ({ name, version });

describe('addressing the exchange', () => {
  it('takes the audience from the registry host', () => {
    expect(audienceFor(REGISTRY)).toBe('npm:registry.npmjs.org');
    expect(audienceFor('https://npm.pkg.github.com')).toBe('npm:npm.pkg.github.com');
  });

  it('escapes the scope separator, which a path segment cannot hold', () => {
    expect(escapedName('@pptx-studio/xml')).toBe('@pptx-studio%2fxml');
    expect(escapedName('fflate')).toBe('fflate');
  });

  it('addresses one package, because publish rights are per package', () => {
    expect(exchangeUrl(REGISTRY, '@pptx-studio/xml')).toBe(
      `${REGISTRY}/-/npm/v1/oidc/token/exchange/package/@pptx-studio%2fxml`,
    );
  });

  it('appends the audience to an issuer that already has a query string', () => {
    const url = idTokenUrl('https://pipelines.example/token?api-version=2.0', 'npm:x.example');
    expect(url).toBe('https://pipelines.example/token?api-version=2.0&audience=npm%3Ax.example');
  });
});

describe('what a release would send', () => {
  const packages = [at('@pptx-studio/xml', '0.1.1'), at('@pptx-studio/opc', '0.1.0')];

  it('is the versions npm does not already hold', () => {
    const onNpm = new Set(['@pptx-studio/opc@0.1.0']);
    expect(pending(packages, onNpm).map((p) => p.name)).toEqual(['@pptx-studio/xml']);
  });

  it('matches on the version, not the name', () => {
    const onNpm = new Set(['@pptx-studio/xml@0.1.0']);
    expect(pending(packages, onNpm).map((p) => p.name)).toEqual([
      '@pptx-studio/xml',
      '@pptx-studio/opc',
    ]);
  });

  it('is empty when the registry already holds every version in the tree', () => {
    const onNpm = new Set(['@pptx-studio/xml@0.1.1', '@pptx-studio/opc@0.1.0']);
    expect(pending(packages, onNpm)).toEqual([]);
  });
});

describe('what blocks a release', () => {
  const attempts: readonly Attempt[] = [
    { name: '@pptx-studio/xml', allowed: false, detail: '404, no trusted publisher' },
    { name: '@pptx-studio/opc', allowed: false, detail: '403, forbidden' },
    { name: '@pptx-studio/cli', allowed: true, detail: 'may publish from this workflow' },
  ];

  it('is a refusal for a package this release would publish', () => {
    const blocked = refusals(attempts, new Set(['@pptx-studio/xml']));
    expect(blocked.map((a) => a.name)).toEqual(['@pptx-studio/xml']);
  });

  it('is not a refusal for a package this release leaves alone', () => {
    expect(refusals(attempts, new Set(['@pptx-studio/cli']))).toEqual([]);
  });

  it('names every refusal at once, so one dispatch finds them all', () => {
    const names = new Set(['@pptx-studio/xml', '@pptx-studio/opc', '@pptx-studio/cli']);
    expect(refusals(attempts, names).map((a) => a.name)).toEqual([
      '@pptx-studio/xml',
      '@pptx-studio/opc',
    ]);
  });

  it('says nothing was published, because a half release is the hazard', () => {
    const message = missingPublisherInstructions([attempts[0] as Attempt]);
    expect(message).toContain('Nothing was published.');
    expect(message).toContain('@pptx-studio/xml: 404, no trusted publisher');
    expect(message).toContain('release.yml');
  });
});
