/**
 * The check that stops a release creating a package it cannot create.
 *
 * `@pptx-studio/render-dom` is the live case: it sits in changesets' `ignore`
 * at 0.0.0, so today it is not publishable and this says nothing. The moment
 * somebody lifts the ignore, it has to speak. ADR 0046.
 */

import { readFileSync } from 'node:fs';

import { describe, expect, it } from 'vitest';

import { repoPath } from '../repo/root.ts';
import { bootstrapInstructions, newcomers, publishableIn } from './bootstrap.ts';

const MANIFESTS = [
  { name: '@pptx-studio/cli', version: '0.2.0' },
  { name: '@pptx-studio/render-dom', version: '0.0.0' },
  { name: '@pptx-studio/studio', version: '0.0.0', private: true },
];

describe('what a release would publish', () => {
  it('is the public packages changesets is not holding back', () => {
    expect(publishableIn(MANIFESTS, ['@pptx-studio/render-dom'])).toEqual([
      { name: '@pptx-studio/cli', version: '0.2.0' },
    ]);
  });

  it('includes a package the moment it leaves the ignore list', () => {
    expect(publishableIn(MANIFESTS, []).map((p) => p.name)).toEqual([
      '@pptx-studio/cli',
      '@pptx-studio/render-dom',
    ]);
  });

  it('never includes a private package, ignored or not', () => {
    const names = publishableIn(MANIFESTS, []).map((p) => p.name);
    expect(names).not.toContain('@pptx-studio/studio');
  });
});

describe('newcomers', () => {
  it('is empty when the registry already has every name', () => {
    const packages = publishableIn(MANIFESTS, []);
    const published = new Set(packages.map((p) => p.name));
    expect(newcomers(packages, published)).toEqual([]);
  });

  it('names the package the registry has never heard of', () => {
    const packages = publishableIn(MANIFESTS, []);
    expect(newcomers(packages, new Set(['@pptx-studio/cli']))).toEqual(['@pptx-studio/render-dom']);
  });

  it('says nothing while the package is still held back', () => {
    const packages = publishableIn(MANIFESTS, ['@pptx-studio/render-dom']);
    expect(newcomers(packages, new Set(['@pptx-studio/cli']))).toEqual([]);
  });
});

describe('the message', () => {
  const text = bootstrapInstructions(['@pptx-studio/render-dom']);

  it('names the package and says nothing was published', () => {
    expect(text).toContain('@pptx-studio/render-dom');
    expect(text).toContain('Nothing was published.');
  });

  it('carries the steps, because they are not guessable mid-release', () => {
    expect(text).toContain('Granular access token');
    expect(text).toContain('Trusted publisher');
    expect(text).toContain('Revoke the token');
    expect(text).toContain('.changeset/config.json');
  });
});

describe('the repository as it stands', () => {
  it('holds render-dom back, so no release can try to create it', () => {
    const config = JSON.parse(readFileSync(repoPath('.changeset/config.json'), 'utf8')) as {
      ignore: string[];
    };
    // The guard above is the safety net; this is the thing it is a net for.
    expect(config.ignore).toContain('@pptx-studio/render-dom');
  });
});
