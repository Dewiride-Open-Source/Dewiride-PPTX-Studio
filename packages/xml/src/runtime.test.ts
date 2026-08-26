import { describe, expect, it } from 'vitest';
import { NS, XML_SPACE_PRESERVE } from './index.js';

describe('@pptx-studio/xml', () => {
  it('exposes the OOXML namespace vocabulary', () => {
    expect(NS.mc).toBe('http://schemas.openxmlformats.org/markup-compatibility/2006');
    expect(NS.a).toBe('http://schemas.openxmlformats.org/drawingml/2006/main');
    expect(NS.p).toBe('http://schemas.openxmlformats.org/presentationml/2006/main');
    expect(XML_SPACE_PRESERVE).toBe('preserve');
  });

  // These two are the reason the core suite runs in Chromium rather than jsdom.
  // Under jsdom both assertions pass while `node:fs` still resolves, so a Node
  // API leak stays invisible until someone loads the package in a tab.
  describe('browser runtime guarantee', () => {
    it('runs with no Node globals in scope', () => {
      expect(typeof globalThis.window).toBe('object');
      expect('process' in globalThis).toBe(false);
      expect('Buffer' in globalThis).toBe(false);
    });

    it('cannot resolve a Node builtin', async () => {
      // Assembled at runtime and marked @vite-ignore so the bundler leaves it
      // alone: the point is to observe the *browser* failing to resolve it. A
      // static `import 'node:fs'` in package source fails earlier still, at
      // transform time, which is the enforcement we actually rely on.
      const nodeBuiltin = ['node', 'fs'].join(':');
      await expect(import(/* @vite-ignore */ nodeBuiltin)).rejects.toThrow();
    });
  });
});
