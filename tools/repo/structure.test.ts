import { describe, expect, it } from 'vitest';

import { checkStructure, MAX_SOURCE_FILES } from './structure.ts';

const many = (dir: string, n: number, ext = 'ts'): string[] =>
  Array.from({ length: n }, (_, i) => `${dir}/file${String(i).padStart(2, '0')}.${ext}`);

describe('the directory-shape rule', () => {
  it('allows a directory holding exactly the limit', () => {
    expect(checkStructure(many('packages/x/src', MAX_SOURCE_FILES))).toEqual([]);
  });

  it('reports one over the limit, with the files that are there', () => {
    const problems = checkStructure(many('packages/x/src', MAX_SOURCE_FILES + 1));
    expect(problems).toHaveLength(1);
    expect(problems[0]?.directory).toBe('packages/x/src');
    expect(problems[0]?.count).toBe(MAX_SOURCE_FILES + 1);
    expect(problems[0]?.files).toHaveLength(MAX_SOURCE_FILES + 1);
  });

  it('counts each directory on its own, not the tree below it', () => {
    const paths = [
      ...many('packages/x/src/a', MAX_SOURCE_FILES),
      ...many('packages/x/src/b', MAX_SOURCE_FILES),
    ];
    expect(checkStructure(paths)).toEqual([]);
  });

  // The line the rule is drawn on: corpus/decks is indexed by a manifest.
  it('does not count data files', () => {
    const decks = Array.from({ length: 42 }, (_, i) => `corpus/decks/a${String(i)}-probe.pptx`);
    const fixtures = many('corpus/ground-truth', 40, 'json');
    expect(checkStructure([...decks, ...fixtures])).toEqual([]);
  });

  it('counts generated source, which costs the same to read past', () => {
    const problems = checkStructure(many('packages/x/src/presets', MAX_SOURCE_FILES + 3, 'gen.ts'));
    expect(problems).toHaveLength(1);
  });

  it('counts the shell scripts an experiment is driven by', () => {
    const problems = checkStructure([
      ...many('tools/e', MAX_SOURCE_FILES - 1),
      'tools/e/author.ps1',
      'tools/e/read.ps1',
    ]);
    expect(problems).toHaveLength(1);
    expect(problems[0]?.count).toBe(MAX_SOURCE_FILES + 1);
  });

  it('sorts the worst offender first', () => {
    const paths = [...many('a', MAX_SOURCE_FILES + 1), ...many('b', MAX_SOURCE_FILES + 9)];
    expect(checkStructure(paths).map((p) => p.directory)).toEqual(['b', 'a']);
  });

  it('handles a file at the repository root', () => {
    expect(checkStructure(['vitest.config.ts'])).toEqual([]);
  });

  it('refuses a limit that is not a positive whole number', () => {
    expect(() => checkStructure([], 0)).toThrow(/positive whole number/);
    expect(() => checkStructure([], 1.5)).toThrow(/positive whole number/);
  });
});
