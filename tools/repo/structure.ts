/** The directory-shape rule of CLAUDE.md hard rule 4, as a pure function. */

/** The most hand-authored source files one directory may hold. */
export const MAX_SOURCE_FILES = 12;

/** Data is not counted: a fixture directory indexed by a manifest is already structured. */
const SOURCE = /\.(ts|tsx|mjs|cjs|ps1)$/;

export interface StructureProblem {
  readonly directory: string;
  readonly count: number;
  readonly files: readonly string[];
}

/** Every directory over the limit, worst first. Takes `git ls-files` paths. */
export function checkStructure(
  paths: readonly string[],
  max: number = MAX_SOURCE_FILES,
): StructureProblem[] {
  if (!Number.isInteger(max) || max < 1) {
    throw new Error(`a limit of ${String(max)} is not a positive whole number of files`);
  }

  const byDirectory = new Map<string, string[]>();
  for (const path of paths) {
    if (!SOURCE.test(path)) continue;
    const cut = path.lastIndexOf('/');
    const directory = cut < 0 ? '.' : path.slice(0, cut);
    const bucket = byDirectory.get(directory);
    if (bucket === undefined) byDirectory.set(directory, [path.slice(cut + 1)]);
    else bucket.push(path.slice(cut + 1));
  }

  const problems: StructureProblem[] = [];
  for (const [directory, files] of byDirectory) {
    if (files.length <= max) continue;
    problems.push({
      directory,
      count: files.length,
      files: [...files].sort((a, b) => a.localeCompare(b)),
    });
  }
  return problems.sort((a, b) => b.count - a.count || a.directory.localeCompare(b.directory));
}
