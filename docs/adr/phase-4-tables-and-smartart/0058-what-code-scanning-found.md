# 0058 — What code scanning found

Date: 2026-09-27
Status: **accepted** — the three open CodeQL alerts are answered, and so are two sites of the same
shape that no alert named. With each old form put back, a test sized for milliseconds took 7 to
25 seconds. 18/18 mutants killed, plus one inert control that survived as it should.

**Maintenance after 4.1 · supersedes
[0055](../phase-3-text/0055-the-website-installs-what-it-documents.md)'s reasons for starting npm and
pnpm through `cmd.exe` (§ Three consumers, one manifest, and its deviation of the same name).**
Dependabot had nothing open. What the repository's security page listed was three CodeQL alerts:

- **#2**, `js/polynomial-redos`, high;
- **#3**, `js/polynomial-redos`, high;
- **#8**, `js/shell-command-injection-from-environment`, medium.

## #2 — `sourcePartNameForRels`

`/^(\/(?:.*\/)?)_rels\/(.*)\.rels$/` backtracks once per `/_rels/` it can try as the directory, and
each try rescans to the end looking for `.rels`. CodeQL's witness is a name starting `/_rels/` with
many repetitions of `/_rels/a`. The function is exported from `@pptx-studio/opc`, so its input is
whatever a caller hands it. Deck names are capped at 1024 characters (X1.5) before they reach it,
which bounds the archive path but not the library.

The replacement makes three checks and one `lastIndexOf('/_rels/')`:

- the name starts with `/`;
- it ends with `.rels`;
- it contains `/_rels/`, whose last occurrence ends the directory.

The last occurrence is right because `relsPartNameFor` puts `_rels/` in the part's own, deepest
directory, and `.rels` holds no `/`, so the directory can never overlap the extension. An empty base
is the package root, as before.

**One behaviour changed on purpose.** The pattern's `.` does not match U+2028, U+2029, `\n` or `\r`,
so the function refused any name containing one. Nothing designed that; it is what `.` means.

`\n` and `\r` can never arrive from a deck, because X1.1 refuses control characters. U+2028 and
U+2029 can: M1.6 only warns on them, a UTF-8-flagged entry name decodes to them, and
`isRelationshipPartName` (`[^/]*`) accepts the `.rels` beside them. `PartStore.danglingRelationships()`
is documented "Never throws". It filters with `isRelationshipPartName`, then called this function,
so on such a deck it threw, and census threw at the same two calls (`census.ts:172`, `:501`). The
function now implements the grammar with `[^]` in place of `.`, and both callers are answered by
the one fix.

The test is exhaustive and compares against that grammar written as a pattern. It covers every name
of up to five pieces from `/`, `_rels/`, `.rels`, `_rels`, `a`, `.` and U+2028: 19,608 names. The
census call is not tested separately. Its tests build decks through opc's public writer, which never
sets the UTF-8 flag, and it reaches the fix through the same two calls the part-store test covers.

## #3 — `decoratedStretches`, and its twin in render-svg

`text.replace(/ +$/, '')` tries each space as the start of the run that reaches the end, and each
try scans to the glyph that stops it: quadratic in a run of spaces followed by anything. Run text
comes from the deck.

The replacement walks back from the end over U+0020 only. It must not be `trimEnd()`, which also
drops tab, no-break space, U+3000 and `\n`: glyphs PowerPoint draws the rule under.

**The twin.** `packages/render-svg/src/text/layout.ts` ran the same pattern in `ruleStretches` for
every underlined or struck piece, and no alert named it. The sweep of `packages/*/src` found no third
site:

- `geometry/src/paths/path.ts:186`'s `/0+$/` runs on `toFixed` output, at most about 123
  characters;
- `xml/src/parse/source.ts`'s encoding pattern reads the first 256 bytes;
- the rest are anchored or carry no trailing constraint.

The first render-svg test missed the twin. Its 2¹⁶ spaces at 32 pt overflowed the 10⁶ pt of an
unwrapped line, so the layout broke after the spaces, and the old pattern ran on a piece that ended
in them, which is its linear case. The mutant survived. At 5 pt, 2¹⁷ spaces fit one line and the test
asserts they do.

## Measured, with the old form put back

Each hostile input is sized so the fixed code answers in milliseconds, against a 250 ms or 2 s bound.
In Chromium, on this machine:

| site                    | input                                       | old form | bound  |
| ----------------------- | ------------------------------------------- | -------- | ------ |
| `sourcePartNameForRels` | `'/_rels/' + '/_rels/a' × 2¹⁵`              | 9.7 s    | 250 ms |
| `sourcePartNameForRels` | `'/' + '_rels/' × 2¹⁵ + '\n.rel'`           | 7.5 s    | 250 ms |
| `decoratedStretches`    | `'a' + ' ' × 2¹⁷ + 'b'`                     | 24.5 s   | 250 ms |
| render-svg rules        | one unwrapped line, `'a' + ' ' × 2¹⁷ + 'b'` | 19.2 s   | 2 s    |

## #8 — the candidate gate's `cmd.exe`

0055 started npm and pnpm through `cmd.exe /d /s /c "…"` on Windows because both are `.cmd` shims.
Node refuses to spawn a shim without a shell (CVE-2024-27980), and `shell: true` with an argument
list prints DEP0190. CodeQL reads a `cmd.exe` + `/c` argument list as a shell. It flagged the joined
string, built from absolute paths.

`tools/release/candidate/spawn.ts` starts neither shim, and no shell at all:

- **npm** runs as `node.exe` on the `node_modules/npm/bin/npm-cli.js` beside it, which is what
  `npm.cmd` runs. It refuses if that file is missing.
- **pnpm** has no fixed place beside `node.exe`. The one authoritative handle is `npm_execpath`,
  which pnpm sets for a script it runs: a `pnpm.mjs` runs under `node.exe`, a native `pnpm.exe`
  directly. Without it the gate refuses and says to run `pnpm candidate <work-dir>`, a new root
  script.
- Launched through pnpm, npm would read pnpm's `npm_config_*` and lifecycle variables as its own
  configuration, so those are dropped before npm starts. `NPM_CONFIG_USERCONFIG`, which the release
  job's `setup-node` exports, is kept.

Off Windows nothing changes: both tools are started by name with no shell, which is how CI and the
release run the gate. The Windows branch runs nowhere in CI. It is proved by an injected host in
`spawn.test.ts` and by one real run on this machine:

- `pnpm candidate tmp/candidate` packed 12, provenance 12/12, and ran typecheck, lint, test, smoke
  and build clean;
- plain `node` refused with the instruction.

## The same class, unflagged: `bisect --oracle command`

`commandOracle` pasted the candidate's path into a `shell: true` command unquoted. The path is a
file inside a `mkdtemp` directory under the temporary directory, whose name can hold a space.

The path now goes in `PPTX_STUDIO_CANDIDATE`, and `{}` becomes a quoted reference to it:

- `"$PPTX_STUDIO_CANDIDATE"` for sh, which never word-splits or re-expands it;
- `"%PPTX_STUDIO_CANDIDATE%"` for cmd.exe, which expands once and cannot see a `"` in a Windows
  path.

A template that quoted `{}` itself is now quoted twice, so `@pptx-studio/cli` takes a minor bump and
the README says `{}` arrives quoted.

The first test was weak. Its right answer was exit 1, and a shell that split the path on `&` also
ended in exit 1, from the unrecognised command after it, so two mutants survived. Now only exit 0
counts, for exactly one argument naming a file that holds the candidate's bytes, in a directory
named `a b & $HOME %PATH% ;x-`.

## Verification

- Every package's suite passes.
- Mutants, 18/18 killed:
  - `sourcePartNameForRels`, 5: first `/_rels/` instead of last; the pattern put back; line
    terminators refused again; no leading-slash check; root case dropped.
  - `decoratedStretches`, 2: `trimEnd`; the pattern put back.
  - render-svg rules, 2: the same two.
  - `bisect`, 3: `{}` as the raw path; the reference unquoted; the variable never set.
  - `spawn.ts`, 6:
    - the npm refusal dropped;
    - any `npm_execpath` accepted;
    - pnpm's variables passed to npm;
    - `NPM_CONFIG_USERCONFIG` stripped too;
    - a pnpm script spawned directly;
    - Windows treated as POSIX.
- One inert control, an unused extra variable in the bisect spawn, survived as it must.
- CI's Linux runner failed five of the injected-Windows cases that passed here. The resolver took
  paths apart with the host's `node:path`, and on Linux `dirname` of a drive path is `.` and
  `basename` does not split on `\`. It now uses `path.win32`, which is what the branch means on any host.

## Open questions

1. `sourcePartNameForRels('/ppt/_rels/.rels')` is `/`, the package root, though
   `relsPartNameFor('/')` only ever writes `/_rels/.rels`. A relationship part named that way would
   describe the part `/ppt/`, which is not a legal name. Whether to refuse it needs an ECMA-376
   Part 2 citation and a look at every caller's error handling. Unchanged here.
2. The tests keep the grammar and `/ +$/` as oracles over short inputs. CodeQL should see no
   uncontrolled source reaching them; the analysis of `main` after the merge says whether it agrees.
