# Security policy

## Reporting a vulnerability

Please report security issues privately through
[GitHub Security Advisories](https://github.com/Dewiride-Open-Source/Dewiride-PPTX-Studio/security/advisories/new),
or by email to **security@dewiride.com**. Please do not open a public issue.

We aim to acknowledge within 3 working days and to have an assessment within 10. If you would like
credit in the advisory, say so and tell us how you want to be named.

## Threat model

PPTX Studio parses untrusted binary input — a `.pptx` is a ZIP archive full of XML written by
someone else — and it does so **in the user's browser**, with no server involved. Nothing is
uploaded anywhere. That removes a whole class of server-side risk and concentrates what remains in
the parser.

Things we consider in scope:

- **Decompression bombs.** Budgets are enforced _during_ decompression with a running counter, not
  from the ZIP header's declared sizes, because those are attacker-controlled. Limits: ≤ 200 MB
  input, ≤ 1 GB inflated, ≤ 200:1 incremental ratio, ≤ 20 000 entries.
- **Path traversal (zip-slip).** Part names are validated against the OPC grammar with
  percent-decoding, and `..`, absolute paths, backslashes and NUL bytes are rejected.
- **XXE and entity expansion.** `DOCTYPE` is rejected outright by the tokenizer. There is no
  external-entity resolution to attack because there is no entity resolution.
- **SVG injection.** SVG blips (`asvg:svgBlip`) are sanitized — `<script>`, `on*` attributes and
  external `<use>` are stripped — and rendered via `<img src={blobURL}>`, never inlined.
- **Denial of service in the parser.** The invariant is that the parser either returns a result with
  diagnostics or throws `PptxParseError`. It must never hang, never over-allocate, and never throw a
  raw `RangeError`. This is fuzzed.
- **Font shadowing.** Embedded and uploaded fonts are registered under a namespaced family name
  (`--pptx-<deck>-<hash>`), never under the name the file claims. `document.fonts` is
  document-global, so a deck embedding a font that calls itself "Arial" must not be able to restyle
  the surrounding application.

Out of scope: the contents of a deck the user opened on purpose, and rendering differences from
PowerPoint (those are correctness bugs — please file them as issues).

## Supply chain

- `minimumReleaseAge` is 24 hours, so a compromised release has a detection window before it can
  enter the tree.
- Dependency postinstall scripts are **deny-by-default**; each entry in `allowBuilds` in
  `pnpm-workspace.yaml` is an explicit, reviewed decision.
- `blockExoticSubdeps` is on: no git or tarball subdependencies.
- Releases are published from CI with npm trusted publishing (OIDC). No long-lived npm token exists,
  and every published artifact carries a provenance attestation.
- GitHub Actions are pinned to commit SHAs, not tags.

## Supported versions

Pre-1.0. Fixes land on `main` and in the next release; there are no maintained release branches yet.
