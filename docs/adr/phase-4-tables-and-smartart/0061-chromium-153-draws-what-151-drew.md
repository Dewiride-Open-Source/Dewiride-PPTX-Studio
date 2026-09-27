# 0061 — Chromium 153 draws what 151 drew

Date: 2026-09-27
Status: **accepted** — Playwright 1.63.0 moves Chromium from 151.0.7922.34 to 153.0.8010.12. On
every committed slide, on both platforms, every raster and SVG digest is the same as 151's. The
records differ only in the browser they name.

**Maintenance after 4.1.** [0060](0060-the-toolchain-moves-and-typescript-waits.md) kept Playwright
back because it moves Chromium, and every Chromium-keyed record with it. This is that move, alone.
No renderer or source file changes here, so digest churn cannot hide a regression.

## The procedure

A baseline records the full user agent, and a run compares it by exact equality
(`FID_BROWSER_CHANGED`). That check also runs before a record, so a new Chromium cannot record over
an old baseline. It has to be removed and written fresh, which is
[0041](../phase-3-text/0041-what-the-harness-was-not-gating.md)'s procedure. Followed:

1. **Windows**, on this machine with Chromium 153 (`pnpm browsers`): both `win32-x64` files
   removed, then `pnpm fidelity --record` and `pnpm gate3 --record`.
2. **Linux**: both `linux-x64` files removed, the branch pushed, and CI dispatched. `record the
linux baselines` ran `--record --bootstrap` and uploaded `baselines-linux-x64`, which was
   downloaded under the ignored `tmp/` and committed with the manifest it wrote.
3. **`font-metrics.json`**: the probe-fonts job on the same dispatch measured on `windows-latest`.
   Its `font-metrics-windows` artifact was scored with the job's own command (`analyse.ts …
--fixture corpus/ground-truth/font-metrics.json`), and the manifest entry updated to the file.

## What moved

| record                        | entries | raster digests moved | SVG digests moved | other change    |
| ----------------------------- | ------- | -------------------- | ----------------- | --------------- |
| fidelity, `win32-x64`         | 254     | 0                    | 0                 | user agent      |
| fidelity, `linux-x64`         | 254     | 0                    | 0                 | user agent      |
| Gate 3 zoom, `win32-x64`      | 500     | 0                    | 0                 | user agent      |
| Gate 3 zoom, `linux-x64`      | 500     | 0                    | 0                 | user agent      |
| `font-metrics.json` (Windows) | —       | —                    | —                 | `chromium` only |

The fidelity score against PowerPoint is unchanged: corpus mean 9844 bp, noise floor 4 bp. Gate 3
holds. Every width in the probe-font fixture is the same, and so is every row of its analysis.

## The runs that are not records

- **Real faces (Linux, release-required):** green on 153. Its tolerance is derived, not recorded, and
  the widths stayed inside it.
- **Real faces (macOS, dispatch only, not required):** red, exactly as on every dispatch since
  2026-09-12 on 151: 182 failures, one `advance-quantum` and 181 `width`.
  - Chromium on macOS returns widths that are not whole pixels, 3363 of them.
  - `/System/Library/Fonts/NewYork.ttf` measures up to 17% narrower in the reader than in the
    browser.

  The two verdicts differ in one key, the Chromium version, which that failure names as its subject.
  153 changed nothing there. It stays what
  [0053](../phase-3-text/0053-what-linux-and-macos-said.md) left open.

## Verification

- `pnpm check` green on this machine against the new win32 records.
- `pnpm corpus`: 175 entries, no violations.
- CI on the branch green, all five release-required jobs among them.
- Dependabot #54 (website) and #12 (the rest of the test group) are superseded, and closed.

## Open questions

1. The macOS real-faces gate, as above; unchanged, and not this change's to answer.
