# @pptx-studio/render-dom

## 0.1.0

### Minor Changes

- 63d04e0: `mountSlide` and `mountOverlay` can be installed.

  The live DOM renderer was the twelfth package and the only one never published. It sat at `0.0.0`
  in changesets' `ignore`, so `@pptx-studio/render-dom` did not exist on the registry at all and the
  two mount entry points were reachable only from a checkout of this repository.

  That was not a decision about the code. A trusted publisher is a setting on an npm package and npm
  has no way to attach one to a name the registry does not hold, so the first publish of a new name
  cannot go out over OIDC — the release would have reached it, failed to authenticate, and stopped
  partway. The name now exists, its trusted publisher points at this repository's `release.yml`, and
  the package is a normal member of the release.

  Nothing about the package itself changed. `shapeOverlay`, the debug overlay, stays in
  `render-svg` where it always was.
