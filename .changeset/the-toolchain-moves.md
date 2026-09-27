---
'@pptx-studio/opc': patch
'@pptx-studio/xml': patch
'@pptx-studio/census': patch
'@pptx-studio/geometry': patch
'@pptx-studio/paint': patch
'@pptx-studio/text': patch
'@pptx-studio/model': patch
'@pptx-studio/render-svg': patch
'@pptx-studio/render-dom': patch
'@pptx-studio/validate': patch
'@pptx-studio/writer': patch
'@pptx-studio/cli': patch
---

Built with tsdown 0.23. The declarations take a new shape, and the API is unchanged.

- Every value is exported where it is declared (`export declare function …`), not in a trailing
  `export { … }`.
- The types are exported in one `export type { … }`.
- The JavaScript is byte-for-byte what 0.22 emitted.
- In geometry, opc, validate and xml, the source maps name fewer symbols.
