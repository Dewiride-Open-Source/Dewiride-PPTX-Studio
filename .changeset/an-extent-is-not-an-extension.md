---
'@pptx-studio/validate': patch
---

V028 no longer mistakes the extent of a transform for an extension.

The rule holds `mc:AlternateContent` and every `ext` inside an `extLst` opaque, by local name so
that vocabularies never read are covered too. It matched any element spelled `ext`, and the
extent inside `a:xfrm` is spelled `ext` as well - so a placeholder given a transform of its own,
which is what PowerPoint writes when one is dragged, was refused as a rebuilt extension. An `ext`
is now an extension only under an `extLst`.
