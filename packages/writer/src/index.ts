/**
 * `@pptx-studio/writer` - handing the file back.
 *
 * Everything up to here has been about reading a package without damaging it.
 * This is where the damage would happen, and the whole design is one idea:
 * **the writer is not allowed to understand the document.**
 *
 * A writer that emits a `.pptx` from a model can only emit the features that
 * model has. Every feature it has not been taught is a feature the file loses
 * on save, which is why a deck that goes through most browser editors comes out
 * without its charts, its SmartArt, its animations, its OLE objects and its
 * macros - not because anyone decided to drop them, but because the export path
 * had no way to write them and no way to notice. Preservation cannot be a
 * feature with a coverage percentage, because the percentage is never a hundred
 * and the missing part is always somebody's.
 *
 * So a part that was not edited is never serialised at all. Its stored DEFLATE
 * stream is copied from the source archive into the output archive without ever
 * being inflated. We do not have to know what a `p:timing` tree means, or what
 * is inside an `mc:AlternateContent` branch, or what an OLE object is, in order
 * to carry them across intact - and because that is the *default*, there is no
 * category of content we can silently drop.
 *
 * ## What this package adds on top of `PartStore.write`
 *
 * The streaming itself belongs to `@pptx-studio/opc`. What lives here is the
 * export *path*: the four things that have to happen around it, in an order
 * that matters.
 *
 * - **`prepare` hooks**, so later phases can contribute save-time work - font
 *   embedding, autofit commit, media relationship dedupe - without this package
 *   learning what a font is.
 * - **Media collection**, which by default removes only what this session
 *   orphaned, so that opening and saving a file never changes it.
 * - **The preservation check**, which compares the archive we emitted against
 *   the archive we read. Not the store's opinion of it: the bytes.
 * - **The firewall**, `assertValid` from `@pptx-studio/validate`, which refuses
 *   to let the bytes leave if a rule this session broke would make PowerPoint
 *   show a repair prompt.
 *
 * The last two are the reason `exportPackage` returns a result object rather
 * than a `Uint8Array`. An export that succeeded silently is indistinguishable
 * from one that skipped its checks, and three of them skip themselves when the
 * caller does not supply a baseline. Saying which ran is the only way the
 * difference is visible.
 */

export {
  WriterError,
  WRITER_ERROR_CODES,
  isWriterError,
  type WriterErrorCode,
  type WriterErrorDetail,
} from './errors.js';

export { reachableParts, orphanedParts, type Reachability } from './gc/reachability.js';

export {
  collectGarbage,
  planCollection,
  isMediaPart,
  type CollectOptions,
  type CollectionPlan,
  type CollectionPolicy,
  type KeptPart,
} from './gc/collect.js';

export {
  runPrepare,
  type PrepareContext,
  type PrepareHook,
  type PrepareRecord,
} from './export/prepare.js';

export { assertPreserved, type PreservationCheck } from './export/preserve.js';

export {
  exportPackage,
  openPackage,
  type ExportOptions,
  type ExportResult,
  type OpenPackage,
} from './export/export.js';

export {
  bisectPackages,
  collectDelta,
  describeChange,
  flattenChanges,
  leafChanges,
  summarizeBisect,
  type BisectOptions,
  type BisectOutcome,
  type BisectResult,
  type Change,
  type ChangeKind,
  type Oracle,
  type RunEvent,
  type Span,
  type Verdict,
} from './oracle/bisect.js';

export {
  comparePackages,
  roundTripPackage,
  summarizeRoundTrip,
  type CompareHow,
  type CompareOptions,
  type Difference,
  type DifferenceKind,
  type PartComparison,
  type Relabelling,
  type RoundTripOptions,
  type RoundTripReport,
  type RoundTripResult,
} from './oracle/roundtrip.js';
