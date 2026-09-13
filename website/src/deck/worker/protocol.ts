import type { PackageCensus } from '@pptx-studio/census';
import type { Report } from '@pptx-studio/validate';
import type { PreservationCheck, RoundTripReport } from '@pptx-studio/writer';

import type { Failure } from '@/design/state';

export interface WorkerEnvironment {
  readonly hasDOMParser: boolean;
  readonly hasOffscreenCanvas: boolean;
  readonly hardwareConcurrency: number;
}

/** A part the visitor rewrote in the tab, to be replaced before validating or exporting. */
export interface ReplacedPart {
  readonly part: string;
  readonly bytes: ArrayBuffer;
}

export type DeckRequest =
  | { readonly kind: 'census'; readonly id: number; readonly bytes: ArrayBuffer }
  | {
      readonly kind: 'validate';
      readonly id: number;
      readonly bytes: ArrayBuffer;
      readonly replaced: readonly ReplacedPart[];
    }
  | {
      readonly kind: 'export';
      readonly id: number;
      readonly bytes: ArrayBuffer;
      readonly replaced: readonly ReplacedPart[];
    };

export interface CensusDone {
  readonly census: PackageCensus;
  readonly ms: number;
}

export interface ValidateDone {
  readonly report: Report;
  readonly ms: number;
}

export interface ExportDone {
  readonly bytes: ArrayBuffer;
  readonly rewritten: readonly string[];
  readonly streamed: number;
  readonly preservation: PreservationCheck;
  readonly comparison: RoundTripReport;
  readonly report: Report | null;
  readonly ms: number;
}

export type DeckResponse =
  | { readonly kind: 'ready'; readonly environment: WorkerEnvironment }
  | {
      readonly kind: 'progress';
      readonly id: number;
      readonly done: number;
      readonly total: number;
    }
  | ({ readonly kind: 'census'; readonly id: number } & CensusDone)
  | ({ readonly kind: 'validate'; readonly id: number } & ValidateDone)
  | ({ readonly kind: 'export'; readonly id: number } & ExportDone)
  | { readonly kind: 'failed'; readonly id: number; readonly failure: Failure };
