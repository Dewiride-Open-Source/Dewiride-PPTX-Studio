import projection from './plan.json';

export type SubPhaseState = 'done' | 'caveat' | 'in-progress' | 'todo';

export interface SubPhaseStatus {
  readonly id: string;
  readonly title: string;
  readonly state: SubPhaseState;
  readonly adr: string | null;
  readonly result: string | null;
  readonly note: string | null;
}

export interface PhaseStatus {
  readonly number: number;
  readonly title: string;
  readonly done: number;
  readonly total: number;
  readonly gate: {
    readonly title: string;
    readonly state: SubPhaseState;
    readonly adr: string | null;
  };
  readonly subPhases: readonly SubPhaseStatus[];
}

export interface PlanStatus {
  readonly recorded: number;
  readonly total: number;
  readonly gatesClosed: number;
  readonly gates: number;
  readonly roundTrip: { readonly decks: number; readonly total: number };
  readonly phases: readonly PhaseStatus[];
  readonly carried: readonly { title: string; detail: string; raisedIn: readonly string[] }[];
}

/** `plan.json` is written by `tools/repo/plan-status.ts` from `docs/plan/phases.json`; `pnpm references` refuses a stale copy. */
export const plan: PlanStatus = projection as PlanStatus;
