/** The mark: a 16:9 slide with one adjust handle on its edge, in the overlay's chrome blue. */
export function Mark({ size = 22 }: { size?: number }) {
  return (
    <svg
      width={size}
      height={size}
      viewBox="0 0 24 24"
      aria-hidden="true"
      focusable="false"
      className="shrink-0"
    >
      <rect
        x="2.5"
        y="5.5"
        width="19"
        height="13"
        rx="1.5"
        fill="none"
        stroke="currentColor"
        strokeWidth="1.6"
      />
      <rect x="17.5" y="9.5" width="5" height="5" rx="0.8" className="fill-accent" />
    </svg>
  );
}

export function Wordmark() {
  return (
    <span className="inline-flex items-center gap-2 font-semibold tracking-tight">
      <Mark />
      PPTX Studio
    </span>
  );
}
