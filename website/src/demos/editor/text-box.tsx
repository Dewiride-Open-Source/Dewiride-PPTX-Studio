'use client';

import { useEffect, useRef, type KeyboardEvent, type ClipboardEvent } from 'react';

import { resolveAnchor, resolveInsets, type Sheet } from '@pptx-studio/model';
import type { Placed } from '@pptx-studio/render-svg';

/** One entry per paragraph, each split at its soft breaks. */
export type Paragraphs = readonly (readonly string[])[];

/** The face, size and colour the drawn text has, read off the first `<tspan>` render-svg emitted. */
export interface DrawnText {
  readonly fontFamily: string;
  readonly fontSize: number;
  readonly fontWeight: string;
  readonly fontStyle: string;
  readonly fill: string;
}

export function drawnText(surface: HTMLElement, cNvPrId: number): DrawnText | null {
  const groups = surface.querySelectorAll(`[data-text="${String(cNvPrId)}"]`);
  const span = groups[groups.length - 1]?.querySelector('tspan');
  if (span === null || span === undefined) return null;
  return {
    fontFamily: span.getAttribute('font-family') ?? 'sans-serif',
    fontSize: Number(span.getAttribute('font-size') ?? '18'),
    fontWeight: span.getAttribute('font-weight') ?? 'normal',
    fontStyle: span.getAttribute('font-style') ?? 'normal',
    fill: span.getAttribute('fill') ?? 'currentColor',
  };
}

const EMU_PER_POINT = 12700;
const JUSTIFY: Record<string, string> = { t: 'flex-start', ctr: 'center', b: 'flex-end' };

/** A paragraph per `<div>`, a soft break per `<br>`, which is how the browser's Enter and Shift+Enter write. */
function fill(root: HTMLElement, paragraphs: Paragraphs): void {
  root.replaceChildren();
  for (const segments of paragraphs) {
    const block = document.createElement('div');
    segments.forEach((segment, index) => {
      if (index > 0) block.append(document.createElement('br'));
      block.append(document.createTextNode(segment));
    });
    if (segments.every((segment) => segment === '')) block.append(document.createElement('br'));
    root.append(block);
  }
}

/** A soft break is a `<br>`, or under `pre-wrap` the newline Chromium writes instead. */
function segmentsOf(block: Node): string[] {
  const segments = [''];
  const walk = (node: Node): void => {
    if (node.nodeType === Node.TEXT_NODE) {
      (node.textContent ?? '').split('\n').forEach((part, index) => {
        if (index > 0) segments.push('');
        segments[segments.length - 1] += part;
      });
    } else if (node.nodeName === 'BR') {
      segments.push('');
    } else {
      for (const one of node.childNodes) walk(one);
    }
  };
  walk(block);
  // The browser keeps a trailing <br> in a block that ends with a line break.
  if (segments.length > 1 && block.lastChild?.nodeName === 'BR') segments.pop();
  return segments;
}

export function read(root: HTMLElement): string[][] {
  const paragraphs: string[][] = [];
  let loose: Node[] = [];
  const flushLoose = (): void => {
    if (loose.length === 0) return;
    const holder = document.createElement('div');
    for (const one of loose) holder.append(one.cloneNode(true));
    paragraphs.push(segmentsOf(holder));
    loose = [];
  };
  for (const node of root.childNodes) {
    if (node.nodeName === 'DIV' || node.nodeName === 'P') {
      flushLoose();
      paragraphs.push(segmentsOf(node));
    } else {
      loose.push(node);
    }
  }
  flushLoose();
  return paragraphs.length === 0 ? [['']] : paragraphs;
}

interface TextBoxProps {
  readonly placed: Placed;
  readonly sheet: Sheet;
  readonly drawn: DrawnText;
  readonly initial: Paragraphs;
  readonly onInput: (paragraphs: string[][]) => void;
  readonly onCommit: () => void;
  readonly onCancel: () => void;
}

/**
 * The selected shape's text, editable where it is drawn.
 *
 * Sits in the same transform as render-svg's text layer - points, rotated with
 * the shape - so the box's insets and size are the shape's own, and the drawn
 * text beneath it is hidden while this is open.
 */
export function TextBox({
  placed,
  sheet,
  drawn,
  initial,
  onInput,
  onCommit,
  onCancel,
}: TextBoxProps) {
  const root = useRef<HTMLDivElement>(null);
  const { frame } = placed;
  const insets = resolveInsets(placed.shape, sheet);
  const anchor = resolveAnchor(placed.shape, sheet)?.value ?? 't';
  const width = frame.cx / EMU_PER_POINT;
  const height = frame.cy / EMU_PER_POINT;
  const transform = `translate(${String(frame.x + frame.cx / 2)} ${String(frame.y + frame.cy / 2)}) rotate(${String(frame.rot)}) translate(${String(-frame.cx / 2)} ${String(-frame.cy / 2)}) scale(${String(EMU_PER_POINT)})`;

  useEffect(() => {
    const element = root.current;
    if (element === null) return;
    fill(element, initial);
    element.focus();
    const range = document.createRange();
    range.selectNodeContents(element);
    range.collapse(false);
    const selection = window.getSelection();
    selection?.removeAllRanges();
    selection?.addRange(range);
    // Only on open: what is typed afterwards is already in the element.
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, []);

  const keys = (event: KeyboardEvent<HTMLDivElement>) => {
    if (event.key === 'Escape') {
      event.preventDefault();
      onCancel();
    } else if (event.key === 'Enter' && (event.ctrlKey || event.metaKey)) {
      event.preventDefault();
      onCommit();
    }
    event.stopPropagation();
  };

  const paste = (event: ClipboardEvent<HTMLDivElement>) => {
    event.preventDefault();
    document.execCommand('insertText', false, event.clipboardData.getData('text/plain'));
  };

  return (
    <g transform={transform}>
      <foreignObject x={0} y={0} width={width} height={height} style={{ pointerEvents: 'auto' }}>
        <div
          style={{
            width,
            height,
            boxSizing: 'border-box',
            display: 'flex',
            flexDirection: 'column',
            justifyContent: JUSTIFY[anchor] ?? 'flex-start',
            padding: `${String(insets.top)}px ${String(insets.right)}px ${String(insets.bottom)}px ${String(insets.left)}px`,
          }}
        >
          <div
            ref={root}
            contentEditable
            suppressContentEditableWarning
            role="textbox"
            aria-multiline="true"
            aria-label={`Text of ${placed.shape.name === '' ? 'the shape' : placed.shape.name}`}
            spellCheck={false}
            onInput={() => {
              if (root.current !== null) onInput(read(root.current));
            }}
            onKeyDown={keys}
            onPaste={paste}
            onBlur={onCommit}
            onPointerDown={(event) => event.stopPropagation()}
            style={{
              outline: 'none',
              fontFamily: drawn.fontFamily,
              fontSize: drawn.fontSize,
              fontWeight: drawn.fontWeight,
              fontStyle: drawn.fontStyle,
              color: drawn.fill,
              lineHeight: 1.2,
              whiteSpace: 'pre-wrap',
              overflowWrap: 'break-word',
              caretColor: drawn.fill,
              minHeight: `${String(drawn.fontSize * 1.2)}px`,
            }}
          />
        </div>
      </foreignObject>
    </g>
  );
}
