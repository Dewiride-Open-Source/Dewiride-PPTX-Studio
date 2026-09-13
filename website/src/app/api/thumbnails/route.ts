import { renderDeck } from '@pptx-studio/cli';

/**
 * Slides to SVG on the server, with no LibreOffice and no headless browser.
 *
 * `@pptx-studio/cli` reads font files off the disk, so this has to be the Node
 * runtime; on Edge the `node:fs` import fails at load. Everything else in this
 * app runs in the tab.
 */
export const runtime = 'nodejs';

/** 16 MB. A deck larger than this belongs on a background job, not a request. */
const MAX_BYTES = 16 * 1024 * 1024;

export async function POST(request: Request): Promise<Response> {
  const body = await request.arrayBuffer();
  if (body.byteLength === 0) {
    return Response.json({ error: 'no bytes' }, { status: 400 });
  }
  if (body.byteLength > MAX_BYTES) {
    return Response.json({ error: `over ${String(MAX_BYTES)} bytes` }, { status: 413 });
  }

  const url = new URL(request.url);
  const width = Number(url.searchParams.get('width') ?? '640');

  try {
    const startedAt = performance.now();
    // Every field is optional, and the three the CLI verb needs for itself -
    // `out`, `json`, `quiet` - are not part of this shape at all.
    const result = renderDeck(new Uint8Array(body), {
      width: Number.isFinite(width) ? Math.min(Math.max(width, 160), 3840) : 640,
    });

    return Response.json({
      slides: result.slides,
      width: result.width,
      height: result.height,
      fonts: result.fonts,
      missing: result.missing.map(
        (code) => `U+${code.toString(16).toUpperCase().padStart(4, '0')}`,
      ),
      facesIndexed: result.facesIndexed,
      fontDirectories: result.fontDirectories,
      ms: performance.now() - startedAt,
    });
  } catch (cause) {
    return Response.json(
      { error: cause instanceof Error ? cause.message : String(cause) },
      { status: 422 },
    );
  }
}
