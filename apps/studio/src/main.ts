import { humanBytes, type PackageCensus } from '@pptx-studio/census';
import { StudioWorker, type CensusResult, type ExportResult } from './client.js';
import type { EditKind } from './export.js';
import type { WorkerEnvironment } from './protocol.js';
import { slidesView } from './slides.js';

/**
 * Gate 0: drop a `.pptx` on the page and get a live explorer of its internals.
 *
 * Framework-free on purpose. The plan puts a Next.js shell here, and it will
 * arrive - but the thing sub-phase 0.8 has to establish is the *Worker
 * boundary*, and `packages/react` does not exist until sub-phase 5.7, so a
 * React shell today would be a wrapper around nothing. Keeping the boundary in
 * plain DOM also keeps it honest: nothing below can quietly start depending on
 * a framework, which is the same contract every package in this repository is
 * held to.
 *
 * The page is also the benchmark harness. `globalThis.pptxStudio` is the hook
 * `tools/bench` drives from Playwright, so the numbers recorded for a 200 MB
 * deck come from this code path and not from a special one built to be fast.
 */

const worker = new StudioWorker();

function el<K extends keyof HTMLElementTagNameMap>(
  tag: K,
  className?: string,
  text?: string,
): HTMLElementTagNameMap[K] {
  const node = document.createElement(tag);
  if (className !== undefined) node.className = className;
  if (text !== undefined) node.textContent = text;
  return node;
}

function requireElement(id: string): HTMLElement {
  const node = document.getElementById(id);
  if (node === null) throw new Error('the page is missing #' + id);
  return node;
}

function table(headings: readonly string[], rows: readonly (readonly string[])[]): HTMLElement {
  const wrapper = el('div', 'scroll');
  const node = el('table');
  const head = el('thead');
  const headRow = el('tr');
  for (const heading of headings) headRow.append(el('th', undefined, heading));
  head.append(headRow);
  node.append(head);
  const body = el('tbody');
  for (const row of rows) {
    const tr = el('tr');
    for (const [index, cell] of row.entries()) {
      tr.append(el('td', index === 0 ? undefined : 'num', cell));
    }
    body.append(tr);
  }
  node.append(body);
  wrapper.append(node);
  return wrapper;
}

function section(title: string, body: HTMLElement, open = true): HTMLElement {
  const details = el('details', 'section');
  details.open = open;
  const summary = document.createElement('summary');
  summary.textContent = title;
  details.append(summary, body);
  return details;
}

function stat(label: string, value: string, hint?: string): HTMLElement {
  const box = el('div', 'stat');
  box.append(el('span', 'stat-label', label), el('span', 'stat-value', value));
  if (hint !== undefined) box.append(el('span', 'stat-hint', hint));
  return box;
}

function ms(value: number): string {
  return value < 1000 ? value.toFixed(1) + ' ms' : (value / 1000).toFixed(2) + ' s';
}

/**
 * The environment strip.
 *
 * `DOMParser: absent` is the reason `@pptx-studio/xml` exists, and it is shown
 * on the page rather than left in a document because it is the kind of premise
 * that quietly stops being true. If a browser ever does expose `DOMParser` to
 * workers, whoever opens this page finds out immediately.
 */
function renderEnvironment(environment: WorkerEnvironment): HTMLElement {
  const body = el('div', 'stats');
  body.append(
    stat(
      'DOMParser in the worker',
      environment.hasDOMParser ? 'present' : 'absent',
      environment.hasDOMParser
        ? 'unexpected - the tokenizer premise is worth re-checking'
        : 'why this project has its own tokenizer',
    ),
    stat('XMLHttpRequest', environment.hasXMLHttpRequest ? 'present' : 'absent'),
    stat(
      'OffscreenCanvas',
      environment.hasOffscreenCanvas ? 'present' : 'absent',
      'phase 3.2 needs it',
    ),
    stat(
      'crossOriginIsolated',
      String(environment.crossOriginIsolated),
      environment.crossOriginIsolated ? 'memory is measurable' : 'no memory measurement',
    ),
    stat('hardwareConcurrency', String(environment.hardwareConcurrency)),
  );
  return section('Worker environment', body, false);
}

function renderTimings(result: CensusResult): HTMLElement {
  const fastest = [...result.runs].sort((a, b) => a.totalMs - b.totalMs)[0];
  const body = el('div');
  const stats = el('div', 'stats');
  stats.append(
    stat('archive', humanBytes(result.bytes)),
    stat(
      'parse, best of ' + String(result.runs.length),
      fastest === undefined ? '-' : ms(fastest.totalMs),
    ),
    stat(
      'XML throughput',
      fastest === undefined || fastest.scanMs === 0
        ? '-'
        : (fastest.xmlBytesScanned / 1024 / 1024 / (fastest.scanMs / 1000)).toFixed(0) + ' MiB/s',
      'tokenize and count',
    ),
    stat('dispatch', ms(result.dispatchMs), 'transfer, not copy'),
    stat('round trip', ms(result.wallMs)),
  );
  body.append(stats);

  if (result.memory.available) {
    const before = result.memory.beforeBytes;
    const after = result.memory.afterBytes;
    body.append(
      el(
        'p',
        'note',
        'Worker memory ' +
          (before === null ? '?' : humanBytes(before)) +
          ' before, ' +
          (after === null ? '?' : humanBytes(after)) +
          ' after.',
      ),
    );
  }

  const stages = result.stages;
  if (stages !== null) {
    const rate = (ms: number): string =>
      ms <= 0 ? '-' : (stages.xmlBytes / 1024 / 1024 / (ms / 1000)).toFixed(0) + ' MiB/s';
    body.append(
      el(
        'p',
        'note',
        'Stage by stage over ' +
          humanBytes(stages.xmlBytes) +
          ' of XML in ' +
          String(stages.xmlParts) +
          ' parts, ' +
          String(stages.tokens) +
          ' tokens. Measured after the census runs, so nothing is charged for JIT warm-up.',
      ),
      table(
        ['stage', 'time', 'throughput'],
        [
          ['inflate', ms(stages.inflateMs), rate(stages.inflateMs)],
          ['decode UTF-8', ms(stages.decodeMs), rate(stages.decodeMs)],
          ['tokenize', ms(stages.tokenizeMs), rate(stages.tokenizeMs)],
        ],
      ),
    );
  }

  body.append(
    table(
      ['run', 'open', 'inflate', 'scan', 'total'],
      result.runs.map((run) => [
        String(run.index + 1),
        ms(run.openMs),
        ms(run.inflateMs),
        ms(run.scanMs),
        ms(run.totalMs),
      ]),
    ),
  );
  return section('Timing', body);
}

function renderSummary(census: PackageCensus): HTMLElement {
  const body = el('div', 'stats');
  const presentation = census.presentation;
  body.append(
    stat('parts', String(census.archive.parts), String(census.archive.entries) + ' zip entries'),
    stat(
      'inflated',
      humanBytes(census.archive.declaredInflatedBytes),
      (census.archive.declaredInflatedBytes / Math.max(1, census.archive.compressedBytes)).toFixed(
        1,
      ) + ':1',
    ),
    stat('slides', presentation === null ? '-' : String(presentation.slides)),
    stat(
      'layouts / masters',
      presentation === null
        ? '-'
        : String(presentation.slideLayouts) + ' / ' + String(presentation.slideMasters),
    ),
    stat(
      'relationships',
      String(census.relationships.total),
      String(census.relationships.external) + ' external',
    ),
    stat('text', String(census.text.characters) + ' chars', String(census.text.runs) + ' runs'),
  );
  return section('Package', body);
}

function renderFeatures(census: PackageCensus): HTMLElement {
  return section(
    'Features (' + String(census.features.length) + ')',
    table(
      ['feature', 'count', 'arrives in'],
      census.features.map((feature) => [
        feature.label,
        String(feature.count),
        feature.phase === 'preserve'
          ? 'carried across, never rendered'
          : 'sub-phase ' + feature.phase,
      ]),
    ),
  );
}

function renderParts(census: PackageCensus): HTMLElement {
  return section(
    'Parts (' + String(census.parts.length) + ')',
    table(
      ['part', 'content type', 'bytes', 'zipped', 'elements', 'depth', 'rels'],
      census.parts.map((part) => [
        part.name,
        part.contentType,
        String(part.bytes),
        String(part.compressedBytes),
        part.elements === null ? '-' : String(part.elements),
        part.maxDepth === null ? '-' : String(part.maxDepth),
        String(part.relationships),
      ]),
    ),
    false,
  );
}

function renderRelationships(census: PackageCensus): HTMLElement {
  const body = el('div');
  body.append(
    table(
      ['relationship type', 'count'],
      census.relationships.byType.map((entry) => [entry.name, String(entry.count)]),
    ),
  );
  if (census.relationships.dangling.length > 0) {
    body.append(
      el('p', 'note', 'Dangling edges, which PowerPoint tolerates and an export preserves:'),
    );
    body.append(
      table(
        ['source', 'id', 'target'],
        census.relationships.dangling.map((entry) => [entry.source, entry.id, entry.target]),
      ),
    );
  }
  return section('Relationships', body, false);
}

function renderNamespaces(census: PackageCensus): HTMLElement {
  return section(
    'Namespaces (' + String(census.namespaces.length) + ')',
    table(
      ['namespace', 'prefixes', 'kind', 'count'],
      census.namespaces.map((namespace) => [
        namespace.uri === '' ? '(no namespace)' : namespace.uri,
        namespace.prefixes.join(' ') || '(default)',
        [
          namespace.standard ? 'ECMA-376' : namespace.extension ? 'extension' : 'unknown',
          namespace.required ? 'Requires' : '',
          namespace.ignorable ? 'Ignorable' : '',
        ]
          .filter((flag) => flag !== '')
          .join(' · '),
        String(namespace.count),
      ]),
    ),
    false,
  );
}

function renderProblems(census: PackageCensus): HTMLElement | null {
  if (census.problems.length === 0) return null;
  return section(
    'Problems (' + String(census.problems.length) + ')',
    table(
      ['severity', 'code', 'message', 'part'],
      census.problems.map((problem) => [
        problem.severity,
        problem.code,
        problem.message,
        problem.part ?? '',
      ]),
    ),
  );
}

// ------------------------------------------------------------------- Gate 1

/**
 * Where the bytes come from a second time.
 *
 * The deck is **transferred** into the worker, which detaches it, so the page
 * no longer has the archive it just inspected - by design, since holding a
 * second copy of a 200 MB file on the main thread is the exact cost the
 * transfer exists to avoid. A `File` can simply be read again, and a URL can be
 * fetched again, so what is kept here is the way back to the bytes rather than
 * the bytes.
 */
interface Source {
  readonly name: string;
  readonly read: () => Promise<ArrayBuffer>;
}

let source: Source | null = null;

/** The content types PowerPoint will accept under each extension. */
const MIME: Readonly<Record<string, string>> = {
  pptx: 'application/vnd.openxmlformats-officedocument.presentationml.presentation',
  pptm: 'application/vnd.ms-powerpoint.presentation.macroEnabled.12',
  ppsx: 'application/vnd.openxmlformats-officedocument.presentationml.slideshow',
  potx: 'application/vnd.openxmlformats-officedocument.presentationml.template',
};

/** `deck.pptm` -> `deck.pptx-studio.pptm`. The extension has to survive. */
function exportName(name: string): string {
  const dot = name.lastIndexOf('.');
  return dot === -1 ? name + '.pptx-studio' : name.slice(0, dot) + '.pptx-studio' + name.slice(dot);
}

function extensionOf(name: string): string {
  return name.slice(name.lastIndexOf('.') + 1).toLowerCase();
}

/**
 * Hand the file to the browser.
 *
 * An object URL and a synthetic click, which is the only download mechanism
 * that works without a server: a `data:` URL would base64 the whole archive -
 * a third more bytes, built on the main thread - and the File System Access
 * API is not in Firefox or Safari. The URL is revoked on the next task rather
 * than immediately, because revoking it before the browser has started reading
 * cancels the download in Chromium.
 */
function download(bytes: Uint8Array, name: string): void {
  const type = MIME[extensionOf(name)] ?? 'application/octet-stream';
  const url = URL.createObjectURL(new Blob([bytes as unknown as BlobPart], { type }));
  const anchor = el('a');
  anchor.href = url;
  anchor.download = name;
  anchor.style.display = 'none';
  document.body.append(anchor);
  anchor.click();
  anchor.remove();
  setTimeout(() => {
    URL.revokeObjectURL(url);
  }, 0);
}

/**
 * What the export checked, said plainly.
 *
 * The two numbers worth reading are `rewritten` and `streamed`: they are the
 * architecture's central claim in a form that can be wrong. A no-op export that
 * rewrote anything at all would mean a part was serialised that nobody edited,
 * and the whole preservation argument rests on that never happening.
 */
function renderExportOutcome(result: ExportResult): HTMLElement {
  const outcome = result.outcome;
  const body = el('div');
  const stats = el('div', 'stats');
  stats.append(
    stat('written', humanBytes(outcome.bytesOut), 'read ' + humanBytes(outcome.bytesIn)),
    stat(
      'parts rewritten',
      String(outcome.rewritten.length),
      outcome.rewritten.length === 0 ? 'nothing was edited' : outcome.rewritten.join(', '),
    ),
    stat('parts streamed', String(outcome.streamed), 'copied still compressed'),
    stat(
      'preservation',
      outcome.preservation.skipped === null
        ? String(outcome.preservation.checked) + ' entries identical'
        : 'skipped',
      outcome.preservation.skipped ?? 'compared against the archive we read',
    ),
    stat(
      'firewall',
      outcome.report === null
        ? 'off'
        : String(outcome.report.checked.length) +
            ' rules, ' +
            String(outcome.report.findings.length) +
            ' findings',
      outcome.report?.ok === true ? 'nothing blocking' : 'blocked',
    ),
    stat('export', ms(outcome.ms), 'in the worker'),
  );
  body.append(stats);

  const differences = outcome.comparison.differences;
  body.append(
    el(
      'p',
      'note',
      differences.length === 0
        ? 'Compared as documents rather than as bytes: ' +
            String(outcome.comparison.parts) +
            ' parts, all identical (' +
            String(outcome.comparison.xml) +
            ' as canonical XML, ' +
            String(outcome.comparison.relationships) +
            ' as relationship graphs, ' +
            String(outcome.comparison.binary) +
            ' by SHA-256).'
        : String(differences.length) +
            ' of ' +
            String(outcome.comparison.parts) +
            ' parts differ, which is what an edit looks like from the outside:',
    ),
  );
  if (differences.length > 0) {
    body.append(
      table(
        ['part', 'kind', 'what'],
        differences.map((difference) => [difference.part, difference.kind, difference.detail]),
      ),
    );
  }

  if (outcome.report !== null && outcome.report.findings.length > 0) {
    body.append(
      table(
        ['rule', 'severity', 'part', 'message'],
        outcome.report.findings.map((finding) => [
          finding.rule,
          finding.severity,
          finding.where.part,
          finding.message,
        ]),
      ),
    );
  }

  return body;
}

async function save(edit: EditKind): Promise<ExportResult> {
  if (source === null) throw new Error('nothing has been opened yet');
  const target = requireElement('export-result');
  target.replaceChildren(el('p', 'note', 'exporting …'));
  setStatus('exporting ' + source.name + ' …', true);

  try {
    const result = await worker.export(source.name, await source.read(), edit);
    target.replaceChildren(renderExportOutcome(result));
    download(result.bytes, exportName(result.name));
    setStatus(
      exportName(result.name) +
        ' — ' +
        String(result.outcome.rewritten.length) +
        ' part(s) rewritten, ' +
        String(result.outcome.streamed) +
        ' streamed',
    );
    return result;
  } catch (error) {
    const message = error instanceof Error ? error.message : String(error);
    target.replaceChildren(el('p', 'note', 'the export was refused: ' + message));
    setStatus(message);
    throw error;
  }
}

/**
 * The Gate 1 panel: two buttons and everything the export checked.
 *
 * Built here rather than in `index.html` because it only makes sense once a
 * deck is open, and a Save button that is present and inert before then is a
 * worse answer than one that is not present.
 */
function renderExport(): HTMLElement {
  const body = el('div');
  body.append(
    el(
      'p',
      'note',
      'Everything is read and written in the Worker in this tab. Nothing is uploaded. ' +
        'The bytes you get are only handed over once the preservation check and the ' +
        'twenty-nine repair-firewall rules have both passed.',
    ),
  );

  const actions = el('div', 'actions');
  const plain = el('button', 'primary', 'Save a copy');
  plain.type = 'button';
  plain.addEventListener('click', () => void save('none'));
  const stamped = el('button', undefined, 'Save with a stamp');
  stamped.type = 'button';
  stamped.title =
    'Sets cp:lastModifiedBy in docProps/core.xml, and nothing else. One part is ' +
    'serialised afresh; every other part is copied out of the source archive.';
  stamped.addEventListener('click', () => void save('stamp'));
  actions.append(plain, stamped);
  body.append(actions);

  const result = el('div');
  result.id = 'export-result';
  body.append(result);
  return section('Hand it back', body);
}

function render(result: CensusResult): void {
  const output = requireElement('output');
  output.replaceChildren();

  const heading = el('h2', 'file');
  heading.textContent = result.name + ' — ' + humanBytes(result.bytes);
  output.append(heading);

  const problems = renderProblems(result.census);
  if (problems !== null) output.append(problems);
  // The slides go first, because a picture of the deck is what anyone who
  // dropped one came for. Everything below it is the package, not the deck.
  if (source !== null) {
    const read = source.read;
    output.append(section('Slides — geometry only, no text yet', slidesView(read)));
  }
  output.append(
    renderExport(),
    renderTimings(result),
    renderSummary(result.census),
    renderFeatures(result.census),
    renderRelationships(result.census),
    renderNamespaces(result.census),
    renderParts(result.census),
    renderEnvironment(result.environment),
  );
}

function setStatus(text: string, busy = false): void {
  const status = requireElement('status');
  status.textContent = text;
  status.dataset['busy'] = busy ? 'yes' : 'no';
}

async function inspect(
  name: string,
  buffer: ArrayBuffer,
  repeat = 1,
  stages = false,
): Promise<CensusResult> {
  setStatus('reading ' + name + ' …', true);
  try {
    const result = await worker.run(name, buffer, {
      repeat,
      stages,
      onProgress: (done, total) => {
        setStatus('scanning part ' + String(done) + ' of ' + String(total), true);
      },
    });
    render(result);
    setStatus(name + ' — ' + String(result.census.archive.parts) + ' parts');
    return result;
  } catch (error) {
    setStatus(error instanceof Error ? error.message : String(error));
    throw error;
  }
}

async function inspectFile(file: File, repeat = 1): Promise<CensusResult> {
  source = { name: file.name, read: () => file.arrayBuffer() };
  return inspect(file.name, await file.arrayBuffer(), repeat);
}

async function inspectUrl(url: string, repeat = 1): Promise<CensusResult> {
  const response = await fetch(url);
  if (!response.ok) throw new Error('could not fetch ' + url + ': ' + String(response.status));
  const name = url.slice(url.lastIndexOf('/') + 1);
  source = {
    name,
    read: async () => {
      const again = await fetch(url);
      if (!again.ok) throw new Error('could not fetch ' + url + ': ' + String(again.status));
      return again.arrayBuffer();
    },
  };
  return inspect(name, await response.arrayBuffer(), repeat, true);
}

function wire(): void {
  const drop = requireElement('drop');
  const picker = requireElement('picker');

  for (const type of ['dragenter', 'dragover'] as const) {
    drop.addEventListener(type, (event) => {
      event.preventDefault();
      drop.classList.add('over');
    });
  }
  for (const type of ['dragleave', 'drop'] as const) {
    drop.addEventListener(type, () => drop.classList.remove('over'));
  }
  drop.addEventListener('drop', (event) => {
    event.preventDefault();
    const file = event.dataTransfer?.files.item(0);
    if (file !== null && file !== undefined) void inspectFile(file);
  });
  drop.addEventListener('click', () => picker.click());
  picker.addEventListener('change', () => {
    const file = (picker as HTMLInputElement).files?.item(0);
    if (file !== null && file !== undefined) void inspectFile(file);
  });

  worker.ready
    .then((environment) => {
      requireElement('output').append(renderEnvironment(environment));
      setStatus('drop a .pptx to inspect it');
    })
    .catch((error: unknown) => {
      setStatus(error instanceof Error ? error.message : String(error));
    });
}

/** The hook `tools/bench` drives. Same code path as a dropped file. */
interface StudioAutomation {
  readonly censusFromUrl: (url: string, repeat?: number) => Promise<CensusResult>;
  readonly environment: () => Promise<WorkerEnvironment>;
  /** Inspect a deck, then export it - the same path the two buttons take. */
  readonly exportFromUrl: (url: string, edit?: EditKind) => Promise<ExportResult>;
}

(globalThis as unknown as { pptxStudio: StudioAutomation }).pptxStudio = {
  censusFromUrl: (url, repeat = 1) => inspectUrl(url, repeat),
  environment: () => worker.ready,
  exportFromUrl: async (url, edit = 'none') => {
    await inspectUrl(url);
    return save(edit);
  },
};

wire();
