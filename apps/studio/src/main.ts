import { humanBytes, type PackageCensus } from '@pptx-studio/census';
import { CensusWorker, type CensusResult } from './client.js';
import type { WorkerEnvironment } from './protocol.js';

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

const worker = new CensusWorker();

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

function render(result: CensusResult): void {
  const output = requireElement('output');
  output.replaceChildren();

  const heading = el('h2', 'file');
  heading.textContent = result.name + ' — ' + humanBytes(result.bytes);
  output.append(heading);

  const problems = renderProblems(result.census);
  if (problems !== null) output.append(problems);
  output.append(
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
  return inspect(file.name, await file.arrayBuffer(), repeat);
}

async function inspectUrl(url: string, repeat = 1): Promise<CensusResult> {
  const response = await fetch(url);
  if (!response.ok) throw new Error('could not fetch ' + url + ': ' + String(response.status));
  const name = url.slice(url.lastIndexOf('/') + 1);
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
}

(globalThis as unknown as { pptxStudio: StudioAutomation }).pptxStudio = {
  censusFromUrl: (url, repeat = 1) => inspectUrl(url, repeat),
  environment: () => worker.ready,
};

wire();
