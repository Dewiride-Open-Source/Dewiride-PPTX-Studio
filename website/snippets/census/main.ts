declare const file: File;

//#region example
// main.ts
const worker = new Worker(new URL('./census.worker.ts', import.meta.url), { type: 'module' });
const buffer = await file.arrayBuffer();
worker.postMessage(buffer, [buffer]); // transferred, not copied
//#endregion

export {};
