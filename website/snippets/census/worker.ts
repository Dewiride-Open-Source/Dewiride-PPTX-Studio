import { censusPackage } from '@pptx-studio/census';

//#region example
// census.worker.ts
self.onmessage = ({ data }: MessageEvent<ArrayBuffer>) => {
  const census = censusPackage(new Uint8Array(data), {
    onProgress: (done, total) => self.postMessage({ progress: { done, total } }),
  });
  self.postMessage({ census }); // plain JSON by design, so it survives postMessage
};
//#endregion
