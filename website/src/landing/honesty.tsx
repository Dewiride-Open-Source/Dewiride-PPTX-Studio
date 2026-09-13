import { Callout } from '@/design/callout';

export function Honesty() {
  return (
    <section aria-labelledby="not-a-clone" className="mx-auto max-w-6xl px-4 py-12 sm:px-6">
      <h2 id="not-a-clone" className="text-2xl font-semibold tracking-tight">
        Not a PowerPoint clone
      </h2>
      <div className="mt-6">
        <Callout kind="warn">
          <p>
            PPTX Studio is not trying to be PowerPoint in a browser tab. Today it opens a deck,
            draws shapes, pictures and text, lets you make three kinds of edit, and writes the file
            back. Tables, charts, SmartArt, OLE objects, ink and 3-D models are carried through an
            export untouched but are not drawn yet - you will see their frames. There is no
            selection toolbar, no resize, no text editing and no layout switching yet; those are
            phases 4 to 7 of a published plan, and the plan says which. The API will change before
            1.0. What is here is measured rather than promised, and the status above will not be
            marked green ahead of the code.
          </p>
        </Callout>
      </div>
    </section>
  );
}
