const PRINCIPLES: readonly { title: string; body: string }[] = [
  {
    title: 'The file is the document.',
    body: 'Most tools read a .pptx into their own model and write a brand-new file from it, so anything the model did not understand is gone when you save. PPTX Studio keeps the original package and edits parts inside it. Anything you did not touch - charts, SmartArt, animations, macros, OLE objects - is streamed back out exactly as it came in, because nothing had to understand it.',
  },
  {
    title: 'Inheritance stays live.',
    body: 'A title on a slide usually has no position of its own; it gets one from the layout, which gets it from the master. Other editors bake those numbers in the moment they open the file, and from then on the slide is a pile of fixed boxes. Here a missing property stays missing and is resolved on demand, so you can always ask where a value came from - and later change the layout and have the slide follow.',
  },
  {
    title: 'Preservation by default, checked before you get the bytes.',
    body: 'You do not opt in to keeping content. The writer re-serialises only the parts you changed, then runs 31 rules that real PowerPoint is known to refuse and will not hand you a file it just broke. That is tested on 55 decks in CI, and the round trip opens in PowerPoint 365 with no repair prompt.',
  },
];

export function Principles() {
  return (
    <section aria-labelledby="what-it-does" className="mx-auto max-w-6xl px-4 py-12 sm:px-6">
      <h2 id="what-it-does" className="text-2xl font-semibold tracking-tight">
        What it does
      </h2>
      <div className="mt-6 grid gap-6 md:grid-cols-3">
        {PRINCIPLES.map((principle) => (
          <article
            key={principle.title}
            className="rounded-panel border border-line bg-surface p-5 shadow-panel"
          >
            <h3 className="text-base font-semibold">{principle.title}</h3>
            <p className="mt-2 text-[14px] leading-relaxed text-fg-muted">{principle.body}</p>
          </article>
        ))}
      </div>
    </section>
  );
}
