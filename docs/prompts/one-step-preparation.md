# One-step preparation: a prompt for a coding agent

This is a prompt for a coding agent working in an app that already has a template studio built on slotdeck,
where a person who uploads a template is met by long lists of suggestions and issues before the template can
be used.

Paste everything below the line to the agent. It asks for one thing: an operation that prepares a template by
itself, a short summary to look at, and publishing. Server logic only, for slotdeck 0.3.1 or later. The full
set of rules for a studio is in [template-studio.md](template-studio.md); this prompt is the part to do first.

---

You are working in an application that has a template studio built on slotdeck. Today, after a person uploads
a template and it is drafted, they are given every suggestion and every issue to go through one by one. There
are too many, and nobody reaches the end. Your task is to put one operation in front of those lists, so that
the path from an upload to a template the chat can use is: upload, prepare, look, publish.

This is about what the server does and returns. Do not design screens. The reference is
`node_modules/slotdeck/docs/template-studio.md`, section "prepareManifest", not the internet.

## Find out first

Change nothing yet. Report in a few lines:

1. The installed version (`node_modules/slotdeck/package.json`). `prepareManifest` is in 0.3.1 and later. Below
   that, say so and stop: the upgrade comes first, with "Upgrading from" in
   `node_modules/slotdeck/docs/changelog.md`.
2. Where the draft runs (`draftManifest`), what is stored after it, and what the person is then given: how many
   suggestions, issues and other items.
3. Whether the chat side gives `presentationTools` an icon set (`icons`) for a template, and whether templates
   are loaded with fonts (`loadTemplate(pptx, manifest, { fonts })`).
4. What makes a template usable by the chat today: a flag, a default, or simply being stored.

## What to build

**1. Prepare.** One operation for a template that has no layouts yet. It drafts once, then settles the draft
without a model:

```ts
const draft = await draftManifest(pptx, { llm, name, onProgress, signal });
if (draft.manifest.patterns.length === 0) throw new Error('no layout was found in this file'); // store nothing

// hasIcons: whether the chat side has an icon set for this template
const prepared = prepareManifest(loadTemplate(pptx, draft.manifest, fonts ? { fonts } : {}), { icons: hasIcons });
```

`prepareManifest` takes the suggestions that are facts of the slide, lets rows widen, adds the mirrored and
text-only variants, solves the issues that have one answer, and takes out a layout that cannot be rendered.
With `icons: false` it keeps the template's own pictograms, so that layouts with icons are still offered. It
renames nothing.

Store `prepared.manifest` as the template's manifest, not yet used by the chat. Keep `prepared.done`,
`prepared.left` and `draft.skipped` with it. Use `prepared.fields` as the first value of each deck field (the
`fields` the chat side passes to `presentationTools`): it is the text the template itself shows there.

**2. A summary as the answer.** The response to prepare holds, in this order, and nothing else:

1. numbers: layouts made, how many of them are added variants (`prepared.added.length`), slides skipped
   (`draft.skipped.length`);
2. `prepared.left`: the few sentences only a person can settle;
3. where to get the sample deck, rendered at the fewest and at the most items:
   `sampleDeck(template, { items: 'min' })` and `{ items: 'max' }`, then
   `renderDeck(template, deck, { assets, layout: 'adaptive' })`. Mark it as a sample filled to every limit, in
   its file name and wherever it is shown, or it is read as a broken real deck;
4. whether it may be published: `prepared.sound`.

**3. Publish.** One operation that makes the prepared manifest the one the chat uses. A person starts it after
reading the summary and looking at the sample. Refuse it while `prepared.sound` is false; `left` says why. The
chat side must build its template and tools from the manifest that is stored now, not from one it loaded at
start-up, or the published template never arrives.

**4. The same for a template that already has layouts.** Never draft it again: a second draft renames every
layout, drops every correction, and leaves stored presentations pointing at layouts that no longer exist. Run
only the second half, on the stored manifest:

```ts
const prepared = prepareManifest(loadTemplate(pptx, storedManifest, fonts ? { fonts } : {}), { icons: hasIcons });
```

Show its `left` and how many things are in `done`; store its manifest when the person agrees. Every layout id
and slot name stays, so stored presentations keep working. A second run changes nothing, so it is safe to
offer at any time.

**5. Keep the lists, behind.** Whatever the studio offers today for going through suggestions and issues one
by one stays reachable, for the person who wants to look closer. It is not what an upload leads to, and
nothing in it is needed before publishing. After prepare, `suggestChanges(template)` lists only the variants
that leave out one optional text.

## Rules

- `prepareManifest` calls no model. The draft is the only model step, and it runs once for a template. Keep
  whatever limits and stop the app has on it.
- Prepare never publishes by itself. A person publishes.
- A failed draft stores nothing.
- To take less by itself, pass `variants: []` or `widen: false`. Do not add `'without'` to `variants` unless
  asked: there is one such layout for every optional text.

## Tests

No test calls a live model. Script one: the draft's first request has `fields` in `request.schema.properties`
and is answered `{ language: 'en', fields: [] }`; a request with `shapes` there is one slide, answered
`{ usable, reason, id, description, figures, shapes: [{ shape, role, name, optional }] }` with ids from the
`enum` of `shape`, or `{ usable: false, reason: 'scripted' }` to skip it.

1. Upload, prepare, publish: the template the chat side then builds has the prepared layouts.
2. The response to prepare holds the numbers, `left`, the sample and `sound`, and no list of suggestions.
3. Nothing reaches the chat side before publish, and publish is refused when `sound` is false.
4. Prepare on a template that has layouts makes no model request, loses no layout id and no slot name, and a
   second run returns an empty `done`.
5. A draft whose model throws stores nothing.

## Done when

- the report was given;
- the five tests pass;
- your final message says what an upload leads to now, what you left as it was, and what a person must still
  decide.
