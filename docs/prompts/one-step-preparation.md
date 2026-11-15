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
are too many, and nobody reaches the end. Your task is to change what an upload leads to, so that the path
from an upload to a template the chat can use is: upload, prepare, look, publish.

This is about what the server does and returns. Do not design screens. The reference is
`node_modules/slotdeck/docs/template-studio.md`, section "prepareManifest", not the internet.

## Find out first

Change nothing yet. Report in a few lines:

1. The installed version (`node_modules/slotdeck/package.json`). `prepareManifest` is in 0.3.1 and later. Below
   that, say so and stop: the upgrade comes first, with "Upgrading from" in
   `node_modules/slotdeck/docs/changelog.md`.
2. The operation an upload leads to today: where the draft runs (`draftManifest`), what starts it, what is
   stored after it, and what the person is then given: how many suggestions, issues and other items.
3. What the chat side gives `presentationTools` for a template: whether an icon set (`icons`), which `layout`
   (`'strict'` when it gives none), and where the values of the deck fields (`fields`) come from. And whether
   templates are loaded with fonts (`loadTemplate(pptx, manifest, { fonts })`).
4. What makes a template usable by the chat today: a flag, a default, or simply being stored. And which
   templates the chat uses now.

## What to build

**1. Prepare, in the place of today's draft.** The operation an upload leads to today becomes this one. Do not
add a second operation beside it and leave the old one where it is: then an upload still ends at the lists.
For a template that has no layouts yet it drafts once, then settles the draft without a model:

```ts
const draft = await draftManifest(pptx, { llm, name, onProgress, signal });
if (draft.manifest.patterns.length === 0) throw new Error('no layout was found in this file'); // store nothing

// hasIcons: whether the chat side has an icon set for this template
// layout: the layout mode the chat side gives presentationTools, 'strict' when it gives none
const prepared = prepareManifest(loadTemplate(pptx, draft.manifest, fonts ? { fonts } : {}), { icons: hasIcons, layout });
```

`prepareManifest` takes the suggestions that are facts of the slide, lets rows widen, adds the mirrored and
text-only variants, solves the issues that have one answer, and takes out a layout that cannot be used. With
`icons: false` it keeps the template's own pictogram in every icon slot the planner would have to fill, so
that layouts with icons are still offered; `layout` tells it which slots those are. It renames nothing.

Store `prepared.manifest` as the template's manifest, not yet used by the chat. Store with it
`prepared.sound`, `prepared.done`, `prepared.left`, `prepared.kept` and `draft.skipped`. Store `prepared.fields`
as the first value of each deck field, where the chat side reads the `fields` it passes to
`presentationTools`: it is the text the template itself shows there, and without a value for a deck field the
planner is offered no layout that has one.

The draft runs once for a template, and the server makes sure of it:

- Prepare refuses to draft a template that has layouts, and runs part 4 on it instead.
- While a draft runs for a template, a second start for the same template is refused, not queued.
- Any other way the app has of starting a draft gets the same two checks. A second draft renames every layout,
  drops every correction, and leaves stored presentations pointing at layouts that no longer exist.

**2. A summary as the answer.** The response to prepare holds, in this order, and nothing else:

1. numbers: layouts made, how many of them are added variants (`prepared.added.length`), slides skipped
   (`draft.skipped.length`);
2. `prepared.left`: the few sentences only a person can settle;
3. the deck fields and their first values (`prepared.fields`), for the person to replace with their own.
   Setting a deck field's value is an operation of the app: it writes where the chat side reads its `fields`
   and changes no manifest;
4. where to get the sample deck, rendered at the fewest and at the most items:
   `sampleDeck(template, { items: 'min' })` and `{ items: 'max' }`, then
   `renderDeck(template, deck, { assets, layout: 'adaptive' })`. Mark it as a sample filled to every limit, in
   its file name and wherever it is shown, or it is read as a broken real deck;
5. whether it may be published: `prepared.sound`.

No list of suggestions, no list of issues, and not the sentences of `done`: their number is enough.

**3. Publish.** One operation that makes the stored manifest the one the chat uses. A person starts it after
reading the summary and looking at the sample. It works out again whether the manifest is fit to use, on the
manifest as it is stored at that moment, and refuses when it is not:

```ts
function fitToUse(template: Template): boolean {
  if (checkManifest(template).some((issue) => issue.severity === 'error')) return false;
  try {
    for (const items of ['min', 'max'] as const) {
      const { deck, assets } = sampleDeck(template, { items });
      renderDeck(template, deck, { assets, layout: 'adaptive' });
    }
    return true;
  } catch {
    return false;
  }
}
```

A `sound` that was stored at prepare says nothing about a manifest that was changed since. The chat side must
build its template and tools from the manifest that is published now, not from one it loaded at start-up, or
the published template never arrives.

When a manifest is not fit to use, a sentence of `left` names the layout and says what is wrong with it. The
way on is an operation the app offers for it: the person corrects that layout, or removes it with
`changeManifest` and `{ op: 'removePattern', pattern }`, and publishes again. The last layout of a manifest
cannot be removed; if that is the one, the person corrects it or uploads the file as a new template.

If "published" is a new state in this app, every template the chat uses today is published from the start.
Adding the state must not take a template away from the chat.

**4. The same for a template that already has layouts.** Never draft it again. Run only the second half, on
the stored manifest, and tell it that the manifest is in use:

```ts
const prepared = prepareManifest(loadTemplate(pptx, storedManifest, fonts ? { fonts } : {}), {
  icons: hasIcons,
  layout,
  remove: false,          // stored presentations name these layouts: nothing they depend on is taken away
  fields: storedFields,   // the values the app has for the deck fields: only a missing one is asked about
});
```

With `remove: false` every layout id and every slot name stays. A layout that cannot be used stays too:
`prepared.sound` is then `false`, and `prepared.left` names it and says what is wrong with it. Show `left` and
how many things are in `done`. Before storing, check each stored presentation of the template with
`validateDeck(manifest, deck)` against the manifest as it is and as prepared, and show the person any that
got worse. One listed because a row now takes no fewer cards than it holds could not be rendered before
either.

Store the manifest when the person agrees. For a template that is published, store it when it is fit to use,
or when the manifest stored now is not fit either and the prepared one has no `checkManifest` error that the
stored one lacks: it is then no worse than what the chat uses, and the way on is the one of part 3. With the
manifest, for each deck field the app has no value for, store `prepared.fields[name]` where the chat side
reads its `fields`, as in part 1, and show it to the person to replace. A second run with the same options
changes nothing, so it is safe to offer at any time.

`prepareManifest` does not know what a person undid by hand. A variant they removed is added again, and rows
whose widening they took away widen again. If the studio lets people do either, pass `variants: []` and
`widen: false` here, or what they chose.

**5. Keep the lists, behind.** Whatever the studio offers today for going through suggestions and issues one
by one stays reachable, for the person who wants to look closer. It is not what an upload leads to, and
nothing in it is needed before publishing a template that is fit to use. After prepare with its default
options, `suggestChanges(template)` lists only the variants that leave out one optional text.

## Rules

- `prepareManifest` calls no model. The draft is the only model step. Keep whatever limits and stop the app
  has on it.
- Prepare never publishes by itself. A person publishes.
- A failed draft stores nothing.
- After a draft, a layout that cannot be used is taken out, and the first sentence of `left` says which and
  why. For a manifest in use, pass `remove: false`, always.
- To take less by itself, pass `variants: []` or `widen: false`. Do not add `'without'` to `variants` unless
  asked: there is one such layout for every optional text.
- `prepared.kept` lists the icon slots that were set to keep the template's pictogram because the app has no
  icon set, each with what it was before (`was`). Keep the list. When the app gets an icon set, those are the
  slots to give back to it, each with `{ op: 'setSlot', pattern, slot, empty: was }`.

## Tests

No test calls a live model. Script one: the draft's first request has `fields` in `request.schema.properties`
and is answered `{ language: 'en', fields: [] }`; a request with `shapes` there is one slide, answered
`{ usable, reason, id, description, figures, shapes: [{ shape, role, name, optional }] }` with ids from the
`enum` of `shape`, or `{ usable: false, reason: 'scripted' }` to skip it.

Each test goes the way the app goes: it starts from the same call an upload makes today, not from a function
that only the test calls.

1. Upload, prepare, publish. Then build the tools the way the chat side builds them, with the `fields`,
   `icons` and `layout` the chat side would pass, and ask for an outline with a scripted model. The catalogue
   in that request names the prepared layouts, among them one with a deck field and one with an icon slot. A
   template that is published and offers the planner next to nothing fails this test.
2. The response to an upload holds the numbers, `left`, the deck fields, the sample and `sound`, and no list
   of suggestions or issues.
3. Nothing reaches the chat side before publish. Publish is refused for a manifest that is not fit to use:
   store one whose only layout has a slot on a shape id its slide does not have, and try. It is also refused
   when the manifest was sound at prepare and was changed to that since.
4. Prepare on a template that has layouts makes no model request and loses no layout id and no slot name,
   also when one of its layouts has such a slot. Once its result is stored, a second run returns an empty
   `done`. Run it for a template that is published as well. After it, for a template whose deck field had no
   value, the catalogue built the way the chat side builds it names a layout with that deck field.
5. A draft whose model throws stores nothing. A second prepare for a template that has layouts makes no model
   request, and a second start while a draft runs is refused.
6. A template the chat could use before your change can still be used by it after, without anyone publishing.

## Done when

- the report was given;
- the six tests pass;
- your final message says what an upload leads to now, what became of the operation it led to before, what
  you left as it was, and what a person must still decide.
