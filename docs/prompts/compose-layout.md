# A new layout in the template's look: a prompt for a coding agent

This is a prompt for a coding agent working in an app that has a template studio built on slotdeck, where a
person needs a layout the template does not have: a text-only slide, three points side by side, an agenda, a
statement. Variants of a layout cannot give that: a text-only variant of a slide with a photo leaves half the
slide empty.

Paste everything below the line to the agent. It asks for one operation that makes such a layout from words,
a short look at the result, and keeping it or not. Server logic only, for slotdeck 0.4.0 or later. The full
set of rules for a studio is in [template-studio.md](template-studio.md).

---

You are working in an application that has a template studio built on slotdeck. A person can upload, prepare
and publish a template, but cannot add a layout the template does not have. Your task is to add that: inside a
template, the person says in words what layout they want ("a text-only slide with three key points"), and gets
a new layout made of the template's own styles and decorations, on the frame of its own slides. They look at
it and keep it or not.

This is about what the server does and returns. Do not design screens. The reference is
`node_modules/slotdeck/docs/template-studio.md`, section "composePattern", not the internet.

## Find out first

Change nothing yet. Report in a few lines, then say what you plan to change, file by file:

1. The installed version (`node_modules/slotdeck/package.json`). `composePattern` is in 0.4.0 and later. Below
   that, say so and stop: the upgrade comes first, with "Upgrading from" in
   `node_modules/slotdeck/docs/changelog.md`.
2. Where a template's pptx and manifest are stored, and whether the studio can already replace the pptx while
   keeping the manifest. `composePattern` returns both, and they belong together.
3. How a change to a published template reaches the chat, and what publishing checks.
4. Which model the studio's other model operations use, how it is chosen, and what limits and stop they have.

## What to build

**1. Compose.** One operation that takes a template and the person's words, and calls:

```ts
const made = await composePattern(template, words, { llm, effort: 'low', signal });
```

`template` is loaded from what is stored now, with fonts if the app loads templates with fonts. Pass
`effort: 'low'`: with the model's default reasoning one call took minutes; with light reasoning it takes about
ten seconds, and the designs were as good. It makes one or two model calls and never more.

Store nothing in the template yet. Keep the result as a candidate of that template: `made.pptx`,
`made.manifest`, `made.pattern`, and which version of the template it was made from. Answer with:

1. `made.summary` and `made.warnings`;
2. the new layout's description and slots: `made.manifest.patterns.at(-1)`;
3. where to get its sample deck, rendered at the fewest and the most items, from the candidate:
   `loadTemplate(made.pptx, made.manifest)`, then for the new pattern alone
   `sampleDeck(template, { items: 'min' })` and `{ items: 'max' }` and
   `renderDeck(template, deck, { assets, layout: 'adaptive' })`. Mark it as a sample filled to every limit.
   If the studio shows pictures of layouts, show the new one the same way.

When `composePattern` throws `DesignError`, answer with `error.problems` and keep nothing. A second compose
for the same template while one runs is refused, not queued.

**2. Keep.** One operation that stores a candidate as the template: `made.pptx` in place of the pptx and
`made.manifest` in place of the manifest, together, or neither. Refuse it when the template was changed after
the candidate was made: a candidate carries the whole manifest, and storing it would undo that change. The
person composes again instead. For a published template, check the candidate the way publishing checks before
storing; the new layout takes nothing away, so stored decks and other layouts stay as they were, and the
template stays published.

**3. Discard.** Drops a candidate. A candidate the person neither keeps nor discards is dropped after a while;
it holds a whole pptx.

**4. From a presentation (slotdeck 0.5.0 and later).** When the chat makes layouts for a presentation (see
[deck-layouts.md](deck-layouts.md)), a stored deck carries their designs in `deck.layouts`. Compose takes one of
those in place of words, and makes the candidate with no model:

```ts
const made = buildPattern(template, deck.layouts[index]);
```

The rest is as above: the answer, the sample deck, Keep and Discard. After Keep, presentations that use that
layout use the template's own. When the template was changed so that the design no longer fits,
`buildPattern` throws `DesignError`: answer with `error.problems`.

## Rules

- Nothing reaches the stored template, or the chat, before a person keeps a candidate.
- The pptx and the manifest are stored together. A manifest whose new pattern points at slide 37 with a pptx
  of 36 slides fails every deck that uses it.
- Never edit the design's boxes, styles or budgets in code of your own. A person who wants a different layout
  composes again with other words.
- The model call has the limits and the stop of the studio's other model calls.
- To take a kept layout away again, remove its pattern with `changeManifest` and `{ op: 'removePattern', pattern }`.
  Its slide stays in the pptx, unused.

## Tests

No test calls a live model. Script one: the request's `schema.properties.elements.items.properties.style.enum`
lists the style keys; answer with a design that uses them, for example three numbered points
(`point1Heading`, `point1Text`, ...) side by side, each on the panel the kit offers, under the kept `title`.
`designKit(template)` gives the keys and the grid, so the scripted answer can be written from it.

1. Compose, then keep: the stored pptx has one more slide, the stored manifest one more pattern, and the
   tools built the way the chat side builds them offer the new layout to the planner.
2. Compose, then discard: nothing stored changed.
3. A scripted design whose texts lie on each other is sent back once; when the second answer is as bad, the
   answer is the problems and nothing is kept.
4. Keep is refused for a candidate of a template that was changed after it was made.
5. On a published template, keeping a candidate leaves it published, and every stored deck of it still passes
   `validateDeck`.

## Done when

- the report and the plan were given;
- the five tests pass;
- your final message says what a person can now do, what is stored when, and what you left as it was.
