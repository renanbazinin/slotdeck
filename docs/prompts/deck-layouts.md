# Layouts made for a deck: a prompt for a coding agent

This is a prompt for a coding agent working in an app whose chat makes presentations with slotdeck. From 0.5.0,
the planner can make a layout the template does not have, in the template's own look, for slides whose content
none of its layouts shows well: three goals on a slide of running text, a plan in four steps. A text-only deck
then has cards, steps or a timeline where it had the same paragraphs slide four times.

Paste everything below the line to the agent. It asks for one setting on the path the chat already takes, a mark
in the plan the person sees, and nothing else in the chat. Server logic only, for slotdeck 0.5.0 or later.
Keeping such a layout in the template is a studio matter: see [compose-layout.md](compose-layout.md).

---

You are working in an application whose chat makes presentations with slotdeck. When a person asks for a
presentation, the planner can now make up to two new layouts for that presentation, in the template's own look,
for slides that none of the template's layouts shows well. Your task is to turn that on in the path the chat
already takes. Do not add a tool, a button or a second way to plan: the person asks for a presentation as
before, and gets better slides.

This is about what the server does and returns. Do not design screens. The reference is
`node_modules/slotdeck/docs/generating.md`, section "New layouts for a deck", and
`node_modules/slotdeck/docs/deck-format.md`, section "A deck's own layouts", not the internet.

## Find out first

Change nothing yet. Report in a few lines, then say what you plan to change, file by file:

1. The installed version (`node_modules/slotdeck/package.json`). This needs 0.5.0 or later. Below that, say so
   and stop: the upgrade comes first, with "Upgrading from" in `node_modules/slotdeck/docs/changelog.md`.
2. Every place the chat plans a presentation: the `presentationTools({ ... })` call, or calls of its own to
   `planOutline`, `planDeck` or `generateDeck`.
3. Where decks are stored, and every place the app's own code reads a deck's slides against the template's
   manifest: `validateDeck(manifest, deck)`, `slideAlternatives(manifest, deck, slide)`, a lookup of a slide's
   pattern in `template.manifest.patterns`, a list of layouts or thumbnails per slide.
4. How a plan is shown to the person, and whether it uses `thumbnail` and `layout_description` from the plan
   tool's output.
5. Which progress stages the chat shows while a presentation is made, and the limits and stop of its model calls.

## What to change

**1. One setting where the chat plans.** Add `composeLayouts` to the existing call, read from the app's
configuration with a default of 2; 0 turns it off.

```ts
const tools = presentationTools({ template, llm, store, /* as before */ composeLayouts: settings.composeLayouts ?? 2 });
```

When the app calls the planner itself, pass the same option to `planOutline`, `planDeck` or `generateDeck`.
Nothing else is called for it: do not call `composePattern` or `composeLayouts` from the chat. It applies to a
new plan only; editing a plan or a presentation keeps the layouts it has.

It costs about 10 to 25 seconds and two to five short model calls on a new plan, after the outline. Show the
progress event `{ stage: 'layouts' }` the way the other stages are shown, for example "Designing layouts for
this presentation".

**2. Mark it in the plan.** A slide whose layout was made for this presentation comes in the plan tool's output
with `new_layout: true`, the new layout's `layout_description`, and no `thumbnail`. Show such a slide with its
description and a mark that its layout is new, instead of a picture. Rendered previews of the presentation need
no change.

**3. Store the deck as it comes.** A deck may now carry `layouts`: the designs of the layouts made for it. Keep
the deck JSON whole. Every later render, edit and preview builds those layouts from it, without a model. Never
strip the field, never copy the layouts into the template, and never change the template's pptx or manifest
from the chat.

**4. The app's own reads of a deck.** Wherever step 3 of the report found code that reads a deck against
`template.manifest`, use the template with the deck's layouts instead:

```ts
const forDeck = withLayouts(template, deck.layouts);
const issues = validateDeck(forDeck.manifest, deck, { layout: 'adaptive' });
```

`renderDeck`, `editDeck`, `planEdit`, `editPlan`, `writeDeck` and the reviews do this themselves. `withLayouts`
builds once for a template object and a list of layouts, so reuse the loaded template object.

## Rules

- A layout that cannot be made is not an error: its slides keep the template's layout and the presentation is
  made. Do not report it to the person.
- When a stored presentation is rendered after its template was replaced and one of its layouts no longer fits,
  rendering throws `DesignError` naming the layout. Answer as for any other render failure.
- When the template gains a layout of the same id, because a person kept it in the studio, the presentation uses
  the template's own.
- Never edit a design's boxes, styles or budgets in code of your own.
- The model calls have the limits and the stop of the chat's other calls; the planner caps its own answers.

## Tests

No test calls a live model. Script one that answers by the request's `system`: the request that decides starts
with "You review the plan" (answer `{ layouts: [{ instruction, slides: [3] }] }`), a design request starts with
"You design one new slide layout" (answer with a design written from `designKit(template)`: its style keys and
its grid, such as three numbered cards under the kept `title`), and the others are the outline and the writing.

1. A text-only presentation of a cover, three slides of running text and a closing slide: the plan tool's slide
   3 has `new_layout: true` and no thumbnail; after creating it, the stored deck has `layouts`, and its pptx has
   the text written for slide 3.
2. The stored deck, read back and rendered with the template loaded afresh, renders again with a scripted model
   that throws when called.
3. A designer that fails every time: the presentation is made with the template's own layouts, and no error
   reaches the chat.
4. With the setting at 0, no request starts with "You review the plan".
5. The app's own code from step 4 accepts a deck with `layouts`.

## Done when

- the report and the plan were given;
- the five tests pass;
- your final message says what a person now gets, what is stored where, and what you left as it was.
