# Free slides and colours: a prompt for a coding agent

This is a prompt for a coding agent working in an app whose chat makes presentations with slotdeck. From 0.6.0
a presentation can hold what no layout of the template has: one paragraph alone on a slide, a chart beside
three cards, twelve charts, two hundred. The planner lays such a **free slide** out in the template's own look.
A person can also ask for another colour ("make it dark green", "slide 3 in red") and limit the layouts ("only
the text layout"), and the planner now does what was asked about the form of a deck before it follows its own
habits.

Paste everything below the line to the agent. Almost all of this is on by default, so the prompt asks for
little: the upgrade, four places where the app must not trip over the new things, and two switches in the
studio. No new tool and no new screen. Server logic only, for slotdeck 0.6.0 or later.

---

You are working in an application whose chat makes presentations with slotdeck, and which has a studio where
templates are prepared. slotdeck 0.6.0 lets a presentation hold slides that the software lays out itself
("free slides"), lets a person ask for other colours than the template's, and lets them limit the layouts a
presentation uses. The person does nothing new for any of it: they ask in the chat, in their own words, and it
happens. Your task is to upgrade and to make the app ready for what a plan and a deck can now contain. Do not
add a tool, a colour picker, or a second way to make a presentation.

The short path: upgrade, and build the tools as before. Free slides, colours and a person's list of layouts
then work in the chat with nothing else changed. The rest of this prompt is the few places where the app's own
code can trip over them, and two switches in the studio.

This is about what the server does and returns. Do not design screens. The references are
`node_modules/slotdeck/docs/free-slides.md`, `node_modules/slotdeck/docs/colors.md`,
`node_modules/slotdeck/docs/chat-tools.md` and "Upgrading from 0.5.2" in
`node_modules/slotdeck/docs/changelog.md`, not the internet.

## Find out first

Change nothing yet. Report in a few lines, then say what you plan to change, file by file:

1. The installed version (`node_modules/slotdeck/package.json`). Below 0.6.0 the upgrade is step 1 below; the
   references above come with it, so read them once it is installed.
2. Every place the chat plans, writes or changes a presentation: the `presentationTools({ ... })` call, or calls
   of its own to the planner. Which options it passes, above all `patterns`, `composeLayouts`, `layout`,
   `fields` and `icons`.
3. How the chat model gets the tools: whether `tools.definitions` is handed over as the tools made it, whether
   one set of definitions is kept for all templates, and whether the app's own system prompt describes the
   tools' arguments or says what the chat cannot do.
4. How a plan and a presentation are shown to the person: which keys of the tools' outputs are used (`layout`,
   `layout_description`, `thumbnail`, `new_layout`, `warnings`).
5. Every place the app's own code reads a deck: `validateDeck(manifest, deck)`, a lookup of a slide's pattern
   in `template.manifest.patterns`, code that walks `slide.slots` by the manifest's slot names.
6. In the studio: where a template's settings are shown and changed (`setManifest` through `changeManifest`),
   and where `checkManifest` results are shown.

## What to change

**1. Upgrade, and pass two settings.** Install slotdeck 0.6.0 or later the way the app installs its packages,
and read "Upgrading from 0.5.2" in its changelog. Free slides and colours are on unless turned off. Read both
from the app's configuration so that they can be turned off without a release:

```ts
const tools = presentationTools({ template, llm, store, /* as before */ free: settings.freeSlides ?? true, colors: settings.deckColors ?? true });
```

Two things to look at in what the report found:

- If the app passes `patterns` (its own list of allowed layouts), free slides are offered only when that list
  names `free`. Add `'free'` to the list unless the list exists to forbid them.
- If the app passes `composeLayouts`: with free slides on offer the planner composes no layouts, because a free
  slide is what a presentation now gets for content no layout shows. Leave the setting where it is; it acts
  again when free slides are off.

**2. Tool declarations belong to a template.** `tools.definitions` are now made for the template and the
options of that instance: the `layouts` argument lists that template's layouts, and the colour arguments are
left out for a template that keeps its colours. Make the tools for the template a conversation uses and hand
the chat model that instance's definitions. Do not keep one set for all templates. Build them from a template
loaded once and kept: making the tools now looks at the template's slides, which takes 0.2 to 0.6 seconds for
a template loaded anew and nothing for one kept. `edit_presentation` and
`edit_presentation_plan` may now be called without `instruction`, when a colour is all the person asked for:
do not reject such a call in code of your own.

**3. The plan and the presentation a person sees.** A slide the software lays out comes in the outputs with
`layout: "free"`, `free_layout: true`, a `layout_description` that says what it holds ("A slide laid out for
this presentation: 6 charts and a note.") and no `thumbnail`. Show its description with a mark that it is laid
out for this presentation, as for `new_layout`. Outputs carry `accent_color` when a colour is set and `layouts`
when the presentation is limited to some: show them where the app shows a plan's details. A warning may now
have no slide; the tools print such a warning as a sentence about the whole presentation, so show it as one.

**4. The app's own reads of a deck.** Wherever the report found code that reads a deck against the manifest:

```ts
const issues = validateDeck(template, deck, { layout: 'adaptive' }); // the loaded template, not template.manifest
```

Given the loaded template, `validateDeck` builds the deck's own layouts in, lays free slides out and checks
their text; no `withLayouts` and no `measure` option is needed any more. A slide may be free:
`isFreeSlide(slide)` tells. Its pattern is `free`, which is in no manifest, and its content is in `slots` under
`b1Text`, `b2Chart` and so on; `freeContent(slide)` gives its blocks. Code that looks a slide's pattern up in
`template.manifest.patterns` must pass such a slide by. Keep the deck JSON whole when storing it: `colors`,
`patterns`, and a slide's `blocks`, `arrange`, `base` and `colors` are part of the deck.

**5. What the chat model is told.** The tools' own descriptions tell the chat model when to fill
`accent_color`, `color_slides` and `layouts`, and to pass every other wish about how slides look in
`instructions` in the person's words. If the app's system prompt lists the tools' arguments by hand, remove
that list. If it says the chat cannot change colours, cannot put several charts on a slide or must keep to the
template's layouts, remove those sentences. Add nothing in their place unless
`node_modules/slotdeck/docs/chat-tools.md` says to.

**6. Two switches in the studio.** Where a template's settings are (its name, its language, its text size),
add two, and nothing more:

- **Presentations may use other colours.** On by default. Off writes `{ op: 'setManifest', colors: false }`, on
  writes `{ op: 'setManifest', colors: null }`. Beside it, show the template's main colour as a swatch:
  `templatePalette(template).accent`. Where the person says that is not their brand colour, let them type one:
  `{ op: 'setManifest', colors: { accent: '0E7C7B' } }`.
- **The software may lay slides out itself.** On by default. Off writes `{ op: 'setManifest', free: false }`,
  on writes `{ op: 'setManifest', free: true }`.

Both go through `changeManifest` like every other change, and what it rejects is shown with its reason. Do not
build pickers for the frame or the text styles of free slides: the software chooses them.

## Rules

- A colour a person asks for lives in the deck. Never write it into the template's pptx or manifest from the
  chat, and never read a colour out of a person's message in code of your own: the chat model passes it.
- A slide with very many parts is written in several model calls, about a minute for every fifty charts. Keep
  the progress and the stop of the chat's other calls; do not shorten their limits.
- Every number on a chart must be in what the person wrote. A chart for which nothing was given is left off
  the slide, and the tools say so in `warnings`. Show that warning.
- A presentation made before the upgrade needs nothing. Do not migrate stored decks or plans.

## Tests

No test calls a live model. Script one that answers by what the request asks for.

1. A plan whose scripted outline has one slide with `pattern: "free"` and `shows: [{ count: 4, kind: "chart" },
   { count: 1, kind: "note" }]`: the plan tool's output for that slide has `free_layout: true` and no
   thumbnail. After creating it, the stored deck's slide has `blocks` of four charts and a note, and the pptx
   has four chart parts for that slide.
2. `edit_presentation` with `accent_color: "1E6B3A"` and no instruction, with a scripted model that throws when
   called: a new presentation is made, its deck has `colors: { accent: "1E6B3A" }`, and no model was asked.
3. The app's own code from step 4 accepts a deck with a free slide and a deck with `colors`.
4. With the studio's colours switch off for a template, the tools made for it declare no `accent_color`.
5. With `free: false`, the outline request has no `shows` in its schema, and the plan is made as before.
6. A stored deck from before the upgrade renders to the same bytes as it did.

## Done when

- the report and the plan were given;
- the six tests pass;
- your final message says what a person can now ask for in the chat, what is stored where, and what you left
  as it was.
