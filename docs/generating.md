# Generating and editing decks with a model

The planner turns a brief into deck JSON by asking a language model you provide. The model never produces a
file and never positions anything: it chooses patterns from the manifest and writes text for their slots, inside
a JSON schema built from the manifest.

```ts
import { generateDeck, editDeck, planDeck, planOutline, editPlan, writeDeck, planEdit, resolveAssets, PlanError, DeckValidationError } from 'slotdeck';
import type { Brief, PlanOptions, GenerateOptions, IconDef, AssetDef } from 'slotdeck';
```

All functions on this page need an `LlmProvider` (see [providers.md](providers.md)).

## `generateDeck(template, brief, options)`: brief in, pptx out

```ts
const result = await generateDeck(template, brief, options);
// result.pptx      Uint8Array   the file
// result.deck      DeckSpec     store this: it is what gets edited and rendered again
// result.assets    Record<string, Uint8Array>   every picture the deck uses, by asset id
// result.warnings  Issue[]      limits the final draft still exceeds
// result.outline   OutlineEntry[]   [{ pattern, purpose, items? }] for each slide
```

It runs `planDeck`, then `resolveAssets`, then `renderDeck`.

### `Brief`

| Field | Type | Default | Meaning |
| --- | --- | --- | --- |
| `prompt` | string | required | What the deck is about, in the person's own words. Include every fact, figure and name that should appear: the model is told to use only what is stated. |
| `slideCount` | number, or `{ min, max }` | 6 to 10 | A fixed number of slides, or a range. |
| `language` | string | the template's | BCP 47 tag of the language to write in. |
| `direction` | `"ltr"` or `"rtl"` | the language's | Reading direction. |
| `imageDensity` | `"none"`, `"low"`, `"medium"`, `"high"` | `medium` | How many slides carry a photo: none, at most a quarter, about half, or wherever it suits. Forced to `none` by `generateDeck` when no image source and no supplied photos were given. |
| `instructions` | string | none | Anything else to respect: audience, tone, points that must be covered. |
| `fields` | object of strings | none | Deck-level values the template needs (see "Deck fields" in [deck-format.md](deck-format.md)). A pattern that needs a field that is not given is not offered. |
| `textSize` | `"normal"`, `"large"`, `"x-large"` | `normal` | Size of body text. Larger text gets smaller budgets, so less is written. |
| `textOnly` | boolean | `false` | Plan every slide as plain text. Only the template's plain text patterns are offered. See "Asking for a kind of slide" below. |

### Options (`GenerateOptions`)

| Option | Type | Default | Meaning |
| --- | --- | --- | --- |
| `llm` | `LlmProvider` | required | The text model. |
| `images` | `ImageProvider` | none | Makes photos from descriptions. Without it, decks are planned with no generated photos. |
| `icons` | `IconDef[]` | none | Your icon set: `{ id, description, data }`. Patterns with icon slots are offered only when this is given. The model picks icons by `description`, so describe what each shows and what it stands for ("a clock: time, schedule"). |
| `assets` | `AssetDef[]` | none | Pictures you already have: `{ id, kind: 'photo' | 'graphic', description, data }`. The model decides where each fits from its description. |
| `layout` | `"strict"` or `"adaptive"` | `strict` | Whether optional slots may be left out. |
| `patterns` | string[] | all | Restrict the deck to these pattern ids. |
| `numberPatterns` | `"from-brief"`, `"always"`, `"never"` | `from-brief` | When patterns marked `figures` may be used. `from-brief` offers them only if the brief contains digits. |
| `maxRepairs` | number | 2 | How many times a draft that breaks the template's limits is sent back to the model. |
| `review` | `DeckReview` | none | A second look at every draft: design checks, a model that looks at pictures of the slides, or a function of your own. What it finds goes back to the model in the repair rounds. See [design-checks.md](design-checks.md). |
| `composeLayouts` | number | 0 (off) | Up to this many new layouts may be made for the deck, in the template's own look, for slides that none of its layouts shows well. About 10 to 25 seconds more. See "New layouts for a deck" below. |
| `cache` | `Map<string, Uint8Array>` | none | Generated pictures by asset id. Pass the same map on later calls and pictures whose description did not change are reused. |
| `imageConcurrency` | number | 3 | How many pictures to generate at once. |
| `onProgress` | function | none | Called with `{ stage: 'outline' }`, `{ stage: 'layouts' }` (with `composeLayouts`), `{ stage: 'write', slides }`, `{ stage: 'repair', attempt, issues }`, `{ stage: 'images', done, total }`, `{ stage: 'render' }`. |
| `signal` | `AbortSignal` | none | Cancels the work. |

### Example

```ts
const { pptx, deck, warnings } = await generateDeck(
  template,
  {
    prompt: 'A deck for new adopters: why adopt a rescue dog, the first month at home, what it costs. We rehomed 286 dogs in 2023, 351 in 2024 and 412 in 2025.',
    language: 'en',
    slideCount: 9,
    imageDensity: 'medium',
    fields: { organizationDepartment: 'Harbor Street Shelter  |  Adoption team' },
  },
  {
    llm,
    images,
    icons: [
      { id: 'clock', description: 'a clock: time, schedule, duration', data: clockPng },
      { id: 'heart', description: 'a heart: care, health, loyalty', data: heartPng },
    ],
    assets: [{ id: 'flow', kind: 'graphic', description: 'diagram of the adoption process in four steps', data: flowPng }],
    layout: 'adaptive',
    onProgress: (event) => console.log(event.stage),
  },
);
```

## How planning works

1. **Outline** (one model call). The model sees every usable pattern with its description and slot limits, and
   answers with one pattern per slide and the point each slide makes.
2. **Write** (one model call). The model fills a JSON schema built for exactly that sequence of slides, so it
   cannot invent patterns or slots.
3. **Repair** (up to `maxRepairs` more calls). The draft is validated against the manifest. Errors and
   over-budget text go back to the model with the list of problems. For a template loaded with fonts the
   text is measured, and the problems are stated in lines ([text-fit.md](text-fit.md)). With a `review`, what
   the review finds on a draft goes back in the same list ([design-checks.md](design-checks.md)).

Which patterns are offered depends on what you supplied:

| A pattern with | Is offered only when |
| --- | --- |
| a required photo slot | `imageDensity` is not `none`. `generateDeck` and the chat tools set it to `none` when there is no image source and no supplied photo. If you call `planDeck` yourself without a way to make photos, set `imageDensity: 'none'`. |
| a required icon slot | `icons` is given |
| a required graphic slot | `assets` has a picture of kind `graphic` |
| a required deck field | `brief.fields` has it |
| `figures: true` | the brief contains digits (or `numberPatterns: 'always'`) |

"Required" follows the layout: in `strict` layout an optional slot (`empty: "remove"`) counts as required,
because nothing may be removed from a slide. Slots with `empty` of `keep` or `drop` are never shown to the
model.

If nothing can be offered, `PlanError` is thrown with the reasons.

### Asking for a kind of slide

By default the planner varies the patterns across a deck. A wish about the kind of slides overrides that, and
there are two ways to state one.

**In words.** A wish in `prompt` or `instructions`, such as "text only", "no tables" or "only cards", decides
which patterns may be used and comes before variety. The planner is told so, and the pattern list it reads marks
each plain text pattern with the words "Plain text.".

**As a flag.** `textOnly: true` is the reliable way to get a deck of text and nothing else. The planner is then
offered only the plain text patterns, so no other kind can be chosen, and it is told that using the same pattern
for several slides in a row is expected.

```ts
await generateDeck(template, { prompt: 'Our move to hybrid work: why, the rules, what changed', textOnly: true }, { llm });
// for example: cover, text-only, text-only, text-only, text-only, closing
```

A pattern counts as **plain text** when a slide made from it holds running text and nothing else: all its
slots are text (a picture that the pattern always drops does not count), it has no row or grid of items, and it
is not marked `figures`. A cover, a closing slide, a quotation and a title-with-paragraphs slide are plain text;
cards, tables, charts and anything with a picture or icon are not.

Two things to know:

- "No pictures" is a different wish. `imageDensity: 'none'` removes photos but still allows cards, tables and
  charts.
- A text-only deck is only as varied as the template's plain text patterns. A template without one cannot make
  such a deck: `PlanError: Template "x" has no plain text pattern, so a text-only deck cannot be planned from it`.
  Add one as a derived pattern (see [layout-freedom.md](layout-freedom.md)).

### New layouts for a deck

A template has the layouts it has. A text-only deck from a template with one plain text layout uses it for
every slide, and three goals or a plan in four steps end up as paragraphs. With `composeLayouts` set, the
planner makes the layouts such slides need, in the template's own look:

```ts
const plan = await planOutline(template, { prompt: 'Quarterly update: three goals and a plan in four steps', textOnly: true }, { llm, composeLayouts: 2 });
// before: cover, text-only, text-only, text-only, text-only, closing
// after:  cover, three-goal-cards, process-cards, process-cards, text-only, closing
// plan.layouts: the two designs
```

1. The outline is planned as usual.
2. One light call looks at it and names the slides whose content none of the offered layouts shows well (points
   of equal weight, steps, a comparison, an agenda, a timeline, a statement, a few key figures) and what layout
   each would need. Only slides of content may get one: a cover, a section divider, a quotation and a closing
   slide keep theirs.
3. Each layout is composed with `composePattern` (see [template-studio.md](template-studio.md)), side by side, on
   the frame of the layout its first slide had, so its title sits where the deck's other titles do.
4. The plan's slides use the new layouts, and the deck is written once, as always.

The designs travel in the plan and the deck (`layouts`, see [deck-format.md](deck-format.md)). Rendering builds
them onto the template each time, so the template file never changes, and a stored deck renders the same later
without a model. A layout that cannot be made leaves its slides as planned; the deck is made either way.
`composeLayouts(template, plan, options)` does steps 2 to 4 on a plan you already have, and says what was made
and what could not be.

What to expect, from live runs on two templates (one in Hebrew) with `gemini-3.8-flash`:

- Every text-only deck got one or two new layouts, and no layout failed to build. Rows of cards, steps with
  numbers, a timeline, a comparison, an action list with owners and key figures all came out as real slides in the
  template's styles.
- A deck with varied layouts got fewer: one in each of the two runs, for the slide that was running text.
- It adds 10 to 25 seconds and two to five light calls to a deck.
- The writer fills a composed box to about the length its design was drawn for; a card can still be roomier than
  its text.

It is off by default: it costs time, and a template whose own layouts suit its decks gains little from it. A
layout that keeps being wanted belongs in the template: add it in a studio with `composePattern`, or from the
deck's design with `buildPattern`. A deck that names that id then uses the template's own.

### Why numbers are guarded

A model given a slide built around a statistic and no data tends to invent one. So patterns marked `figures` are
held back unless the brief itself contains figures, and chart values that the brief does not state come back as
warnings such as `chart values not stated in the brief: 89`. Values worked out from the brief, such as a sum, are
flagged too; treat the warning as a prompt to check.

Contact details and names are not guarded. A model asked for a closing slide may invent an email address. State
such details in the brief, or review the deck.

## The steps on their own

Use these when a person should see and approve a plan before the content is written.

```ts
// 1. Decide the slides. One call, quick, small result: the thing to show a person.
const plan = await planOutline(template, brief, { llm, icons, layout: 'adaptive' });
// plan = { version: 1, title, brief, slides: [{ pattern, purpose, items? }], layouts? }

// 2. Change the plan by instruction. One call. Returns a new plan and one sentence on what changed.
const { plan: revised, summary } = await editPlan(template, plan, 'Add a slide about pricing before the end', { llm, icons, layout: 'adaptive' });

// 3. Write the content. One call plus repairs.
const { deck, warnings, outline } = await writeDeck(template, revised, { llm, icons, layout: 'adaptive' });

// 4. Pictures and file.
const assets = await resolveAssets(template, deck, { images, icons, assets: supplied });
const { pptx } = renderDeck(template, deck, { assets, layout: 'adaptive' });
```

`planDeck(template, brief, options)` is steps 1 and 3 in one call and returns `{ deck, warnings, outline }`.

Pass the same `icons`, `assets` and `layout` to every step: the plan was made on the assumption that they are
available.

### `resolveAssets(template, deck, options?)`

Collects the bytes of every picture a deck refers to: supplied icons and assets as they are, described pictures
from `cache` or from `images`. Returns a map from asset id to bytes, ready for `renderDeck`.

| Option | Meaning |
| --- | --- |
| `images` | The `ImageProvider` that makes described pictures. |
| `icons`, `assets` | The same sets the deck was planned with. |
| `cache` | A `Map` of pictures already made. New ones are added to it. |
| `concurrency` | How many to generate at once (default 3). |
| `onProgress`, `signal` | As above. |

The image provider is asked for a picture with the aspect ratio of the frame it goes into.

## Editing a deck

```ts
const edited = await editDeck(template, deck, 'Shorten slide 3 to two bullets and translate the deck to Hebrew', {
  llm,
  images,
  icons,
  assets,
  cache, // pictures of the earlier version; see extractAssets in rendering.md
  layout: 'adaptive',
});
// edited.deck       the new deck
// edited.pptx       the new file
// edited.assets     its pictures
// edited.summary    one sentence from the model on what it changed
// edited.rewritten  numbers of the slides in the new deck that were written or rewritten
// edited.warnings
```

An edit is two model calls. The first decides what the deck looks like after the change: which slides stay,
move, go, get rewritten or are new. The second writes only the slides that change. Reordering and deleting need
no second call.

- Slides the instruction does not touch are copied by code, so their text cannot drift.
- A rewritten slide keeps its pictures unless the instruction is about them.
- A slide given a different pattern takes its photo along.
- A change of language rewrites every slide, sets the reading direction and translates deck-level fields.
- "Make the text larger" sets `textScale` and rewrites only the slides whose text no longer fits.
- "Use three cards instead of four" changes `items` on that slide.

`planEdit(template, deck, instruction, options)` does the same without pictures or rendering and returns
`{ deck, summary, rewritten, warnings }`.

Extra options for edits (`EditOptions`):

| Option | Meaning |
| --- | --- |
| `fields` | Replaces the deck-level values, for example a translated header text. Defaults to the deck's own. |
| `direction` | Reading direction after the edit. Defaults to the direction of the deck's language. |
| `imageDensity` | Whether slides added by the edit may use photo patterns. Default `medium`. |

Changes made by hand in PowerPoint are not read back. The deck JSON is the source of truth.

## Errors

| Error | When | What to do |
| --- | --- | --- |
| `PlanError` | No pattern can be used with what was supplied; or the model's outline or edit plan is unusable; or the plan uses patterns that cannot be filled | Read the message: it names the patterns and reasons. Supply the missing icons, fields or images, or retry. |
| `DeckValidationError` | After all repairs the draft still has error-level issues | `error.issues` lists them. Retry, raise `maxRepairs`, or simplify the brief. |
| Anything your provider throws | The model call failed | Handled by you: slotdeck does not retry model calls. |

Warnings (text a little over a budget, for instance) do not throw. They are returned with the result.

## Cost and time

For one deck: two text calls, more if repairs are needed, and one image call per generated photo. Measured on
the demo template with `gemini-3.8-flash`: a plan takes about 5 seconds, an eight-slide deck about 30 seconds
from brief to file (with drawn stand-in pictures), and an edit that rewrites one slide about 8 seconds.

Planning calls (the outline, a change to a plan, the first step of an edit) ask the model for light reasoning,
because they choose and arrange rather than write. This made them several times faster and removed the runaway
answers described in [providers.md](providers.md). Writing the slides keeps the model's default.

To keep cost down: reuse `cache` across calls, set `imageDensity: 'low'` or `'none'`, and use `planOutline`
first so a person can correct the plan before content and pictures are produced.
