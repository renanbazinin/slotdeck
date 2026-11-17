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
// result.outline   OutlineEntry[]   [{ pattern, purpose, items?, words? }] for each slide; a free slide has blocks and arrange too
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
| `textOnly` | boolean | `false` | Plan every slide as plain text. Only the template's plain text patterns are offered, and free slides of text. See "Asking for a kind of slide" below. |
| `patterns` | string[] | none | The only patterns the person asked for, by id, `free` among them for free slides. Every slide then uses one of them. Kept with the plan and the deck; a later change keeps to it unless it asks for another pattern. |
| `colors` | `DeckColors` | none | Colours in place of the template's; copied to the deck. See [colors.md](colors.md). |

### Options (`GenerateOptions`)

| Option | Type | Default | Meaning |
| --- | --- | --- | --- |
| `llm` | `LlmProvider` | required | The text model. |
| `images` | `ImageProvider` | none | Makes photos from descriptions. Without it, decks are planned with no generated photos. |
| `icons` | `IconDef[]` | none | Your icon set: `{ id, description, data }`. Patterns with icon slots are offered only when this is given. The model picks icons by `description`, so describe what each shows and what it stands for ("a clock: time, schedule"). |
| `assets` | `AssetDef[]` | none | Pictures you already have: `{ id, kind: 'photo' | 'graphic', description, data }`. The model decides where each fits from its description. |
| `layout` | `"strict"` or `"adaptive"` | `strict` | Whether optional slots may be left out. |
| `patterns` | string[] | all | Restrict the deck to these pattern ids. A list that does not name `free` gets no free slides. |
| `numberPatterns` | `"from-brief"`, `"always"`, `"never"` | `from-brief` | When patterns marked `figures` may be used. `from-brief` offers them only if the brief contains digits. |
| `maxRepairs` | number | 2 | How many times a draft that breaks the template's limits is sent back to the model. |
| `review` | `DeckReview` | none | A second look at every draft: design checks, a model that looks at pictures of the slides, or a function of your own. What it finds goes back to the model in the repair rounds. See [design-checks.md](design-checks.md). |
| `composeLayouts` | number | 0 (off) | Up to this many new layouts may be made for the deck, in the template's own look, for slides that none of its layouts shows well. About 10 to 25 seconds more. None is made while free slides are on offer. See "New layouts for a deck" below. |
| `free` | boolean | `true` | Whether a slide may be a free slide: one the software lays out itself from the parts the plan names, for content that no pattern shows. `false` turns them off. See "Free slides" below. |
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
   cannot invent patterns or slots. A free slide of more than nine parts is written by calls of its own (see
   "Free slides" below).
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

A free slide is offered beside the patterns, under conditions of its own: see "Free slides" below.

If nothing can be offered, `PlanError` is thrown with the reasons.

### Asking for a kind of slide

By default the planner varies the patterns across a deck. A wish about the kind of slides overrides that, and
there are three ways to state one.

**In words.** A wish in `prompt` or `instructions`, such as "text only", "no tables" or "only cards", decides
which patterns may be used and comes before variety. The planner is told so, and the pattern list it reads marks
each plain text pattern with the words "Plain text.".

**As a list.** `patterns` on the brief limits the deck to the patterns named. Every slide uses one of them, the
first and the last too. The list is kept with the plan and the deck, and a later change keeps to it unless it
asks for another pattern.

**As a flag.** `textOnly: true` is the reliable way to get a deck of text and nothing else. The planner is then
offered only the plain text patterns, so no other kind can be chosen, and it is told that using the same pattern
for several slides in a row is expected. Where free slides are on offer, a free slide may be planned too, of
texts, lists, quotations and notes alone.

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
- A text-only deck is only as varied as the template's plain text patterns. A template without one makes such a
  deck from free slides of text. Where those are not on offer either, it cannot:
  `PlanError: Template "x" has no plain text pattern, so a text-only deck cannot be planned from it`. Add one as a
  derived pattern (see [layout-freedom.md](layout-freedom.md)).

### Asking for an amount of text

A person may ask for an amount of text on a slide: "200 words on the expansion", "a paragraph of about 80
words". Every slot says how much to write, in its description ("About 420 characters in all") and its limit,
and without more the writer keeps to that: 200 words asked came out as 76. From 0.6.1 the planner reads such a
wish and makes room for it.

1. The outline gives the slide `words`: the number of words asked. A slide with no such wish has none, and is
   planned and written as before.
2. Code finds the text a slide that holds it, counting about 6 characters to a word with its space, 6.5 in
   Hebrew (200 words are about 1,200 characters), at the deck's text size:
   - the layout planned, when nine tenths of the limits of its running text (paragraphs, lists, texts of 200
     characters or more) hold it. A layout's limits were found for the template's own text, and a text that
     fills them in many paragraphs is the one most likely to run over its box;
   - else, for a slide of plain text, the smallest layout of plain text that holds it;
   - else a free slide, where free slides are on offer, set out for that length: one text alone that long is
     set at the template's running-text size in two columns where it needs them, and its budget is what its
     box then holds. A slide of a layout with a chart or a table keeps it beside the text.
3. Where none holds it and the slide shows text alone, the text goes over as many slides as it needs, while
   the deck may have more (`slideCount` up to its most, or ten more in a change), each with its share of the
   words, counted in whole words, and its purpose numbered "(1/3)". A slide with a chart, a table, a photo or
   cards is not repeated for its text: every copy would ask for that content again. Such a slide, and one in a
   deck that may have no more slides, holds what it can.
4. The writer is told the length before the rest of the slot's description, from which what it said of length
   is taken out, in words and as a range of characters that starts at the length:
   `Write 200 words, as was asked: 1200 to 1370 characters in all`. A model writes short of a length it is told
   "about", by a fifth. A long text with nothing else to say of itself is asked for in paragraphs. The slot's
   limits stay as they are. A slide that holds less is told what it holds, and the deck gets a warning:
   `200 words were asked for this slide, and it holds about 92: ask for the text over more slides, or for a shorter one`.

```ts
const plan = await planOutline(template, { prompt: 'Our expansion north', instructions: 'One slide of 200 words of running text.', slideCount: 1 }, { llm });
// plan.slides = [{ pattern: 'free', purpose: 'Our expansion north', words: 200, blocks: ['text'] }]
```

`words` is part of a plan an app keeps, from 1 to 3000. A change to the plan keeps it unless the request changes
it, and `writeDeck` writes a plan that has it as above. A change to a deck reads it the same way: see "Editing a
deck" below.

### New layouts for a deck

A template has the layouts it has. A text-only deck from a template with one plain text layout uses it for
every slide, and three goals or a plan in four steps end up as paragraphs. With `composeLayouts` set, the
planner makes the layouts such slides need, in the template's own look. Free slides answer the same need, and a
deck gets one of the two: while free slides are on offer, which they are unless `free` is `false`, no layout is
made.

```ts
const plan = await planOutline(template, { prompt: 'Quarterly update: three goals and a plan in four steps', textOnly: true }, { llm, composeLayouts: 2, free: false });
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

It is off by default: it costs time, and a template whose own layouts suit its decks gains little from it. And
it needs `free: false`, since a free slide is otherwise what such content gets. A
layout that keeps being wanted belongs in the template: add it in a studio with `composePattern`, or from the
deck's design with `buildPattern`. A deck that names that id then uses the template's own.

### Free slides

A pattern has the places it was drawn with: three cards, one chart. Content that no pattern of the template
shows (six charts on one slide, twelve cards, a table beside two figures) gets a **free slide**, which the
software lays out itself from the parts the plan names. What such a slide is and how it is laid out is in
[free-slides.md](free-slides.md). The planner offers free slides by itself:

```ts
const plan = await planOutline(template, { prompt: 'Sales by quarter for each of our six stores: ...', fields }, { llm });
// plan.slides[1] = { pattern: 'free', purpose: 'Every store at a glance',
//                    blocks: ['chart', 'chart', 'chart', 'chart', 'chart', 'chart', 'note'], arrange: { columns: 3 } }
```

- The list of patterns the outline reads ends with one line for `free`. It names the kinds of part a slide may
  have, says that a slide a pattern shows well uses that pattern, and says about how many parts read well at the
  template's text size.
- A slide of the answer that chooses `free` names its parts in `shows`, like parts together, one entry for each
  run with its `count`: `"shows": [{ "count": 6, "kind": "chart" }, { "count": 1, "kind": "note" }]`. It may say
  how many stand side by side (`side_by_side`), and where the first part stands when it is the main one
  (`shows_main`). The plan keeps them as `blocks` and `arrange`, the two fields a free slide adds to an entry.
  Entries of one kind that follow one another and count up from 1 (one entry for each chart, numbered) are read
  as that many parts.
- A number of things the brief asks to see on one slide is that number of parts, however many: "all twelve
  store charts on one slide" is twelve charts.
- The kinds on offer follow what the deck has. Figures, charts and tables need digits in the brief
  (`numberPatterns`), a photo needs an `imageDensity` other than `none`, a graphic needs a supplied asset of kind
  `graphic`, and a card is offered an icon only with `icons`. A text-only deck is offered texts, lists,
  quotations and notes.

Free slides are on offer when all of this holds:

| Condition | Otherwise |
| --- | --- |
| The option `free` is not `false`. | `free: false` plans without free slides: the prompts and schemas say nothing of them. |
| The template can carry them: its manifest does not say `"free": false`, it has a layout with a title for them to stand on, that layout leaves room under the title, the template's running text is in text boxes of its own and not only in placeholders, and `brief.fields` has the deck fields that layout shows. See "What a template gives a free slide" and "Limits" in [free-slides.md](free-slides.md). | No free slide is offered. |
| No list of patterns is in force, or the list names `free`. That goes for the app's list (`patterns` in the options) and for the person's (`Brief.patterns`). | An app that passes `patterns: ['cover', 'text-only']` gets no free slides by upgrading; it adds `'free'` to have them. |

A new brief whose own list names `free` where none can be made is a `PlanError` that says why, such as
`The deck cannot be limited to a pattern that cannot be filled here: free (free slides are turned off)`.

While free slides are on offer, `composeLayouts` makes nothing: a free slide is what a deck gets for content
that no layout shows. Pass `free: false` to have layouts made in their place.

The option decides what a new slide is offered and nothing else. A plan or a deck that already has a free
slide is written, changed and rendered whatever the option says.

**Changing a plan.** `editPlan` lists a free slide with its parts, such as
`2. free (6 charts and a note; 3 side by side): Every store at a glance`, and tells the model to repeat them
unless the request changes what the slide shows. A free slide whose parts come back empty keeps the parts it
had. However its parts come back, the same or others, the slide stays set out as it was (three side by side,
the main part at the left) wherever the answer does not say otherwise.

**Writing.** A free slide is written against slots worked out for its parts. What the writer reads of a slot
is the slot's own words, then its limits, each said once:
`The one line to take away. Write 80 to 150 characters. at most 153 characters` for a note,
`A small chart. When the brief gives no numbers for it, leave it empty. at most 6 categories, at most 2 series`
for a chart, followed by the rule of every chart. The planner takes nothing from a slot's description and puts
nothing into it. A text with too little room for a sentence on how much to write has no description, and
reads `at most 22 characters`. See "What a slot says before the slide is written" in
[free-slides.md](free-slides.md).

- Up to nine parts: the slide is written with the other slides, in the one writing call.
- More than nine: the slide is written by calls of its own, with an array for each run of like parts in place
  of a property for each slot. The answer of one such call is capped at 24,000 tokens, and its schema has at
  most twelve arrays. A slide that needs more is cut into consecutive ranges of parts (about 40 small charts
  each), written three calls at a time, and each call is told that the other parts are written apart. Where
  like parts stand before its own on the slide, a call is also told which of their kind its parts are:
  `- r01: parts 31 to 60, 30 charts (the 31st to the 60th chart of the slide)`. A call of charts and tables
  alone asks the model for light reasoning.
- An array says of its parts what the smallest of them holds: the description of that slot, and of each limit
  the lowest that a part of the array states. A chart whose slot shows one series (a chart drawn as its bars
  or its line alone) is asked for as one list of numbers, named by its caption, and is told no number of
  series. The room a call's answer gets follows the same limits: a chart is counted as long as the categories
  and series its slot says it shows.
- A call whose answer has problems is asked again by itself. What it then returns is laid over its answer
  before: an answer that holds only the items that were to be mended loses none of the others. The same goes
  for a free slide that is written with the deck.
- Text far over the budget the writer was shown (by more than a quarter) goes back for repair, like any text
  over a slot's budget. A budget is what the slide holds when its texts are about as long as texts of their
  kind usually are, so it is one to write to, on a slide of sixteen cards as on a slide of three. Only a budget
  under 12 characters is not held to in this way: that is a word or two, the value of a figure, and the check
  of the deck reports what does not fit there.
- Every photo a writer describes is a picture to generate, and at most six are made for one slide. A slide
  that was planned with more (it may be, where photos were supplied) tells the writer so, a seventh
  description goes back for repair, and what is still over after the last repair is left out:
  `slide 3: 1 photo was left out, since at most 6 are made for one slide`.
- Every number a chart shows must appear in the brief; the check does not tell whether it stands in the right
  chart. A chart and a table are asked of the writer like any part, and may be given empty: in a slide written
  with the deck their lists say how (`[] for a chart that is left empty`), and in a call of arrays every
  property may be `""` or `[]`. A chart, a table or
  a figure the brief has no numbers for comes back empty and is taken off the slide, without a repair round; a
  table with headings and no rows counts as empty. A part whose content is still missing after the last repair
  is taken off too. The parts
  that stay are numbered anew, and the deck is told in a warning without a slide number:
  `slide 3: 13 parts had no content and were left out`. A slide left with no part says its purpose as one
  text; where the plan gave it none, the title that was written for it, or else the title of the deck, or
  else the first words of the brief. A free slide never makes `writeDeck` throw for missing content.

What it costs: about a minute of model time for every 50 charts. A slide of 200 charts is five calls, made in
two rounds: three side by side, then two.

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
// plan = { version: 1, title, brief, slides: [{ pattern, purpose, items?, blocks?, arrange?, words? }], layouts? }

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

A plan is plain JSON that an app may keep and show for editing, so `writeDeck` checks its counts before it
asks a model anything. `items` of a slide has to be a whole number the pattern's row takes, and a free slide
has at most 400 parts. Otherwise it throws a `PlanError` that names the slide:
`Slide 2: pattern "stat-cards" takes 2 to 6 stats, not 7` or
`Slide 2: a free slide takes at most 400 parts, not 2000`.

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

`resolveAssets` does not check the deck. A slide whose pattern the template does not have, or whose `items` is
a number its row does not take, is passed over, and no picture is made for it: `renderDeck` then says what is
wrong with it.

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
- "Make slide 2 about 200 words" gives that slide the room for the text, as a plan does (see "Asking for an
  amount of text" above): its layout, another layout of plain text, or a free slide written from what the slide
  held, which keeps the slide's chart, table or photo as it is. Where nothing holds it and the slide shows text
  alone, the text goes over new slides after it, of the same kind, and a warning without a slide number says
  where it went:
  `the 200 words asked for slide 2 go over slides 2 to 4: one slide holds about 55`.
- A free slide is changed part by part: the parts the instruction does not touch stay exactly as they are. See
  "Free slides in an edit" below.

`planEdit(template, deck, instruction, options)` does the same without pictures or rendering and returns
`{ deck, summary, rewritten, warnings }`.

### Free slides in an edit

A deck that has a free slide is changed like any deck, and an edit can make one: a new slide, or a slide of a
pattern that is to hold more than its pattern shows ("add two more charts to slide 3"). For a new slide, free
slides are on offer under the conditions of "Free slides" above, with the list the deck is limited to
(`DeckSpec.patterns`) in the place of the brief's. A free slide the deck already has is listed, changed and
kept whatever is on offer.

```ts
const edited = await planEdit(template, deck, 'Add charts for the two new stores to slide 3: Airport 10, 12, 14, 18 and Harbour 18, 14, 12, 10', { llm });
// edited.deck.slides[2].blocks   eight charts and the note, where there were six charts and the note
// edited.rewritten               [3]
```

- The first call is shown a free slide by its parts, numbered and without their data:
  `3. [free: 6 charts and a note; 3 side by side] title: "Sales by store"; 1 chart "North" (column, 4 categories); ...; 7 note "Both grew."`.
  A slide of more than 24 parts is listed in short: the first three and the last of each run of like parts,
  and how many lie between.
- Its answer names the parts the slide has after the change, one by one: which existing part each one is, or
  that it is new, and whether its content is to be written again. It also says how the parts are set out. A
  part the answer does not name is taken off the slide; an answer that leaves most of a slide's parts unnamed
  is asked about once before it is done. A part that is named by its number with a kind that is no kind is the
  part of that number, and stays what it is.
- **A part that stays is carried over by code**, under its new number. No model is shown its content or asked
  for it again: adding two charts to a slide of sixty cannot change a number in the other sixty.
- A part taken off, parts in another order and another arrangement (two side by side, the first part on top)
  are done by code too. Such an edit writes nothing, and the slide is among `rewritten` all the same.
- The plan is told to leave its words for the slide (`change`) empty when parts are only taken off or moved.
  Words it gives all the same, with no part to write, can only be about the title, which the plan has no other
  way to name: one small call then writes the title alone. That is how "take the note off and call the slide
  Sales 2025" gets its new title, and how a plan that names every part as it is and gives a new title changes
  the title and no chart.
- A plan may have words that are about a part and still mark no part to be written again ("rename the chart of
  Old Mill", with every part named as it is). The title is then all the writer is offered, and it comes back as
  it was: nothing on the slide changed. The warnings say so, without a slide number, and no call is added:
  `slide 3: nothing on it was changed: the plan of the change named no part to write again`. The slide is then
  not among `rewritten`. Ask again, naming the part.
- Wherever the title of a slide that has one is offered (with parts that are written, or alone), the writer is
  told to repeat the current title unless the request is about it. A title that is left out stays. An empty
  title takes it off a slide of up to nine parts; on a larger slide, which is written by calls of its own, it
  keeps the title the slide had, so an edit changes the title of such a slide and does not remove it.
- The plan is told to repeat how a slide is set out unless the request is about it. A value it states (`left`,
  `right`, `top`, `none`; a number side by side of 1 or more) is always set. `auto` and 0 hand a slide back to
  the software's own choice only on a slide the plan changes: one it has words for, or whose parts change. On a
  slide the plan leaves alone they state nothing, since they are also what a model answers for a slide it has
  nothing to say of, and the slide is copied as it is. An answer that leaves the two properties out changes
  nothing on any slide. `editPlan` is more careful still: there `auto` and 0 never undo an arrangement.
- A part that is to be written again is shown to the writer with what it holds, and keeps its pictures unless
  the instruction is about them. A chart with several series is shown and asked for with all of them, also on a
  slide of so many charts that a new one is asked for as one list of numbers. Where nothing comes back for the
  part, or nothing that can be shown, it stays as it was: an edit takes no part off a slide by leaving it
  unsaid. The title likewise. What the plan asked for outright and did not come is asked for again in the
  repair rounds, and when it still has not come the warnings say that the change was not made:
  `slide 3: nothing was written for 1 part, and it stays as it was`. A new part that comes back without
  content is left out, as in `writeDeck`, and the warnings say so.
- The slide is planned again with the optional parts it has. By its rules the software plans a caption for the
  photos of a slide that has several, and a heading over the texts of a slide of several parts. A slide may
  have either all the same: a caption under its one photo, a heading over its one text. The change counts
  those in. They keep their room on the slide, and one whose part is written again is offered again with what
  it says, so that a translation reaches the heading too. A photo that is replaced is offered its caption with
  the new picture; on a slide of several photos a caption is offered also where the photo had none.
- A chart or a table the slide holds may be larger than its place reads well at: three series where a new
  chart of that size is told two. It was written when the slide had fewer parts, or by code. It is offered as
  it is: the limits the writer reads, and those the schema enforces, are no lower than what the part holds, so
  that a change that is not about its numbers costs it no series and no row. Such a part is not sent back for
  its size when it comes back no larger than it was. The warnings still say
  `the chart has 3 series; the slot fits 2`, as a check of the deck says before the change. A part that comes
  back larger than it was is sent back like any chart over its limits.
- Like charts of a slide are drawn on one value axis. New numbers for one chart can therefore change the scale
  the charts beside it are drawn on. Their numbers in the deck are not touched.
- When the instruction changes the content of the slide and the plan names no parts, every part is offered to
  the writer with what it holds. A part that is offered so may be left out of the answer and keeps what it had:
  on a slide written by calls of its own, an array that holds the changed parts alone is taken as it is, and no
  repair round is spent on the number of its items. The number is a problem only for parts that are new or that
  the plan named to be written again, and the repair names those. An item of such a short array that does not
  say which part it is is not given to a part that may stay.
  A change of language writes every part again, whatever the plan says of them.
- A chart or a table is asked of the writer in an edit as in `writeDeck`: required, and empty where there is
  nothing for it. One the slide has and that comes back empty keeps what it had.
- A slide of a pattern that becomes a free slide hands its chart and its table to the first new parts of
  those kinds that the plan does not ask to be written again, and its photos to the first photo parts unless
  the instruction is about pictures. The plan is told to mark such a part as not to be written again unless
  the request changes what it shows. A chart it marks to be written is written from the content of the slide,
  which the writer is then shown whole; the other texts of the slide are the writer's material in both cases.
  A free slide that takes a pattern is laid out again like any slide, and takes its first photos along.
- A new free slide whose parts all come back empty says the title that was written for it, or else the title
  of the deck: the words of the plan say what to do with the slide, and are not shown on it.
- A slide of up to nine parts is written with the other slides. A larger one is written by calls of its own, as
  in `writeDeck`, for the parts that change and no others.
- "Make the text larger" writes again only the parts of a free slide whose text no longer fits.

The chart numbers of a free slide are checked against the brief, and warned about, for the charts an edit wrote.
A chart the slide kept is not looked at again.

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
| `PlanError` | No pattern can be used with what was supplied; or the model's outline or edit plan is unusable; or the plan or the edit uses patterns that cannot be filled, or has a free slide the template cannot carry; or a plan given to `writeDeck` has a number of items its pattern does not take, or a free slide of more than 400 parts | Read the message: it names the patterns and reasons, or the slide and its number. Supply the missing icons, fields or images, correct the plan, or retry. |
| `DeckValidationError` | After all repairs the draft still has error-level issues | `error.issues` lists them. Retry, raise `maxRepairs`, or simplify the brief. |
| Anything your provider throws | The model call failed | Handled by you: slotdeck does not retry model calls. |

Warnings (text a little over a budget, for instance) do not throw. They are returned with the result.

## Cost and time

For one deck: two text calls, more if repairs are needed, and one image call per generated photo. Measured on
the demo template with `gemini-3.8-flash`: a plan takes about 5 seconds, an eight-slide deck about 30 seconds
from brief to file (with drawn stand-in pictures), and an edit that rewrites one slide about 8 seconds. A free
slide of many charts takes about a minute of model time for every 50 charts, in calls that run three at a time.

Planning calls (the outline, a change to a plan, the first step of an edit) ask the model for light reasoning,
because they choose and arrange rather than write. This made them several times faster and removed the runaway
answers described in [providers.md](providers.md). Writing the slides keeps the model's default.

To keep cost down: reuse `cache` across calls, set `imageDensity: 'low'` or `'none'`, and use `planOutline`
first so a person can correct the plan before content and pictures are produced.
