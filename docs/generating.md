# Generating and editing decks with a model

The planner turns a brief into deck JSON by asking a language model you provide. The model never produces a
file and never positions anything: it chooses patterns from the manifest and writes text for their slots, inside
a JSON schema built from the manifest.

```ts
import { generateDeck, editDeck, planDeck, planOutline, editPlan, writeDeck, planEdit, planTextFit, resolveAssets, PlanError, TextCutError, DeckValidationError } from 'slotdeck';
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
| `document` | string | none | A document to make the deck from, word for word: a report, minutes, an article. The slides carry its own text, every part of it, in its order. `prompt` then says what the deck is for, and `instructions` any wish ("only the risks", "a short summary"). At most 12,000 characters (`MOST_DOCUMENT_CHARS`). Copied to the deck. See "Making a deck from a document" below. |

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
| `countSpacing` | boolean | `true` | Whether the writer of a layout's slot of paragraphs is told what its box holds, the template's own spacing counted. `false` tells every slot its stored budget, as 0.6.7 did. See "What a writer is told a layout holds" below. |
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

### What a writer is told a layout holds

A slot's budget (`maxChars`, `maxParagraphs`) is counted once, for the template's own text, without the space
the template keeps around paragraphs. A text that fills it in many paragraphs takes that space many times, and
can run past its box while it keeps to the budget. From 0.6.8 the writer of a slot of two paragraphs or more is
told what its box holds, counted without fonts from the template. This applies where the template keeps space
before or after the slot's paragraphs, or sets their lines further apart than 1.2 times their size:
- the lines the box holds at the slot's size, the deck's text size counted;
- less the space PowerPoint draws between paragraphs, and before the first and after the last where the shape
  says so (`spcFirstLastPara`);
- each paragraph counted in whole lines, each line 85% full, characters as wide as `characterWidth` or the
  estimate.

So a text of any paragraph lengths that keeps to what it is told fits. The writer is told one pair:
- the stored budget where that fits;
- else the same paragraphs with fewer characters, where that costs few;
- else fewer, longer paragraphs.

The pair stands in the slot's limits ("at most 357 characters in total, at most 5 items"), in the outline's list
of layouts, and as the writer's `maxItems`. A length the slot's description states as a clause of its own
("About 420 characters in all") is left out where the pair is less. The deck is still checked against the
stored budget, so a writer is not sent back for a text between the two. `countSpacing: false` turns the count
off.

### Asking for an amount of text

A person may ask for an amount of text on a slide: "200 words on the expansion", "a paragraph of about 80
words". Every slot says how much to write, in its description ("About 420 characters in all") and its limit,
and without more the writer keeps to that: 200 words asked came out as 76. From 0.6.1 the planner reads such a
wish and makes room for it, and from 0.6.2 the text stays on the slide it was asked for.

1. The outline gives the slide `words`: the number of words asked. A slide with no such wish has none, and is
   planned and written as before.
2. Code finds the text a slide that holds it, counting about 6 characters to a word with its space, 6.5 in
   Hebrew (200 words are about 1,200 characters), at the deck's text size:
   - the layout planned, when its running text (paragraphs, lists, texts of 200 characters or more) holds it. A
     slot the count above speaks for holds what its box holds in some number of paragraphs, and its writer is
     told that number and what it holds. Any other slot holds nine tenths of its limits: they were found for the
     template's own text, and a text that fills them in many paragraphs is the one most likely to run over its
     box;
   - else, for a slide of plain text, the smallest layout of plain text that holds it;
   - else a free slide, where free slides are on offer, set out for that length: one text alone that long is
     set at the template's running-text size in two columns where it needs them, and its budget is what its
     box then holds. A slide of a layout with a chart or a table keeps it beside the text;
   - else the same free slide set dense: its text alone, set smaller down to the least size that reads (9 pt on
     a slide 13.33 inches wide, 7 pt on one 10 inches wide, or the template's `leastTextSize`), in two or three columns in the larger room under
     the title. A photo gives way. So do the texts of a plan's slide of several texts, where nothing is written
     yet and they would be cut: they are one text (`planOutline`, `editPlan`). In a plan, so do a list, a
     chart, a table or cards that leave the slide less than half of the text asked, where the text alone holds
     more (from 0.6.3). In a change to a plan (`editPlan`) only the parts the plan had give way: a part a
     change asks for stays, in that change and the ones after it, where it is a part of a free slide or a
     layout's chart or table. A second text of a deck's slide does not give way (`planEdit`).

   A free slide of one text that carries an asked length is always marked dense (`arrange.dense`), so that the
   paragraphs the writer chooses never overflow it; where its text fits at four fifths of the template's size
   it is set as large as it fits, as without the flag. From 0.6.3 one rule says whether it fits: its
   paragraphs counted by length one by one, its columns filled as PowerPoint fills them, each keeping a line
   free. See "A long text a person asked for" in [free-slides.md](free-slides.md).
3. A slide asked for words is never spread over more slides. A text longer than a slide holds at the least
   size is cut to what it holds, and the plan says so before anything is written (`planTextFit`, below). A
   slide that keeps a chart, a table, a list or cards beside which it holds half of its text or more, or a
   deck's second text, holds what it can beside them. So does a slide that keeps such parts in a change to a
   plan (`editPlan`) where a change asked for them, or in a change to a deck (`planEdit`), whatever they hold:
   a free slide's parts, and a layout's chart or table. A layout of cards or of a list gives way to the text,
   as before.
4. The writer is told the length before the rest of the slot's description, from which what it said of length
   is taken out, in words and as a range of characters that starts at the length, or at 95% of what the slot
   holds where that is less: `Write 200 words, as was asked: 1200 to 1370 characters in all`. A model writes
   short of a length it is told "about", by a fifth. A long text with nothing else to say of itself is asked
   for in paragraphs, at most 12 on a dense slide, whose range ends where the text fits in any number of them
   of any length. The slot's limits stay as they are, except on a slot whose box the count of its template's
   spacing speaks for: its limits and `maxItems` are then what it holds in the paragraphs it is told (see
   "What a writer is told a layout holds" above). A text written under 85% of the least it was told is
   sent back once; still short, the deck gets a warning: `500 words were asked for this slide, and it has
   about 380: ask for it again to have them all`. A slide that holds less is told what it holds (`Write 446
   words, 2610 to 2900 characters: of the 1000 words asked for this slide, the most it holds`), and the deck
   gets a warning: `200 words were asked for this slide, and it holds about 92: ask for the text over more
   slides, or for a shorter one`, or on a slide set dense, `1000 words were asked for this slide, and it holds
   about 446 even with its text set small, at 7 pt: the rest is cut. Ask for the text over more slides, or for
   a shorter one`.

With free slides off a slide holds what its layout holds: no text is set smaller, no slide is added, and the
rest of a long text is cut, which the plan says before writing. From 0.6.3 the deck's warning says why, at its
end: `200 words were asked for this slide, and it holds about 55: ask for the text over more slides, or for a
shorter one; free slides are off, so its text is not set smaller`. So do the plan tool's `note` and the
`TextCutError` of a change that stops before writing: `Free slides are off, so the text is not set smaller.`
Free slides are off where `free: false` is passed, the manifest says `"free": false`, or a list of patterns
(the app's or the person's) does not name `free`.

What to expect: a slide 10 inches wide of a right-to-left template holds at most about 2,900 characters at 7
pt, some 440 Hebrew words: 500 of them (about 3,250 characters) are cut to that. 500 English words (about
3,000 characters) go on one slide 13.33 inches wide at 10 pt, which holds at most about 4,350 characters, some
720 words, at 9 pt. A text of many paragraphs holds less (see "A long text a person asked for" in
[free-slides.md](free-slides.md)).

```ts
const plan = await planOutline(template, { prompt: 'Our expansion north', instructions: 'One slide of 200 words of running text.', slideCount: 1 }, { llm });
// plan.slides = [{ pattern: 'free', purpose: 'Our expansion north', words: 200, blocks: ['text'], arrange: { dense: true } }]
```

`words` is part of a plan an app keeps, from 1 to 3000. A change to the plan keeps it unless the request changes
it, and `writeDeck` writes a plan that has it as above. `editPlan` shows each such slide with its words, the size
its text is set at where that is small, and the most it holds (`2. free (a text; 500 words; text at 7 pt; at
most about 440 words): ...`), so that "make it longer" stays within what the slide holds. It adds no slide for a
text unless the request asks for more slides, and "keep it to 2 slides" gives the two slides the words of
those they take in. A change to a deck reads `words` the same way: see "Editing a deck" below.

### Making a deck from a document

A person who gives a report, minutes or an article and asks for slides wants that text on them, not a summary
of it. Pass it as `document` (from 0.6.2), and keep `prompt` to what the deck is for:

```ts
const plan = await planOutline(template, { prompt: 'A briefing for the team on the office move', document: memo, language: 'en' }, { llm });
// for example: plan.slides[1] = { pattern: 'free', purpose: 'Why we move', blocks: ['text'], arrange: { dense: true }, passages: [1, 2, 3] }
const { deck } = await writeDeck(template, plan, { llm });
// deck.document is the memo, and deck.slides[1].passages is [1, 2, 3]
```

**How the text is split.** Code splits the document into numbered passages, p1 to pN, the same way every time
(`documentPassages`):

- A line ends a paragraph. A line of 60 characters or more that ends with no stop goes on in the next one: a
  paragraph broken where a page was.
- A line that starts with `-`, `•`, a number or a letter followed by `.` or `)` is an item; items in a row are
  one list. An `o` or a `§` followed by a tab is a bullet too: a word processor writes the bullets of a list
  within a list so, and a slide shows them as `◦` and `▪`. A line with a tab between two of its cells, or with
  two or more `|`, is a row; two rows or more on lines one after another are a table. A tab at the start of a
  line only sets it in, and a tab after an item's marker is the item's, as text pasted from a word processor has
  them. Tabs that set in every row of a table are no cells; one more, as the empty corner of a table pasted from
  a spreadsheet, is an empty cell.
- A line that starts with `#`, or a short line on its own (at most 80 characters and 12 words, with no stop at
  its end, under 60 characters or with a blank line under it), is a heading over the passages after it. A line
  right under a line of text with no stop at its end is no heading, and nor is the first of three short lines or
  more with no stops, one under another: such a run, as a list typed without bullets, is text, all of it. A
  short line over one other is a heading over it. Headings are not passages.
- Short lines of one section (each under 120 characters) go together, up to 500 characters. A paragraph or a
  list over 500 characters is cut at the ends of its sentences, or between its items, into even pieces. A table
  is never cut. A document is cut into at most 150 passages.

**The outline places them.** The outline is shown the document in its passages, with their sizes, and gives
each slide those it carries (`OutlineEntry.passages`): every passage on one slide, in the document's order
unless the brief asks for another. An opening, a closing or a section slide carries none. A passage goes in
`left_out` only when the brief asks to leave that part out, or when it only asks for the presentation ("make two
slides from this:"); what such a passage asks is read as part of the brief. Code puts right what the answer
missed: a passage on two slides stays on the first, one on no slide joins the slide of the passage before it,
and an answer that places more than half of the document nowhere is asked again once. One that still places
none has the passages shared over its slides of content; with no number of slides and no such slide, over
slides added for them, never more than the document has passages. The outline also says the language the
document is written in, which the plan keeps (`documentLanguage`). A language the text's letters belie, with
hardly any of them (under a tenth) in its script, is not kept: a Hebrew text said to be English. The script most
of the letters are in then decides. A Hebrew text full of English names and terms, said to be Hebrew, is Hebrew.

**How many slides.** With no number of slides, the document takes as many as it needs. The outline is told
what one slide holds at the template's own text size: nine tenths of what a free slide of one text holds there,
or else the room of the largest layout of plain text. If the text takes K such slides, the outline may answer
with 1 to K + 4 slides, or up to 10 where that is more. A slide given more than one slide holds then goes over
more slides, each with its own run of passages and its purpose numbered "(1/3)", while the deck may have more:
up to 30 slides, or the most of a `slideCount` range. The table of a slide, a free slide's or a layout's, goes
to the part that carries the document's table. A number of slides is fixed by a `slideCount` of one
number, by a number of slides the brief names ("2 slides"), or by "each part on its own slide". Without a
`slideCount` the outline answers `asked_slides` with the number it read, and an answer that does not have that
many slides is asked again once. With a fixed number no slide is added and nothing is left out to make room:
the passages are shared over the slides that carry them, each set dense where it needs it (see "Asking for an
amount of text" above). A slide of several texts that would be cut is one text, set dense. Where one still holds
less, the passages are shared again, in order, over the slides that carry them, and the table of a slide goes
with the document's table. Only where one is cut even so do the slides of text that carry none take their share
too (free slides of texts and photos, and plain text layouts), so that no slide stands empty while another is
cut; a chart or any other slide keeps what it was planned for. A slide asked to be shorter than its passages
keeps them, and the passages are shared again only between such slides. A deck of three slides or fewer has no
opening or closing slide unless the brief asks for one: an answer of that many slides with a slide that carries
no passage is asked again once, and one that keeps an opening slide then stands. An opening or a closing slide
that shows no running text gives the passages the answer gave it to the slide next to it. In such a deck it
keeps them, as a slide of text, where they are more than a title line and the slide next to it would not hold
them beside its own (from 0.6.3): a cover given the document's title line stays a cover, and the other slides
carry the rest. A layout of a table given only the document's tables keeps them where it is, at an end of a
deck or in its middle, and passages shared again go round it (from 0.6.3). A change to a plan that names a
number of slides ("keep it to 2 slides") shares the passages again over all of them, a slide it keeps as it was
too, and asks again in the same way.

**The person's text comes first** (from 0.6.3). A slide that is still cut beside cards, figures, steps, a list,
a chart, a quote, a note, a table or a photo is its text alone, set dense, where free slides are on offer and
that keeps more of the text. The other parts give way. A table of the slide's passages is then part of its
running text. This is so in a plan, also one a layout the outline chose becomes (a chart beside a text, a
table), and in a change to a plan (`editPlan`) for the parts the plan had: a part a change asks for stays, in
that change and the ones after it, and the plan says what of the text it leaves out. A slide of the change is
the plan's slide of the same parts, in any order, of the same purpose, else in the same place, else in order;
one beyond the slides the plan had of those parts is new. A change to a deck (`planEdit`) keeps the parts a
slide has. A slide that holds its passages beside its other parts keeps them, and where sharing the passages
again makes the room, the parts stay. With free slides off, a slide of a plan on a layout that holds a few
characters of its text (cards, figures, a quote) moves to the layout of plain text that holds most of it, and
so, in a change to the plan, does a free slide of a text and cards that a plan made with free slides on, where
the layout keeps more of the text. Of the layouts of plain text that hold a text, one of a single text that
takes its paragraphs as they are comes first, so that code sets them there word for word. A layout of a table
counts the rows it shows in its table as kept, and moves only where the plain text keeps more.

**Writing.** A slide that holds its passages gets them word for word, set by code, where the document is in
the deck's language, nothing asks for it shorter, it has no table that goes into a table, and its pattern has
one running text. The writer then writes only its short slots, and its title where code does not set it. From
0.6.6 code also sets the title: the heading the slide's text opens under, word for word, where the title holds
it (on a free slide, in one line). The text then does not say that heading again, and a slide whose every slot
code sets asks the writer nothing. The other headings in the text are set in bold, each at the head of the
paragraph under it (`SlideSpec.headings`). A heading too long for the title is the writer's to title; a text
the writer writes keeps a heading as a paragraph of its own, and code marks it the same where it is word for
word. A text whose first paragraph says its title again (`Background:` under the title `Background`) loses
that paragraph. In a layout's list, an item set by code has the list's bullet in place of its own; a number
stays. Any other slide carrying passages is given
its text and told to keep the text's words and order, shortened only where a limit makes it: a translation, a
summary, a slide that is cut, a table on a slide with a table (which goes into the table, cut to the limits of
its cells), a layout of several running texts. On a slide with no table, a table of the document is part of
the running text, a paragraph for each row with its cells joined by " | ", and the slide's room is counted for
it. A slide in the deck's language that was not asked to be shorter, has no table that goes into a table and
has room for nearly all of its text is checked for the words it keeps: under 85% it is sent back once, and the
deck is told when it is still so: `the slide keeps 72% of the words of its text from the document (p3-p5),
though it has room for nearly all of it`. A slide that holds less than its passages is cut to fit, and the
deck is told: `the text from the document for this slide (p3-p5) is 4100 characters, and the slide holds about
3700 even with its text set small, at 7 pt: it is cut to fit. Ask for more slides to keep it whole`. What a
slide is said to hold is the room the writer is told, and what `planTextFit` counts. A slide set dense holds a
document's text as code would set it, its headings among its paragraphs, and code sets it there word for word
only where the slide holds it so (from 0.6.3); otherwise its writer writes it, told what the slide holds, in
at most twelve paragraphs, or, where the slide holds all of it in more, in the most that do, and joins the
rest. A text of more paragraphs than a layout's slots take has some of them joined, which loses no room. Such
a text goes on to a layout of plain text whose one running text takes its paragraphs, else, with free slides
on, to a free slide, where code sets its paragraphs as they are; where neither holds it, a layout that holds
its characters keeps it, and its writer joins some paragraphs. Each paragraph ends on a line of its own, and
on a right-to-left template five such paragraphs, each just over a line, ran 19 pt past the box of a text
layout that holds 417 characters. So a slot of paragraphs is told what any text of the text's paragraphs holds,
its template's spacing counted, or fewer paragraphs that hold more (see "What a writer is told a layout holds"
above), and its writer joins them. Code sets such a text there as it is where its own paragraphs fit, each
counted in its own whole lines; a slide whose text is written again (a change in other words, a change to its
title) is held to what its writer is told. Where the count does not speak for the slot (`countSpacing: false`,
or a template with no such spacing), a slot that takes four paragraphs or more takes one fewer of a text of
three or more, as from 0.6.3.
With free slides off the warning ends `; free slides are off, so its text is not set smaller`. The numbers a
chart shows may come from the document as well as from the brief.

**A summary or some parts.** "A short summary" or "short slides" in the brief makes the outline set `words` on
the slides that carry passages: the slide is then written shorter, from its own passages. `words` of four
fifths or more of the words of a slide's passages asks for no shorter text, and is dropped. "Only the part on
risks" puts the other passages in `left_out`.

**Limits.** At most 12,000 characters, about 1,800 Hebrew or 2,000 English words: a longer document is a
`PlanError` that asks for the part the deck is about. One document for a deck. A deck that is changed later
keeps its document and draws on it, but a new document needs a new plan. Without a `slideCount`, a deck of a
document has as few slides as its text takes, down to one. `imageDensity` defaults to `low`. Titles, headings
and table cells are cut to their limits. A heading of 60 characters or more with no `#` and no blank line under
it is read as text, and so is one right under a line of text with no stop at its end, or one over two short lines
with no stops: give a heading `#`, or a blank line under it. A short line over one other short line is read as
a heading over it, which a slide may take as its title. Passages shared again over a fixed number of slides go
by the size of their text, not by each slide's room: a slide with a table beside its text may be cut while
another has room.

### Before writing: `planTextFit`

`planTextFit(template, plan)` works out by code, without a model, what each slide that was asked for words or
carries passages will hold when the plan is written: the size its running text is set at on a free slide, the
words it holds, and whether it is set dense. A caller can tell the person before anything is written:

```ts
const plan = await planOutline(template, brief, { llm });
for (const fit of planTextFit(template, plan)) {
  if (fit.holds < fit.words) console.log(`slide ${fit.slide} holds ${fit.holds} of ${fit.words} words`);
  else if (fit.size !== undefined && fit.templateSize !== undefined && fit.size < 0.8 * fit.templateSize) console.log(`slide ${fit.slide}: text at ${fit.size} pt`);
}
```

Each `SlideTextFit` has `slide` (from 1), `words` (asked, or of its passages) and `holds`; on a free slide `size`
and `templateSize`, the size running text has there when nothing makes it smaller; on a dense slide `dense` and
`most`, the most words it holds at the least size that reads; and on a slide of a document `passages`, as
`"p3-p5"`. A plan with neither words nor a document gives `[]`. Pass it the options the plan was made with:
`patterns`, `free`, `icons`, `assets`, `layout`, `numberPatterns` and `countSpacing`. The chat tools say the same in `text_pt`,
`holds_words` and `note` (see [chat-tools.md](chat-tools.md)).

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
// plan = { version: 1, title, brief, slides: [{ pattern, purpose, items?, blocks?, arrange?, words?, passages? }], layouts?, documentLanguage? }

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
- "Make slide 2 about 200 words" gives that slide the room for the text (see "Asking for an amount of text"
  above): its layout, another layout of plain text, or a free slide written from what the slide held, which
  keeps the slide's chart, table or photo as it is, or a free slide of the text alone set dense. Unlike a plan,
  a free slide keeps its cards, list, chart or table, and a layout its chart or table, whatever they leave the
  text, and the deck says what it holds. A layout of cards or of a list that does not hold the text gives way
  to it, as before.
  A photo that gives way to such a text is left out, and the deck is told:
  `the photo of this slide is left out to make room for the 500 words asked for it`. The text stays on that
  slide: a change adds no slide by itself, only where the instruction asks for more slides.
- A dense slide whose text is written again (in other words, in another language, as a whole) keeps its length
  unless the instruction asks for another: its characters, counted as words of about 6 characters. A change
  that keeps the text, such as a better title, writes the title alone, and the text stays as it was.
- A part added beside a dense text, or beside a text from the document ("add three key points next to the
  text"), leaves that text less room. A text that still holds there stays as it is. One that does not is
  written again for the room it has, and cut, as below. The length it is held to is the one it has, so a text
  asked to be shorter before stays short, and a new text beside it takes none of that length: it is written
  as the request says. A photo that leaves a dense text too little room gives way to it,
  as for an amount of text asked: the text stays as it is, and the deck is told
  `the photo of this slide is left out to make room for its text of about 300 words`.
- A free slide is changed part by part: the parts the instruction does not touch stay exactly as they are. See
  "Free slides in an edit" below.
- In a deck made from a document (`deck.document`), the first call is shown the document with the passages of
  each slide and those on no slide, and the passages each slide carries after the change. "Put the part on the
  budget on its own slide" moves passages; "make slide 3 longer" gives it more from its neighbour or from those
  on no slide. A slide whose passages change is written again, its text set by code as the document has it
  where the slide holds it and the document is in the deck's language. Where the instruction asks for that
  slide's text in other words ("slide 2 in simpler words, and move the part on bikes onto it"), the first call
  says so, and the writer writes it. A slide whose passages stay keeps its text unless the instruction is about
  it.

`planEdit(template, deck, instruction, options)` does the same without pictures or rendering and returns
`{ deck, summary, rewritten, warnings }`.

A slide the change writes that would hold less than it was asked for, or less than its passages, is cut to
what it holds, and `warnings` says so after it was written. To be told before anything is written, pass
`stopBeforeCut`: `planEdit` then stops after its first call with a `TextCutError`, which names the slides. What
a slide holds is counted as it will be written, with the heading or caption it keeps, so the number the error
gives is the one the slide would have. A slide the change asks no length of, and whose parts and passages stay
(other words, a new title, another language), does not stop the change: its text is written to what it holds.
In a deck made from a document, the length of a slide's text in another script is a guess from the words of its
passages; only a deck put back into the language of its document, where that text is known, says after
writing that a slide holds less of it.

```ts
try {
  const { deck: next } = await planEdit(template, deck, 'Make slide 2 1000 words', { llm, stopBeforeCut: true });
} catch (error) {
  if (error instanceof TextCutError) console.log(error.message, error.fits); // nothing was written
  else throw error;
}
```

The message reads `Nothing was changed: slide 2 would hold about 446 of its 1000 words (text at 7 pt). Ask for
more slides or a shorter text, or accept that the rest is cut.`; `error.fits` has a `SlideTextFit` for each such
slide, numbered as after the change. Call again without `stopBeforeCut` to make the change and cut the text.

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
| `stopBeforeCut` | Stop before writing, with a `TextCutError`, when a slide the change writes would hold less than was asked for it or less than its passages. Off: such a slide is cut, and `warnings` says so. |

Changes made by hand in PowerPoint are not read back. The deck JSON is the source of truth.

## Errors

| Error | When | What to do |
| --- | --- | --- |
| `PlanError` | No pattern can be used with what was supplied; or the model's outline or edit plan is unusable; or the plan or the edit uses patterns that cannot be filled, or has a free slide the template cannot carry; or a plan given to `writeDeck` has a number of items its pattern does not take, or a free slide of more than 400 parts; or the brief's `document` is no text, or longer than 12,000 characters | Read the message: it names the patterns and reasons, or the slide and its number. Supply the missing icons, fields or images, correct the plan, or retry. |
| `TextCutError` (a `PlanError`) | `planEdit` with `stopBeforeCut`: a slide the change writes would hold less text than it carries. Nothing was written. | Tell the person what `error.message` says. Ask for more slides or a shorter text, or call again without `stopBeforeCut` to accept the cut. |
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
answers described in [providers.md](providers.md). Writing the slides keeps the model's default; the included
Gemini adapter asks for light reasoning only when it sends a request again after an answer cut off at its limit.

A deck made from a document or with an amount of text adds no call of its own: the outline is asked again at
most once, as for any answer it cannot use, and that now also covers one that places most of the document
nowhere or misses the number of slides asked. The prompts that plan carry the document, and the writer's answer
carries the text asked for: 500 Hebrew words are about 1,000 to 1,700 tokens.
A slide filled by code is not written by the model, so a document whose slides hold their text costs the writer
little more than its titles. A text written well short of what was asked, or a slide that keeps too few of its
document's words, takes at most one repair round, within `maxRepairs`. `planTextFit` asks no model.

To keep cost down: reuse `cache` across calls, set `imageDensity: 'low'` or `'none'`, and use `planOutline`
first so a person can correct the plan before content and pictures are produced.
