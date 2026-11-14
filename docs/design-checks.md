# Design checks and reviews

Validation (`validateDeck`) says whether a deck can be rendered and whether its text fits. This page is about
what comes after that: a deck that renders and fits can still look wrong. A body of two words in a box drawn
for a paragraph, one short card in a row of full ones, a title whose last word sits alone on a second line.

There are two tools:

- **`checkDesign`** finds such problems by measurement alone. No picture is made and no model is called.
- **A review** is a second look at each draft while a deck is being written. What it finds goes back to the
  writing model in the same repair round as the limits the draft broke. A review can be the design checks, a
  model that looks at pictures of the slides, or a function of your own. Nothing reviews a draft unless you ask.

Both need a template loaded with its fonts: see [text-fit.md](text-fit.md).

## `checkDesign(template, deck, options?)`

```ts
import { checkDesign } from 'slotdeck';

const problems = checkDesign(template, deck, { layout: 'adaptive' });
// [
//   { severity: 'warning', slide: 1, slot: 'title', kind: 'orphan',
//     message: 'the last of its 2 lines holds one word, "Share". Text of about 24 characters ends a line sooner; a few words more fill the line.' },
// ]
```

It is synchronous and takes a fraction of a second for a deck. It throws when the template has no fonts.

### What it finds

| `kind` | Problem | Message |
| --- | --- | --- |
| `overflow` | Text runs over its box. The same finding, with the same words, as the warning of `validateDeck` with fonts. | `text takes 7 lines at 18 pt; the box holds 5 (2 lines too many). Text of about 290 characters fits.` |
| `empty` | Text takes a third or less of the lines the template's own text takes in that box, so the box looks empty. | `text takes 1 line where the template's own text takes 5, so its box will look empty. Text of about 45 characters fills it as designed.` |
| `uneven` | In a row of cards, one card's text takes a third or less of the lines of its fullest neighbour. | `text takes 1 line where "card1Text" beside it takes 5, so the row will look uneven. Text of about 66 characters or more evens it out.` |
| `orphan` | A heading ends with one word on a line of its own. | `the last of its 2 lines holds one word, "Share". Text of about 24 characters ends a line sooner; a few words more fill the line.` |

Every message says what to do about it, with a length in characters. That is what the writing model can act
on, so a message can be sent to it as it is.

### The result

```ts
interface DesignIssue {
  severity: 'warning';
  slide: number;   // 1-based
  slot: string;    // the slot name as the deck uses it
  kind: 'overflow' | 'empty' | 'uneven' | 'orphan';
  message: string;
}
```

A `DesignIssue` is an `Issue` with a `kind`, so the list can be shown or stored together with the warnings of
`validateDeck` and `renderDeck`.

### Options

| Option | Default | Meaning |
| --- | --- | --- |
| `layout` | `'strict'` | As for `renderDeck`: with `'adaptive'`, optional slots may be left out. |
| `kinds` | all four | Which problems to look for, e.g. `['empty', 'orphan']`. |
| `emptyBelow` | `1/3` | A text is `empty` when it takes this share or less of the lines the template's own text takes; a card is `uneven` at this share or less of its fullest neighbour. |
| `headingSize` | `20` | The font size in points from which a single paragraph counts as a heading, for `orphan`. |

### How each check decides

**`overflow`.** The measured text needs more height than its box gives. If the manifest has `shrinkText`, a
text that fits once it is set smaller is not reported.

**`empty`.** The reference is the template itself: the lines its own sample text takes in the slot. A slot is
looked at only when that is three lines or more, so a title or a one-line caption is never "empty". A slot that
the manifest resized (`box`, `fontSize`: a derived pattern) has no sample to compare with; its box is
then taken to be meant half full. A `box` that only moves the shape, as in a mirrored pattern, keeps the
comparison. In a row that widens its items, the reference is the item as drawn: a short text in a wide card
takes fewer lines, so it is reported sooner, with the length that fills the wider card. A slot fed by a deck field is not checked, because that text is the app's.

**`uneven`.** The slots of one kind across a row are compared: every card's text with every other card's text,
every heading with every heading. Rows come from the pattern's `repeat` or its `sequences`. A row is looked at
only when its fullest item takes three lines or more. Items whose text overflows are left out of the
comparison.

**`orphan`.** A text of one paragraph, set at `headingSize` or larger, on two lines or more, whose last line
is a single word no wider than two fifths of the widest line. An address (`name@example.org`, a web address)
on its own line is not reported: it is meant to stand there.

### What it does not see

- Tables, charts and pictures. Only text slots are measured.
- Slots set in a typeface that was not supplied. They are passed over (see `missingFonts` in
  [text-fit.md](text-fit.md)).
- Anything that needs a picture to judge: colours, contrast, a photo that does not suit the text.
- Whether a line break looks good in the middle of a heading. Only the last line is looked at.

On this project's hand-written example deck the checks report two things, both real: the cover title ends with
"Share" on a second line, and the title of slide 4 with "trip". PowerPoint's own picture of the deck shows
both.

## Reviews: a second look while a deck is written

The planner writes a draft, validates it, and sends a draft that breaks the template's limits back to the
model (see "How planning works" in [generating.md](generating.md)). A review adds its own findings to that
round.

```ts
import { designReview, planDeck } from 'slotdeck';

const { deck, warnings } = await planDeck(template, brief, {
  llm,
  review: designReview(template), // empty boxes, uneven cards, orphan words
});
```

The `review` option is accepted wherever a deck is written: `planDeck`, `writeDeck`, `generateDeck`,
`planEdit`, `editDeck`, and `presentationTools`.

How it works:

1. After each draft, validation runs as always. If the draft has no errors, the review is called with it.
2. What validation and the review found on the slides just written goes back to the model, each as one line:
   `- s3.title: the last of its 2 lines holds one word, ...`.
3. This repeats up to `maxRepairs` times (2 by default). A review therefore runs up to three times for a deck.
4. Whatever is still found after the last round is returned in `warnings`, next to the validation warnings.

A review can make the planner use more model calls: a draft that was within its limits may now be sent back.
With `designReview` the review itself costs nothing; each repair it causes is one writing call.

### `designReview(template, options?)`

The code-only checks as a review. It reports `empty`, `uneven` and `orphan`; `overflow` is left to validation,
which reports it anyway. `options` are those of `checkDesign`.

### A review of your own

A review is a function from a draft deck to problems:

```ts
import type { DeckReview, Issue } from 'slotdeck';

const review: DeckReview = async (deck, { signal }) => {
  const problems: Issue[] = [];
  deck.slides.forEach((slide, index) => {
    if (typeof slide.slots.title === 'string' && slide.slots.title.endsWith(':')) {
      problems.push({ severity: 'warning', slide: index + 1, slot: 'title', message: 'a title must not end with a colon' });
    }
  });
  return problems;
};
```

Each problem is an `Issue`: `severity`, a 1-based `slide`, a `slot` when it is about one, and a `message`
written for the model. A problem without a `slide`, or on a slide the model was not asked to write, is not
sent back; it is returned in `warnings` all the same.

### `visionReview(options)`: a model that looks at the slides

```ts
import { planDeck, visionReview } from 'slotdeck';

const review = visionReview({
  template,
  llm: visionModel,   // an LlmProvider whose model can look at pictures
  preview,            // a PreviewProvider that gives PNG or JPEG
  icons,              // optional: the app's icons and supplied pictures, drawn as they are
  assets,
});
const { deck, warnings } = await planDeck(template, brief, { llm, review });
```

For each draft it:

1. renders the deck, with a plain grey rectangle for every picture that does not exist yet (photos are
   generated after planning);
2. asks the preview provider for a picture of each slide;
3. shows the model the pictures, six slides to a request, with the text of each slide's slots, and asks only
   for what can be seen and fixed by changing text: text cut off or lying over something, text too small to
   read, a box or slide left mostly empty, a heading that ends in a lone word;
4. returns what the model reports, each as `seen in the picture of the slide: ...`, with the slot when the
   model named one that the slide has.

| Option | Default | Meaning |
| --- | --- | --- |
| `template` | required | The template the deck is made from. |
| `llm` | required | A provider whose model accepts pictures (`JsonRequest.images`). It can be another model than the one that writes. |
| `preview` | required | Draws the draft. It must give bitmap pictures: PNG, JPEG, WebP or GIF. |
| `icons`, `assets` | none | Pictures the deck may use, drawn as they are instead of as rectangles. |
| `layout` | `'adaptive'` | As for `renderDeck`. |
| `slidesPerRequest` | `6` | How many slides are shown to the model in one request. |

Each request is capped at 300 output tokens plus 200 for each slide in it. A draft that cannot be rendered is
not looked at: its errors are validation's to report.

**Read this before turning it on.**

- **It is only as good as the pictures.** A model that is shown a preview which differs from PowerPoint will
  report problems that are not there. The Pagus adapter that comes with slotdeck gives SVG, which a model
  cannot read, and `visionReview` refuses it with an error. Use a provider whose pictures can be trusted, such
  as PowerPoint's own export on a machine that has it. See "How close the pictures are" in
  [preview.md](preview.md).
- **It has not been run against a live model, or against faithful pictures.** It was built and tested with a
  scripted model and made-up pictures: what is sent, how the answer is read, how it takes part in the repair
  round. How well a real model judges real slides with this prompt is not known. Try it on your own decks
  before relying on it.
- **It costs model calls with pictures:** one request for every six slides, for every draft, so up to three
  times a deck with the default `maxRepairs`.
- The code-only checks already cover overflow, empty boxes, uneven cards and orphan words, exactly and at no
  cost. A vision review is for what measurement cannot see.

### Reviews in the chat tools

```ts
const tools = presentationTools({ template, llm, store, review: designReview(template) });
```

The review then runs for every deck `create_presentation` writes and for the slides `edit_presentation`
rewrites. What it leaves unfixed is in the tool's `warnings`.
