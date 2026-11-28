# Style-only mode

From 0.10.0. A template usually comes with a manifest: someone names the layouts and their slots, with a model's
help and a person's eye (see [template-studio.md](template-studio.md)). Style-only mode needs none of that.
`styleManifest` reads a `.pptx` and makes a manifest by rules, in a fraction of a second and with no model. A
deck made on it takes the file's **style** and not its layouts:

- the file's own **cover** and **closing slide**, filled as they are drawn;
- between them, **slides the software lays out itself** ([free-slides.md](free-slides.md)): cards, figures, a
  process, a timeline, a comparison, charts, tables and running text, in the file's typefaces, sizes and colours,
  on the frame of one of its slides (its background, logo, header, footer and title).

Use it for a file nobody has prepared: a deck a person uploads and wants more decks "that look like this one".
Use a prepared manifest where the layouts themselves matter: a designed photo layout, a row of exactly four
figures, a slide that must look the same in every deck.

## In short

```ts
import { styleManifest, loadTemplate, generateDeck, gemini } from 'slotdeck';

const style = styleManifest(pptx, { name: 'Sales deck' });
if (style.reason) throw new Error(style.reason);          // this file needs a prepared manifest

const template = loadTemplate(pptx, style.manifest);
const { pptx: deck, warnings } = await generateDeck(
  template,
  { prompt: 'Our plan for 2027 ...', patterns: style.patterns, fields: { field1: 'Sales | North region' } },
  { llm: gemini({ apiKey }), layout: 'adaptive' },
);
```

Three things carry the mode:

- `style.manifest` is a manifest like any other. `loadTemplate`, `validateDeck`, `renderDeck`, `checkManifest` and
  the chat tools take it as it is.
- `style.patterns` is what a deck is planned from: `['cover', 'free', 'closing']`, without the cover or the closing
  slide where the file has none. Give it as `patterns` to a brief or to `presentationTools`. The manifest holds
  every slide of the file as a layout, because the text styles are read from them, but nobody named their slots:
  a planner that is not held to `style.patterns` would fill them by guesswork.
- `style.reason` is there where the file cannot carry slides the software lays out. Then there is no style-only
  mode for this file, and the sentence says why.

Write a deck by hand the same way: `{ pattern: 'cover', slots: { title, subtitle } }`, then `freeSlide({ title,
blocks })` for each slide, then `{ pattern: 'closing', slots: { title } }`.

## What comes back

```ts
interface StyleManifest {
  manifest: TemplateManifest;
  patterns: string[];              // what to plan a deck from: the cover, 'free', the closing slide
  reason?: string;                 // why the file cannot carry slides the software lays out; left out where it can
  cover?: string;                  // 'cover' where the first slide can be one
  closing?: string;                // 'closing' where the last slide can be one
  fields: Record<string, string>;  // the deck's own texts of the frame: field name -> the sample the file shows there
  notes: string[];                 // what was decided and why, in sentences
}
```

`fields` are the texts a deck gives once for all its slides (`DeckSpec.fields`, `Brief.fields`): the name of a unit
in the header, a date in the footer. Each is named `field1`, `field2` and so on, with the text the file shows
there, so that an app can ask a person "what goes where it says *Organization name | Department*?". A deck that
gives a field no value shows nothing there: the sample is taken off.

`notes` say what the rules decided: which slide the free slides stand on and why, what size running text has,
whether it is enlarged, how many slides were read. They are for a person who looks at the result, and for a log.

The same file always gives the same manifest. Store it beside the file, or make it again whenever the file is
loaded; a manifest made by a later version may differ, as the rules get better.

## What the rules decide

Nothing is asked of a model. Everything is measured on the file:

| What | How it is found |
| --- | --- |
| Direction and language | By the letters of the file's texts: right to left where most are Hebrew or Arabic. |
| The title of a slide | Its title placeholder; else its largest text in the upper half; else, on a slide with nothing up there (a cover), its largest text when that is clearly larger than running text. |
| The line under a title | A subtitle placeholder; else the nearest smaller single line under the title. It is optional: a deck that leaves it out has it removed. |
| A heading | A short line with a longer text right under it, in its column. |
| A big number | A few characters with a digit, set far larger than running text. |
| The deck's own texts | Footer, header and date placeholders with text, and short texts that repeat in the margins of many slides. They become `fields`. The page number stays the software's. |
| Photographs and tables | A picture that fills 3% of the slide or more and is no pictogram is a photo slot that keeps the file's own picture; a table is a table slot. |
| The cover and the closing slide | The first and the last slide, where each is a title slide: a title, the line under it, a text or two. A slide of content that comes first or last (a table, a list, headed columns) is neither. |
| The frame of free slides | Of the slides whose title sits where most titles sit, in the upper part of the slide: the one with the least drawn under its title by its slide layout, then one whose page keeps the file's glow in a corner, then the one with the fewest shapes. |
| Text that is too small to read | Running text under what 14 pt is on a slide 13.33 inches wide is enlarged on every slide (`manifest.textScale`), by no more than half. |
| Colours | By how much the slides use each colour, as for every template. A file whose slides draw nothing in a colour takes the first two accents of its theme (`manifest.colors`). |

## Files whose text is in placeholders

Most decks people make by hand have no text box of their own: a title and a bulleted body, both placeholders
that take their place, typeface, size and bullets from the slide layout and the master. Style-only mode reads
them as they are drawn:

- A placeholder is measured where its layout puts it and at the size it inherits, so the cover's title has the
  room of its real box.
- A free slide copies a placeholder as **a shape that states its own look**: no longer a placeholder, with the list
  style of every level, the insets and the anchor written into it, the nearest statement of each property kept.
  Put on any slide of the file, the copy reads as the placeholder did.
- Running text taken from a bulleted body has no bullet; a list keeps the file's bullets.
- A body set large (the 32 pt under a 44 pt title of an old default template) stands in at under half the title's
  size: that is what running text is beside such a title.
- Running text is the text most of the content is set in, the body of the content slides. The line under a
  cover's title is the last choice, and a placeholder its layout draws a line round (a boxed caption) is none.
- The blocks stand where the file's own body stands: beside a decoration its layout draws down one side, inside
  a card in the middle of the page, under a banner behind the title.

This is the engine's doing, not only this mode's: from 0.10.0 any template whose running text is all in
placeholders can carry free slides. A template with a text box to copy keeps the styles it always had; the
placeholders are copied only where there is nothing else, or where the manifest names one
(`manifest.free.styles`).

## When it cannot be used

`style.reason` is set, `style.patterns` has no `free`, and the manifest says `"free": false`:

| Reason | What it means |
| --- | --- |
| `no slide has a title in its upper part with room under it ...` | The file is all covers, dividers and full-page pictures (a photo album), or its titles stand under the content, at the foot of the slide. Slides the software lays out need one slide with a title at the top and content under it. |
| `the template's layouts have no running text that reads on this layout's background ...` | No text of the file can be copied as running text on the frame: every text is a title or a picture of text, or its colour does not read on the frame's page. |
| another sentence of the engine | As `validateDeck` would say it of a free slide: see "Limits" in [free-slides.md](free-slides.md). |

Such a file can still be a template with a prepared manifest.

## With the chat tools

```ts
const style = styleManifest(pptx, { name });
const tools = presentationTools({ template: loadTemplate(pptx, style.manifest), patterns: style.patterns, llm, store, layout: 'adaptive' });
```

Everything else is as with a prepared template: the deck fields a person gave, icons, pictures, colours, edits.
`layout: 'adaptive'` is the mode to use: a file that was never prepared has no budgets a person checked.

## What it does not do

- It does not say what a layout is for. The file's own content slides are in the manifest for their text
  styles, and a deck in this mode does not use them. To use them, prepare the manifest: `draftManifest` and
  `prepareManifest` start from the file, not from this manifest.
- It finds no rows of cards (`sequences`, `repeat`) and no icon slots: the software's own cards, figures and steps
  take their place.
- A cover's title has the room of the cover's own box. A long title is told so by `validateDeck`, as on any
  template.
- A file of fewer than six slides has too few to tell what repeats on every slide: only footers and page numbers
  are taken as its frame, and a logo drawn on each slide by hand may be missing from the free slides. A logo on
  the slide layout or the master is always there.
- Pictures that the file shows as text (a title saved as an image) cannot be read.

## Checking a result

- `checkManifest(loadTemplate(pptx, style.manifest))` has no error for a manifest this function made.
- `validateDeck(template, deck, { layout: 'adaptive' })` says of a deck what it says on any template.
- Look at one rendered deck before a file is offered to people: the rules measure, they do not see.
