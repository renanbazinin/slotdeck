# Template studio

The studio functions cover everything needed to turn a new `.pptx` into a usable template: drafting its
manifest with a model, checking it, seeing whether its budgets hold, and changing it by hand or by instruction.

```ts
import {
  draftManifest, checkManifest, sampleDeck, changeManifest, editManifest,
  studioView, describeManifest, describeSlot, inspectTemplate, ROLES, PptxPackage, loadTemplate, renderDeck,
} from 'slotdeck';
import type { ManifestChange, SlotPatch, ChangeResult, ManifestIssue, Role, StudioView } from 'slotdeck';
```

## The workflow

1. **Draft** with `draftManifest`. A model labels the shapes; the code measures everything.
2. **Check** with `checkManifest`. It should report nothing.
3. **Look** at a sample: `sampleDeck` then `renderDeck`, and open the file. Every slot is filled to its budget,
   so text that overflows here will overflow in real decks, and shapes that should not be there are easy to see.
4. **Change** what is wrong: `changeManifest` with operations, or `editManifest` with an instruction.
5. Repeat 2 to 4 until the sample looks right. Store the manifest as JSON next to the pptx.

A draft is a starting point for review, not a finished manifest. What a draft does not contain, and a person
adds: budgets tuned by eye, derived patterns (a text-only variant, a native chart), wider rows, picture areas,
chart colours, and alignment choices for decks that read the other way.

## `draftManifest(pptx, options)`

```ts
const { manifest, skipped, issues } = await draftManifest(pptxBytes, {
  llm,                                           // required: an LlmProvider
  name: 'marketing',                             // optional name for the manifest; default 'template'
  slideImage: (slide) => thumbnails.get(slide),  // optional: { mimeType, data } of a picture of slide N (1-based); may be async
  slides: [1, 2, 5],                             // optional: only these slides; default all
  concurrency: 4,                                // optional: slides sent to the model at once; default 4
  onProgress: ({ done, total, slide }) => {},    // optional
  signal,                                        // optional
});
// manifest: TemplateManifest
// skipped:  [{ slide, reason }]   slides that did not become patterns, and why
// issues:   ManifestIssue[]       what checkManifest finds in the draft
```

Model calls: one about the deck as a whole, then one per slide.

What the code works out without the model:

- Every shape's id, kind, position, size, font size and sample text.
- **Frame elements**: whatever sits in the same place on many slides (a logo, a rule), and page-number, date,
  footer and header placeholders. They are left alone and not shown to the model. (Spotting what repeats needs
  a template of at least six slides; with fewer, only the placeholders are treated as frame.)
- **Deck fields**: texts that appear on many slides become one value for the whole deck.
- The reading direction, from the letters of the sample text.
- A character budget for each text box, from its size and font size, and never less than its sample text.
- A table's columns and rows, and how many characters fit a cell before the table leaves the slide.
- Rows and grids of like items, from numbered slot names (`card1Heading`, `card2Heading`).

What the model decides, for each slide:

- Whether the slide is a reusable layout at all. A sheet of brand assets or a broken slide is skipped, with a
  reason.
- An id and a description for the pattern, and whether it is built around numbers.
- A role, a name and "optional or not" for each shape.

A label for a shape that is not on the slide is discarded, so the model cannot make the manifest point at
something that does not exist.

`slideImage` is optional. With it, the model sees the slide; without it, it judges from the measurements alone.
Any PNG or JPEG of the slide works; 480 pixels wide is enough.

To redo one slide after changing the pptx, call `draftManifest` with `slides: [n]` and add the pattern it
returns to your manifest.

### Roles

`ROLES` lists them. A role says what a shape is for and decides the kind of slot it becomes.

| Role | For | Becomes |
| --- | --- | --- |
| `title`, `subtitle` | The headings of the slide | Text slot |
| `text` | A paragraph, or a short label | Text slot |
| `list` | A bulleted or multi-line list, one paragraph per item | Text slot with `maxParagraphs` |
| `number` | A large figure | Text slot whose budget is its sample's length |
| `caption` | A few words under a figure, icon or picture | Text slot |
| `photo` | A photograph to replace | Image slot, `imageKind: "photo"` |
| `icon` | A small pictogram to replace | Image slot, `imageKind: "icon"` |
| `graphic` | A chart, diagram or screenshot picture to replace | Image slot, `imageKind: "graphic"` |
| `chart` | A native chart in place of a picture (not offered by the drafter; set it yourself) | Chart slot |
| `table` | A table | Table slot |
| `decoration` | Something drawn for the sample content, to remove from every slide made with the pattern | Slot with `empty: "drop"` |
| `fixed` | Part of the design that stays exactly as it is | No slot |

Which roles a shape can take depends on what it is: a text box can be any text role or a decoration; a picture
can be `photo`, `icon`, `graphic`, `chart` or `decoration`; a table only `table`; a line, a group or a drawn shape
without text only `decoration`.

## `checkManifest(template)`

```ts
const issues = checkManifest(loadTemplate(pptxBytes, manifest));
// [{ severity: 'error' | 'warning', pattern?: string, slot?: string, message: string }]
```

No model. It verifies that the manifest matches the pptx: slides and shapes exist, slots point at the right
kind of shape, names are unique, rows are consistent. The rules are listed in
[manifest-format.md](manifest-format.md). Run it when a template is loaded, so a template that someone changed
shows up there and not in the middle of a render.

For a template loaded with fonts it also tries every text budget out and warns about one that lets text run
over its box. See "The studio: checking and setting budgets" in [text-fit.md](text-fit.md).

## `sampleDeck(template, options?)`

```ts
const { deck, assets } = sampleDeck(template, { fill: 0.95, items: 'max' });
const { pptx } = renderDeck(template, deck, { assets, layout: 'adaptive' });
```

Builds a deck that uses every pattern once and fills every slot close to its limit, with the template's own
filler words and plain coloured rectangles for pictures.

| Option | Default | Meaning |
| --- | --- | --- |
| `fill` | `0.95` | How full to make each text slot, as a share of its budget. |
| `items` | `'drawn'` | For patterns with a row or grid: `'drawn'` uses the template's number, `'max'` the most allowed, `'min'` the fewest. |
| `fit` | `'budget'` | What "full" means for a text slot: `'budget'` is its character budget; `'box'` is what its box holds as measured with the template's fonts (see [text-fit.md](text-fit.md)). |

To look at the sample deck without opening PowerPoint, draw it with a preview provider:
`previewDeck(template, deck, preview, { assets, layout: 'adaptive' })`. See [preview.md](preview.md).

## `changeManifest(template, changes)`

```ts
const result = changeManifest(template, [
  { op: 'setSlot', pattern: 'cover', slot: 'title', maxChars: 40 },
  { op: 'setRepeat', pattern: 'icon-cards', max: 5 },
]);
// result.manifest  the new manifest; the template that was passed in is not touched
// result.applied   [{ index, change, message }]   what was done, each with a sentence saying what is now true
// result.rejected  [{ index, change, reason }]    what was refused, and why
// result.issues    checkManifest on the result
```

No model. Changes are applied in order. One that cannot be made, or that would leave the manifest pointing at
something the pptx does not have, is refused with a reason and skipped; the rest are still applied. To continue
from the result, load it: `loadTemplate(pptxBytes, result.manifest)` or `{ pkg: template.pkg, manifest: result.manifest }`.

### Operations

Every change is an object with `op`. `pattern` is a pattern id, `slot` a slot name, `shape` a shape id (string).

#### `setManifest`

```ts
{ op: 'setManifest', name?: string, language?: string, direction?: 'ltr' | 'rtl', alignOnDirectionChange?: 'flip' | 'keep', textScale?: number | null }
```

`textScale` enlarges the body text of every deck made from the template (`1.3` for 30% larger); `null` or `1`
returns to the designed size.

#### `addPattern`

```ts
{ op: 'addPattern', slide: number, id: string, description?: string }
```

Makes a slide a pattern with no content slots yet. Texts that other patterns already treat as deck fields are
filled in as deck-field slots. Follow it with `addSlot` operations, or use `draftManifest` with `slides: [n]`
instead to get a complete pattern.

#### `copyPattern`

```ts
{ op: 'copyPattern', pattern: string, id: string, description?: string }
```

A second pattern from the same slide, to turn into a variant with further operations.

#### `removePattern`

```ts
{ op: 'removePattern', pattern: string }
```

The last pattern cannot be removed.

#### `setPattern`

```ts
{ op: 'setPattern', pattern: string, id?: string, description?: string, figures?: boolean }
```

`id` renames the pattern. `figures: true` offers it only for briefs that contain numbers.

#### `addSlot`

```ts
{ op: 'addSlot', pattern: string, shape: string, role: Role, name?: string, optional?: boolean }
```

Makes a shape a slot in a role, with limits measured from the shape (budget from the box and font, table limits
from the table). On a shape that already is a slot, it replaces that slot, keeping its name unless `name` is
given: this is how a role is changed. Role `fixed` removes the slot. `name` defaults to the role.

#### `removeSlot`

```ts
{ op: 'removeSlot', pattern: string, slot: string }
```

The shape goes back to being part of the design.

#### `setSlot`

```ts
{ op: 'setSlot', pattern: string, slot: string, ...SlotPatch }
```

`SlotPatch` properties (all optional; `null` clears an optional setting):

| Property | Applies to | Meaning |
| --- | --- | --- |
| `name` | any | Renames the slot. |
| `description` | any | Guidance for the planning model. |
| `empty` | any | `error` (required), `remove` (optional), `keep`, `drop`. |
| `maxChars`, `maxParagraphs` | text | Budgets. |
| `deckField` | text | Ties the text to a deck-level field. |
| `alignOnDirectionChange` | text | `flip` or `keep` for this slot. |
| `styleRun` | text | `first` or `last`. |
| `fontSize` | text | A font size in points other than the template's. |
| `paragraphGap` | text | Space after each paragraph, in points. |
| `scales` | text | Whether a deck's text size applies. |
| `box` | text, image, chart | Another area for the shape: `{ x, y, w, h }` as fractions of the slide. |
| `imageKind`, `fit` | image | `photo`, `icon` or `graphic`; `cover` or `contain`. |
| `maxColumns`, `maxRows`, `maxCellChars` | table | Limits. |
| `headerText`, `bodyText` | table | `{ size?, bold?, color?, font? }`, merged into what is already set. |
| `maxCategories`, `maxSeries`, `colors`, `text`, `lineColor` | chart | Limits and styling. |

A property that does not apply to the slot's kind is refused (`"title" is a text slot; maxColumns does not apply to it`).

#### `setRepeat`

```ts
{ op: 'setRepeat', pattern: string, min?: number, max?: number, item?: string, span?: { left: number, right: number } | null }
```

Sets how many items a row or grid takes. When the pattern has no row yet, it is found from numbered slot names:
name the slots of each item alike with a number (`card1Heading`, `card1Text`, `card2Heading`, `card2Text`) and
they become a row.

#### `removeRepeat`

```ts
{ op: 'removeRepeat', pattern: string }
```

The items stop being a row whose number can change.

### What the code adds to each change

- **Budgets follow.** After a new `fontSize` or `box` on a text slot, or a new text size, column count or row
  count on a table, the character budget is worked out again, unless the same change states one. The message
  says so: `maxChars follows at 324`.
- **Rows are checked for room.** Asking for more items than the row draws widens the row as far as the slide
  allows: out to the slide's margin, or to the next shape beside the row. Asking for more than fit is refused
  with the most that do: `5 cards do not fit side by side: each is 21% of the slide wide and the row has 93%. The most is 4`.
- **Row items stay alike.** A `setSlot` on one item's slot (`card1Icon`) is made to that slot of every item.
  Slots of a row cannot be renamed, moved or removed one by one; remove the row first.
- **Nothing is left broken.** After each change the manifest is checked against the pptx. A change that
  introduces a new error is refused.

### Recipes

The pattern ids, slot names and shape ids below are those of the demo template in the package's repository;
with your template, use your own.

Remove a shape from every slide made with a pattern (a sticker, a stray page number, an ornament):

```ts
{ op: 'addSlot', pattern: 'quote', shape: '5', role: 'decoration' }
```

A native chart in place of a sample chart picture, as a second pattern:

```ts
[
  { op: 'copyPattern', pattern: 'chart-with-takeaway', id: 'chart-native', description: 'A native, editable chart with a short explanation.' },
  { op: 'addSlot', pattern: 'chart-native', shape: '6', role: 'chart', name: 'chart' },
  { op: 'setSlot', pattern: 'chart-native', slot: 'chart', colors: ['0E7C7B', 'F4A259'], text: { size: 12, color: '5C6B73' } },
]
```

A text-only slide from a slide with a photo:

```ts
[
  { op: 'copyPattern', pattern: 'text-and-photo', id: 'text-only' },
  { op: 'addSlot', pattern: 'text-only', shape: '2', role: 'decoration' },              // the photo
  { op: 'setSlot', pattern: 'text-only', slot: 'body', box: { x: 0.04, y: 0.49, w: 0.72, h: 0.39 }, fontSize: 20, maxParagraphs: 3 },
]
```

Smaller table text and fewer columns (the cell budget is measured for you):

```ts
{ op: 'setSlot', pattern: 'table', slot: 'table', maxColumns: 5, bodyText: { size: 9 }, headerText: { size: 10, bold: true } }
```

Larger body text in every deck, for a template whose own body text is small (render the sample deck afterwards
to see that everything still fits):

```ts
{ op: 'setManifest', textScale: 1.3 }
```

Mirror alignment for right-to-left decks, but keep the line beside the logo where it is:

```ts
[
  { op: 'setManifest', alignOnDirectionChange: 'flip' },
  { op: 'setSlot', pattern: 'cover', slot: 'organizationDepartment', alignOnDirectionChange: 'keep' },
]
```

## `editManifest(template, instruction, options)`

```ts
const result = await editManifest(template, 'Let the cards take up to six and make the subtitle optional everywhere', { llm, signal });
// the same fields as changeManifest's result, plus:
// result.note   what the model could not do or needs to know, in the person's language; often ''
```

The model reads the manifest next to the measured shapes of each pattern's slide and answers with operations
for `changeManifest`. It writes no manifest itself, so its answer goes through the same checks as a change made
by hand. One model call, and a second when some operations were refused, so the model can correct them.

- Instructions can be in any language.
- A person may refer to a slide by number or to a pattern by id.
- The model cannot set positions: `box` and `span` exist only for changes made by hand.
- An operation the model got wrong shows up in `rejected` with the reason; nothing of it is applied.

`describeManifest(template)` returns the text the model reads: every pattern with its slots and limits, and the
shapes of its slide that are not slots. An app can show it to its own chat model, or print it for review.

## `studioView(template)`

For a screen that edits a manifest by hand.

```ts
const view = studioView(template);
// view.aspectRatio            slide width divided by height
// view.slides[i].slide        1-based slide number
// view.slides[i].patterns     ids of the patterns made from this slide ([] when unused)
// view.slides[i].shapes[j]:
//   id        the id a change names the shape by
//   kind      'text' | 'picture' | 'shape' | 'line' | 'group' | 'table' | 'other'
//   box       { x, y, w, h } as fractions of the slide: lay it over a picture of the slide
//   groupId   the group the shape sits in, when it does
//   text      sample text, when it has some
//   fontSize  largest font size of the sample text, in points
//   frame     true when the shape repeats across slides (logo, page number): rarely content
//   roles     the roles addSlot accepts for this shape
//   slots     [{ pattern, name, kind, empty, summary }] slots that point at the shape
```

A typical screen draws a picture of the slide, lays a rectangle over each shape from `box`, colours the ones
that are slots, and on a click offers the shape's `roles` (sending `addSlot`) or the slot's settings (sending
`setSlot`). Nothing is stored in the view: call it again after every change.

`describeSlot(slot)` gives a slot in a few words, such as `text, up to 40 characters, required`.

## `inspectTemplate(pkg)`

```ts
const info = inspectTemplate(PptxPackage.load(pptxBytes));
// info.slideWidth, info.slideHeight   in EMU (914400 per inch)
// info.defaultFontSize                points
// info.slides[i] = { number, part, layout, shapes: ShapeInfo[] }
```

The raw list of shapes, without a manifest. `studioView` is this joined with a manifest.

## Storing a manifest

A manifest is plain JSON. Keep it next to the pptx it describes, and keep the two in step: if the pptx is
edited in PowerPoint, shape ids can change, and `checkManifest` will say which slots no longer match.
