# Template studio

The studio functions cover everything needed to turn a new `.pptx` into a usable template: drafting its
manifest with a model, checking it, seeing whether its budgets hold, and changing it by hand or by instruction.

```ts
import {
  draftManifest, prepareManifest, checkManifest, sampleDeck, changeManifest, editManifest, suggestChanges, SUGGESTION_KINDS,
  patternVariants, VARIANTS, studioView, describeManifest, describeSlot, inspectTemplate, ROLES, PptxPackage,
  loadTemplate, renderDeck,
} from 'slotdeck';
import type {
  ManifestChange, SlotPatch, ChangeResult, ManifestIssue, ManifestSuggestion, SuggestionKind, PatternVariant, PreparedManifest,
  Role, StudioView, VariantOption,
} from 'slotdeck';
```

## The workflow

1. **Draft** with `draftManifest`. A model labels the shapes; the code measures everything, and settles what
   measurement decides: how many items each row takes, and which small pictures are icons.
2. **Prepare** with `prepareManifest`. In one call and without a model it takes what the slides allow, adds
   the mirrored and text-only variants, solves the issues that have one answer, and returns the few sentences
   that are left for a person. A studio can stop here: draft, prepare, look at the sample, use it. The steps
   below are for looking closer.
3. **Check** with `checkManifest`. It should report nothing.
4. **Look** at a sample: `sampleDeck` then `renderDeck`, and open the file. Every slot is filled to its budget,
   so text that overflows here will overflow in real decks, and shapes that should not be there are easy to see.
5. **Take suggestions** one by one with `suggestChanges`: what the slides allow that the manifest does not use
   yet, each ready to apply. No model. `prepareManifest` has taken most of them already.
6. **Change** what is wrong: `changeManifest` with operations, or `editManifest` with an instruction.
7. Repeat 3 to 6 until the sample looks right. Store the manifest as JSON next to the pptx.

A draft can be used as it is: a deck may ask it for two cards or for six, and gets them. It is still worth a
review. What a draft does not contain, and a person adds: budgets tuned by eye, a native chart, picture areas,
chart colours, and alignment choices for decks that read the other way. Variants and rows that widen are a
matter of taste; `suggestChanges` lists them and a click adds them.

**Do not draft again to repair a manifest.** A model names and describes the same slide differently each time,
so a second draft replaces every pattern id and description, drops every correction made since, and leaves
stored decks pointing at ids that no longer exist. To improve a manifest that is in use, ask `suggestChanges`
and apply what it returns: ids and slot names stay as they are. See "Repairing a manifest that is in use" below.

## `draftManifest(pptx, options)`

```ts
const { manifest, skipped, issues } = await draftManifest(pptxBytes, {
  llm,                                           // required: an LlmProvider
  name: 'marketing',                             // optional name for the manifest; default 'template'
  slideImage: (slide) => thumbnails.get(slide),  // optional: { mimeType, data } of a picture of slide N (1-based); may be async
  slides: [1, 2, 5],                             // optional: only these slides; default all
  variants: ['mirror', 'text-only'],             // optional: also add these variants of every pattern that allows them; default none
  rows: 'fit',                                   // optional: 'fit' (default) lets a row take as many items as the slide has room for; 'drawn' only what the template draws
  widen: 2,                                      // optional: let the items of rows and grids grow up to this factor when a line holds fewer than drawn; default off
  concurrency: 4,                                // optional: slides sent to the model at once; default 4
  onProgress: ({ done, total, slide }) => {},    // optional
  signal,                                        // optional
});
// manifest:    TemplateManifest
// skipped:     [{ slide, reason }]   slides that did not become patterns, and why
// issues:      ManifestIssue[]       what checkManifest finds in the draft
// notes:       string[]              what the code settled by measurement, in sentences
// suggestions: ManifestSuggestion[]  what the slides allow beyond the draft; see suggestChanges
```

Model calls: one about the deck as a whole, then one per slide. Each call carries a ceiling on the size of its
answer, as `maxOutputTokens` on the request: for a slide, 400 tokens plus 60 for each shape listed; for the
deck, 300 tokens plus 40 and the size in bytes of each repeated text, which the answer has to repeat. The
ceiling ends a model that runs away inside a text field. A provider should pass it on as its output limit; the
Gemini adapter does.

Before any model call, a draft is refused for a `rows` that is neither `'fit'` nor `'drawn'`, a `widen` that is
not a number above 1 and at most 4, and a `variants` entry other than `'mirror'` and `'text-only'`.

What the code works out without the model:

- Every shape's id, kind, position, size, font size and sample text.
- **Frame elements**: whatever sits in the same place on many slides (a logo, a rule), and page-number, date,
  footer and header placeholders. They are left alone and not shown to the model. (Spotting what repeats needs
  a template of at least six slides; with fewer, only the placeholders are treated as frame.)
- **Deck fields**: texts that appear on many slides become one value for the whole deck.
- The reading direction, from the letters of the sample text.
- A character budget for each text box, from its size and font size, and never less than its sample text.
- A table's columns and rows, and how many characters fit a cell before the table leaves the slide.
- Rows and grids of like items, from numbered slot names (`card1Heading`, `card2Heading`), and how many items
  each row has room for.
- Which pictures are pictograms: no side longer than 1.1 inch, not a JPEG, not cropped into a frame. The model
  reads `small: a pictogram` on such a shape's line.

What the model decides, for each slide:

- Whether the slide is a reusable layout at all. A sheet of brand assets or a broken slide is skipped, with a
  reason.
- An id and a description for the pattern, and whether it is built around numbers. It is told to name and
  describe the layout by its structure ("a row of cards, each with an icon, a heading and a short text"), never
  by the subject of the filler ("profiles of team members") and without a number of items: the planner chooses
  a pattern by its description, and a layout described as team profiles is passed over for anything else.
- A role, a name and "optional or not" for each shape.

A label for a shape that is not on the slide is discarded, so the model cannot make the manifest point at
something that does not exist.

`slideImage` is optional. With it, the model sees the slide; without it, it judges from the measurements alone.
Any PNG or JPEG of the slide works; 480 pixels wide is enough.

To redo one slide after changing the pptx, call `draftManifest` with `slides: [n]` and add the pattern it
returns to your manifest.

### What a draft settles by itself

Two things are facts of the slide, so the code decides them and says so in `notes`.

**How many items a row takes.** A row drawn with four cards takes four by default in the manifest format, and a
deck that wants five then has nowhere to go: the planner picks a table or a grid. With `rows: 'fit'`, the
default, a drafted row takes as many items as fit side by side at their drawn width, while the space between
items stays at a third of the drawn space or more. The row may spread for it, and its `span` is then set: out
to the slide's margin, to the next shape beside it, or to the edge of a shape it is drawn on, such as a tinted
panel. `rows: 'drawn'` keeps every row at what the template draws; the change then comes back as a `row-max`
suggestion.

Some rows take what they draw, whatever the room:

- A grid, because one more item would need another line.
- A row whose items touch, or differ in width by more than its spacing can absorb. Further items are copies of
  the last, set out by the step from the first item to the second; where that step leaves less than a third of
  the drawn space after the widest item, the promise cannot be kept.
- A row with a line, an arrow or a highlight among its items. It stays where it is, and further items would be
  drawn across it. So would fewer: such a row is drafted to hold exactly the number it draws, and is offered no
  widening. A rule that runs the whole length of the row, or a panel behind all of it, is not among the items:
  it is something the row is drawn on, and the row may take more inside it. A shape the pattern takes off every
  slide (a slot with `empty: 'drop'`) is in nobody's way, and neither is anything inside a group it takes off.

```ts
// notes
// '"icon-columns": the row of columns takes 2 to 6. The row may now spread from 4% to 96% of the slide\'s width.'
```

A person may still ask for more with `setRepeat`, up to the number at which items would touch; the refusal
says that number.

**Which pictures are icons.** A model that does not see the slide was observed to label the small pictograms of
a row of cards as photographs. Such a layout then asks for one photograph per card, and is not offered at all
when pictures are turned off. A picture labelled `photo` that measures as a pictogram becomes an icon slot. Its
name stays as the model gave it, so the row its names describe is the row that is found.

```ts
// notes
// 'Slide 15: "card1Photo", "card2Photo", "card3Photo", "card4Photo" were labelled photo, but their pictures are no more than 1.1 inch on a side: icon slots.'
```

For a row this is decided for all its items together, because every item of a row is filled alike and items
beyond the drawn ones are copies of the last. When each picture of the row is a pictogram, each one labelled
`photo` becomes an icon. When one is not (a portrait cropped into a frame among pictograms) and the model's
labels disagree, some `icon` and some `photo`, they all become photo slots, and a note says so. Labels that
agree are left as the model gave them.

An icon slot is filled from the icon set the app gives the planner (`icons` on `planDeck`, `generateDeck` and
`presentationTools`). An app that gives none should know that a pattern with a required icon slot is then not
offered at all (`needs icons, and no icon set was supplied`). Such an app can keep the template's own
pictogram on every card with `{ op: 'setSlot', pattern, slot, empty: 'keep' }`; one change covers a row.

`widen` is off unless asked for: whether two cards should grow to fill a row drawn for four is a matter of
taste. With `widen: 2` it is set on every drafted row or grid where it can change a slide; a grid whose fewest
items already fill a line gets none. Without the option, it comes back as a suggestion.

### Variants in a draft

With `variants`, the draft also gets the mirror image and the text-only version of every drafted pattern whose
slide allows one (see `derivePattern` below). They are worked out by code after the model has labelled the
slides, so they cost no model call. A pattern that allows no variant simply gets none. A variant's row is
settled for the room the variant has, which without the picture can be more than its pattern's.

This is off by default, for three reasons. Every variant is one more layout to review in the sample deck.
Whether a mirrored slide suits a design is for a person to say. And a variant is only as good as the labels it
is derived from, which a draft has not had reviewed yet. So the draft leaves them out, and they are added
after it: by `prepareManifest` in a studio that prepares a template in one step, or one by one with
`derivePattern` in a studio where a person reviews each layout first. Use `variants` only when the draft itself
should show everything a template can do. The `without` variant is never added by a draft, because it needs a
slot to be named.

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
over its box. A row with `widen` gives its texts a larger budget when a slide has fewer items; that budget is
tried in the wider box, for each number of items that widens the row, and a warning then says which
(`on a slide with 2 of its 4 items`). See "The studio: checking and setting budgets" in
[text-fit.md](text-fit.md).

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
| `items` | `'drawn'` | For patterns with a row or grid: `'drawn'` uses the template's number, or the nearest the row takes when it does not take that number; `'max'` the most allowed; `'min'` the fewest. |
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

A second pattern from the same slide, to turn into a variant by hand with further operations. For the mirror
image, a text-only version or a version without an optional line, use `derivePattern`, which works out the
positions.

#### `derivePattern`

```ts
{ op: 'derivePattern', pattern: string, variant: 'mirror' | 'text-only' | 'without', slot?: string, id?: string, description?: string }
```

A second pattern from the same slide, laid out by the code. You name the variant; every position and budget is
worked out from the shapes of the slide. `VARIANTS` lists the variant names.

| Variant | What it makes | Default id |
| --- | --- | --- |
| `mirror` | The picture and what sits beside it change sides. | `<pattern>-mirrored` |
| `text-only` | The picture is left out, and the text beside it widens over the room it leaves. The items of a row are never widened one by one; where the only text beside the picture is a row, the variant is refused. | `<pattern>-text-only` |
| `without` | The optional text named in `slot` is left out, and what sits under it moves up. | `<pattern>-no-<slot>` |

- `slot` is needed for `without` only, and must name an optional text slot (`"empty": "remove"`).
- `id` and `description` replace the defaults. The default description is the source's with a sentence added,
  such as `Mirrored: here the picture is on the left and the text on the right. Alternate it with "photo-stat" for variety.`
  Write your own when the source's description names a side ("a photo on the right"), so that the planner
  reads one clear statement.
- The new pattern is placed right after its source. The source is not changed.
- A variant can be derived from a variant: the mirror image, then that without its subtitle.
- A variant the slide does not allow is refused with the reason, such as
  `"cover" has no photo, graphic or chart of its own, so there is nothing to change sides with`.

How each variant is worked out, what it leaves alone and when it is refused is described under "Variants the
code works out" in [layout-freedom.md](layout-freedom.md). `patternVariants` (below) says in advance which
variants a pattern allows.

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
{ op: 'setRepeat', pattern: string, min?: number, max?: number, item?: string, span?: { left: number, right: number } | null, widen?: number | null }
```

Sets how many items a row or grid takes. When the pattern has no row yet, it is found from numbered slot names:
name the slots of each item alike with a number (`card1Heading`, `card1Text`, `card2Heading`, `card2Text`) and
they become a row.

A row may take more items than it draws, as long as they fit side by side at their drawn width; a refusal says
the most that do. `suggestChanges` proposes a number that keeps a third of the drawn spacing, which is fewer
than the most that fit. A grid takes at most what it draws. More than drawn is also refused for a row whose
items are drawn overlapping, or differ in width so much that the step from the first to the second is shorter
than the widest: further items are copies of the last, set out by that step, and would lie on top of each
other. The fit is counted when a change gives `max` or a `span`, or finds the row; a change to something else,
such as `min` or `widen`, is not held up by it.

A `span` is refused when the drawn row does not fit in it, and for a row whose items sit inside a group: the
renderer lays out between the limits of a span only a row whose items are on the slide. `span: null` takes a
span away, and with it the room it gave: a row that then takes more than fit has its most lowered to what
fits, and the message says so.

A row is laid out from where the slide draws its items. Where a slot of the items carries a `box`, and so is
shown in another place, the items are not made a row, and a row that exists takes no more items and no wider
ones; `setSlot` with `box: null` takes the box away.

`widen` lets the items grow wider when a line holds fewer of them than the template draws, so that two or
three cards fill a row drawn for four. It is the most an item may grow, as a factor on its drawn width,
between 1 and 4; `null` or `1` turns it off. See "Fewer items that fill the row" in
[layout-freedom.md](layout-freedom.md). Render `sampleDeck(template, { items: 'min' })` to look at the result.

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
- **Variants are laid out.** `derivePattern` moves and resizes the shapes of a variant by itself, measures
  the budgets of every text that changed size, and writes a description the planner can choose the variant by.
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

The mirror image of a slide with a photo, for variety in a deck with several such slides:

```ts
{ op: 'derivePattern', pattern: 'photo-stat', variant: 'mirror', description: 'One key figure with a title and a short explanation on the left, and a photo on the right.' }
```

A text-only slide from a slide with a photo, then larger text with room for three paragraphs (the budget
follows the new size):

```ts
[
  { op: 'derivePattern', pattern: 'text-and-photo', variant: 'text-only', id: 'text-only' },
  { op: 'setSlot', pattern: 'text-only', slot: 'body', fontSize: 20, maxParagraphs: 3, paragraphGap: 12 },
]
```

The same slide for decks that have no subtitle, with the body moved up into its place:

```ts
{ op: 'derivePattern', pattern: 'text-and-photo', variant: 'without', slot: 'subtitle' }
```

Two or three cards that fill a row drawn for four, each holding more text:

```ts
{ op: 'setRepeat', pattern: 'icon-cards', widen: 2.2 }
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
- The model cannot set positions: `box` and `span` exist only for changes made by hand. It can ask for a
  variant by name ("add a mirrored version of the photo slides", "a version of slide 2 without the picture"),
  which becomes a `derivePattern` operation: the model names the variant, and the code works out where
  everything goes.
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

## `prepareManifest(template, options)`

```ts
const draft = await draftManifest(pptxBytes, { llm });
const prepared = prepareManifest(loadTemplate(pptxBytes, draft.manifest), { icons: true });
// prepared.manifest  the manifest to store
// prepared.sound     true: no error, and the sample deck renders at the fewest and at the most items
// prepared.left      ['One text is the same on every slide and shows what the template shows until you set your own: department ("Department name").',
//                     "Text was not measured with the template's fonts, so budgets are estimates. Look at the sample deck for text that runs over its box.",
//                     'Read the description of each layout once. The planner chooses a layout by it, so it should say what the layout holds and not what the sample is about.']
// prepared.done      what was done, a sentence for each change
// prepared.added     ids the manifest has now and did not have: the variants that were added
// prepared.removed   ids the manifest had and no longer has: patterns that could not be used
// prepared.kept      [{ pattern, slot }]: icon slots set to keep the template's pictogram, with icons: false
// prepared.fields    { department: 'Department name' }
```

Prepares a manifest for use in one step, for a studio that puts one control where a list of suggestions and a
list of issues would be. No model, and the template passed in is not touched. It renames no pattern and no
slot. A second run with the same options changes nothing and has an empty `done`.

What it does, in this order:

1. Repairs a row whose items hold a slot that is not the row's, as an earlier draft made one where the
   pictogram of each card was labelled a decoration. On a slide with fewer items that slot's shape is gone
   with its item, and the slide cannot be rendered. Where every item has such a slot under its own number and
   no deck fills it (`empty: 'keep'` or `'drop'`), the slot becomes part of the row. Otherwise the row takes
   no fewer items than the one the slot is in. Every deck that rendered still does.
2. Takes out a pattern for which `checkManifest` reports an error, unless `remove` is `false`. Two patterns
   with one id are left for a person: nothing says which is meant. A grid said to take more than it draws
   takes what it draws.
3. Takes the `icon` suggestions (see `suggestChanges` below). Then, with `icons: false`, sets every required
   icon slot to keep the template's own pictogram, and lists those slots in `kept`. A layout with a required
   icon slot and no icon set is not offered to the planner at all; this keeps it offered. An icon slot a deck
   may leave out stays as it is.
4. Takes `row` and `row-max`, lets rows and grids widen by the factor in `widen`, adds the variants in
   `variants` wherever a slide allows them, and settles the rows of those variants: a layout without its
   picture has room its pattern did not have.
5. With fonts, sets the measured budget where a budget lets text run over its box. The items of a row get one
   budget, the smallest any of them measures. Without fonts no such issue is reported, and none is solved.
6. Renders the sample deck at the fewest and at the most items. A pattern whose own sample cannot be rendered
   is taken out, unless `remove` is `false`.
7. Works out a value for each deck field: the one given in `fields`, or the text the template itself shows in
   that place. The manifest does not hold deck-field values; pass `prepared.fields` to the planner as `fields`
   until a person sets their own.

| Option | Default | Meaning |
| --- | --- | --- |
| `icons` | required | Whether the app has an icon set for this template, which it gives the planner as `icons`. |
| `variants` | `['mirror', 'text-only']` | The variants to add. `[]` adds none. `'without'` is not in the default: there is one for every optional text, and an optional text can be left empty anyway. |
| `widen` | `2` | The factor rows may widen by, above 1 and at most 4. `false` leaves them at their drawn width. |
| `remove` | `true` | Whether a pattern that cannot be used may be taken out. `false` for a manifest that stored decks were made with. |
| `fields` | none | The values the app already has for deck fields. A field that has one is not asked about in `left`. |

`left` is the short list a studio shows first. It holds what code cannot decide, in this order: patterns that
were taken out, errors that could not be solved, a sample that does not render, warnings with no single
answer, the deck fields that have no value of their own yet, whether text was measured, and a reminder to read
the descriptions. `sound` says whether the result is fit to use. Throws when `icons` is not `true` or `false`,
and for a `widen` or a variant that does not exist.

### After a draft, and for a manifest in use

After a draft nothing depends on the manifest yet, and the defaults are right. A pattern that cannot be used
is taken out; `done` and the first sentence of `left` say which and why. The last pattern of a manifest is
never taken out.

A manifest that stored decks were made with must keep every pattern: a deck whose pattern is gone can no
longer be changed. Pass `remove: false`. A pattern that cannot be used then stays, `sound` is `false`, and
`left` says what is wrong with it, so that a person decides. With `variants: []` and `widen: false` as well,
no layout is added and no look changes: it settles the facts only.

`prepareManifest` does not know what a person undid by hand. A variant that was removed is added again, and a
row that was left at its drawn width widens again. For a later run on a manifest a person has worked on, pass
`variants: []` and `widen: false`, or the variants and the factor they chose.

`kept` is for the day the app gets an icon set: those are the slots to give back to it, each with
`{ op: 'setSlot', pattern, slot, empty: 'error' }`. A slot a person set to keep by hand is not in the list.

After it, with its default options, `suggestChanges` lists only the `without` variants, which are left to a
person. With `widen: false` it lists `row-widen` as well, and with fewer `variants` the ones that were left
out.

## `suggestChanges(template, options?)`

```ts
const suggestions = suggestChanges(template);
// [
//   { kind: 'icon', pattern: 'team-cards', slide: 15,
//     message: '"card1Photo", "card2Photo", "card3Photo", "card4Photo" are photo slots, but their pictures are no more than 6% of the slide wide: pictograms. As icon slots they are filled from the icon set given to the planner, and the layout no longer waits for photographs. Without an icon set, a required icon slot keeps the layout from being offered.',
//     changes: [{ op: 'setSlot', pattern: 'team-cards', slot: 'card1Photo', imageKind: 'icon' }] },
//   { kind: 'row-max', pattern: 'team-cards', slide: 15,
//     message: 'The row of cards in "team-cards" takes up to 4, and the slide has room for 5 side by side at their drawn width, with the space between them no less than a third of what the template draws. For that the row spreads from 4% to 68% of the slide\'s width.',
//     changes: [{ op: 'setRepeat', pattern: 'team-cards', max: 5, span: { left: 0.04, right: 0.68 } }] },
//   { kind: 'row-widen', pattern: 'team-cards', slide: 15, message: '...', changes: [{ op: 'setRepeat', pattern: 'team-cards', widen: 2 }] },
//   { kind: 'variant', variant: 'text-only', pattern: 'text-and-photo', slide: 3, message: '...', changes: [{ op: 'derivePattern', pattern: 'text-and-photo', variant: 'text-only' }] },
// ]

const result = changeManifest(template, suggestions.filter((s) => s.kind !== 'variant').flatMap((s) => s.changes));
```

No model, and nothing is changed. It reads the manifest against the pptx and lists what the slides allow that
the manifest does not use yet. It works on any manifest: one a draft made, one written by hand, one that has
been in use for months. Every suggestion has been tried against the pptx before it is listed, so its `changes`
apply as they stand.

| `kind` | What was found | The change |
| --- | --- | --- |
| `icon` | Photo slots whose pictures measure as pictograms (no side over 1.1 inch, not a JPEG, not cropped into a frame). For a row, and for items a `row` suggestion would make a row, only when every item's picture does. Slots that nothing fills (`empty: 'keep'` or `'drop'`) are left alone. | `setSlot` with `imageKind: 'icon'`. One change covers a row; where the row is not made yet, one for each slot. |
| `row` | Slots that name several items alike (`card1Text`, `card2Text`) in a pattern with no row, so every slide draws them all. | `setRepeat`, which finds the row. |
| `row-max` | A row with room for more items than it takes, at a third of the drawn spacing or more. Never a grid, a row with an arrow, a line or a highlight among its items, or a row whose items touch or differ in width by more than the spacing can absorb. | `setRepeat` with `max`, and `span` when the row has to spread for it. |
| `row-widen` | A row or grid where some number of items it takes leaves a line short of what the template draws, and the items keep their drawn width. Never a row with an arrow, a line or a highlight among its items. | `setRepeat` with `widen`. |
| `variant` | A mirrored, text-only or shortened layout the slide allows and the manifest does not have. `variant` says which, and `slot` the text a `without` leaves out. A variant is known by the id a derived pattern gets, or by what it does to the slide (which shapes it takes away, what fills its pictures, on which side they sit), so one that was renamed and tuned is not offered again. | `derivePattern`. |

`SUGGESTION_KINDS` lists the kinds. `row` and `row-max` are facts of the slide and safe to take. `icon` is a
fact too, with one condition: icon slots are filled from the icon set the app gives the planner, and without
one a pattern with a required icon slot is not offered. `row-widen` and `variant` are choices about how a deck
should look.

| Option | Default | Meaning |
| --- | --- | --- |
| `kinds` | all | Only these kinds. A kind that does not exist throws. |
| `widen` | `2` | The factor suggested by `row-widen`: a number above 1 and at most 4. Anything else throws when `row-widen` is asked for. |

Ask again after applying: one change can open another. A row found by `row` is then asked how many items it
takes, and a variant derived from a pattern copies the row as it is at that moment, so take `row-max` before
`variant`. Each suggestion is tried by itself, so two that do not go together can be listed side by side, such
as `icon` for a picture and the `text-only` variant that leaves it out; when both are sent, the later one may
come back in `rejected` with its reason. An unknown kind in `kinds` throws. A variant is offered once, for the first pattern of a slide; a pattern that is itself a variant gets
none.

A studio screen can show each suggestion as a line of text with a button that sends its `changes`.

### Repairing a manifest that is in use

A manifest drafted with an earlier version keeps every row at what the template draws, and may have photo
slots where the slide has pictograms. `prepareManifest(template, { icons, remove: false })` repairs it in one
call, and with `variants: [], widen: false` as well it adds no layouts and changes no look. The same repair by
hand takes three lines, with no model and with every pattern id and slot name unchanged, so stored decks keep
working:

```ts
const template = loadTemplate(pptxBytes, savedManifest);
const facts = suggestChanges(template, { kinds: ['icon', 'row', 'row-max'] });
const { manifest, applied, rejected } = changeManifest(template, facts.flatMap((s) => s.changes));
// store `manifest`; show `applied[i].message` to whoever reviews it
```

In an app that gives the planner no icon set, follow the repair with
`{ op: 'setSlot', pattern, slot, empty: 'keep' }` for each such row. It keeps the template's own pictogram on
every card, and the pattern is offered whatever the setting for pictures. A pattern with a required icon slot
and no icon set is not offered at all. Leaving `'icon'` out of `kinds` instead keeps the slots as photographs:
the layout then asks for one photograph for each card, and is not offered when pictures are off.

What the code cannot repair is a description that names the subject of the filler ("Four team member profiles
with photos"). The planner reads it and passes the layout over for other subjects. Change it with
`{ op: 'setPattern', pattern, description }`, or ask `editManifest` in a sentence.

## `patternVariants(template, pattern)`

```ts
const options = patternVariants(template, 'text-and-photo');
// [
//   { variant: 'mirror',    available: true, note: 'the mirror image of "text-and-photo", with the picture on the left and what sat beside it on the right. ...' },
//   { variant: 'text-only', available: true, note: '"text-and-photo" without its picture "photo". Widened over the freed area, with budgets to match: ...' },
//   { variant: 'without', slot: 'subtitle', available: true, note: '"text-and-photo" without "subtitle": body moves up and has more room (431 characters).' },
// ]
```

No model, and nothing is changed. It lists the variants `derivePattern` could make from a pattern: the mirror
image, the text-only version, and one `without` for each optional text slot. An entry with `available: false`
carries the reason in `note`; an available one says there what the variant would do. A studio screen can show
the available ones as buttons that send `{ op: 'derivePattern', pattern, variant, slot }`. It throws when the
template has no pattern of that id.

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
