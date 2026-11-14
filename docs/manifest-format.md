# Manifest format

A manifest is a JSON object of type `TemplateManifest`. It describes one template: which of its slides are
reusable patterns and which shapes on them are slots.

```ts
import type { TemplateManifest, PatternDef, SlotDef, TextSlotDef, ImageSlotDef, TableSlotDef, ChartSlotDef, RepeatDef } from 'slotdeck';
```

A manifest is written once per template. You can have a model draft it (`draftManifest`), change it with
operations (`changeManifest`) or by instruction (`editManifest`), and verify it (`checkManifest`): see
[template-studio.md](template-studio.md). This page is the reference for what the JSON means.

## Example

```json
{
  "version": 1,
  "name": "demo",
  "language": "en",
  "direction": "ltr",
  "alignOnDirectionChange": "flip",
  "patterns": [
    {
      "id": "cover",
      "sourceSlide": 1,
      "description": "Opening slide with the presentation's title and a subtitle.",
      "slots": [
        { "name": "title", "kind": "text", "shapeId": "5", "size": 50, "maxChars": 36 },
        { "name": "subtitle", "kind": "text", "shapeId": "6", "size": 20, "maxChars": 86, "empty": "remove" },
        { "name": "organizationDepartment", "kind": "text", "shapeId": "7", "size": 13, "maxChars": 61, "deckField": "organizationDepartment", "alignOnDirectionChange": "keep" }
      ]
    },
    {
      "id": "icon-cards",
      "sourceSlide": 5,
      "description": "A row of cards, each with an icon, a heading and a short text.",
      "slots": [
        { "name": "title", "kind": "text", "shapeId": "5", "size": 34, "maxChars": 41 },
        { "name": "card1Icon", "kind": "image", "imageKind": "icon", "shapeId": "9" },
        { "name": "card1Heading", "kind": "text", "shapeId": "10", "size": 19, "maxChars": 15 },
        { "name": "card1Text", "kind": "text", "shapeId": "11", "size": 15, "maxChars": 117 },
        { "name": "card2Icon", "kind": "image", "imageKind": "icon", "shapeId": "15" },
        { "name": "card2Heading", "kind": "text", "shapeId": "16", "size": 19, "maxChars": 15 },
        { "name": "card2Text", "kind": "text", "shapeId": "17", "size": 15, "maxChars": 117 }
      ],
      "repeat": {
        "item": "card",
        "slots": ["card{n}Icon", "card{n}Heading", "card{n}Text"],
        "shapes": [["6"], ["12"]],
        "min": 2,
        "max": 4
      }
    }
  ]
}
```

## `TemplateManifest`

| Field | Type | Required | Meaning |
| --- | --- | --- | --- |
| `version` | `1` | yes | Always 1. |
| `name` | string | yes | A name for the template. Used in messages. |
| `direction` | `"ltr"` or `"rtl"` | yes | The reading direction the template was designed for. |
| `patterns` | `PatternDef[]` | yes | The reusable layouts. |
| `language` | string | no | BCP 47 tag of the template's sample text. Decks default to it. |
| `alignOnDirectionChange` | `"flip"` or `"keep"` | no | What happens to left and right paragraph alignment in a deck that reads the other way. `flip` (default) mirrors it. `keep` leaves it as designed. A text slot can override this. |
| `textScale` | number | no | Enlarges the body text of every deck made from this template, as a factor on the designed size: `1.3` turns 9 pt into about 12 pt. For a template whose body text is too small to read comfortably. Budgets shrink to match. See "Larger text" in [layout-freedom.md](layout-freedom.md). |
| `shrinkText` | `{ to, minSize? }` | no | Lets text that does not fit its box be set smaller: down to `to` of its size (`0.85` is 85%) and never under `minSize` points. Acts only for a template loaded with fonts. See "Shrinking text to fit" in [text-fit.md](text-fit.md). |

## `PatternDef`

| Field | Type | Required | Meaning |
| --- | --- | --- | --- |
| `id` | string | yes | What decks call the pattern. Unique in the manifest. Lower-case words joined by hyphens is the convention. |
| `sourceSlide` | integer | yes | Number of the template slide the pattern is cloned from, counted from 1. |
| `description` | string | yes | What the pattern is for. The planning model reads it to choose patterns, so write it for an author: what the slide holds and when to use it. |
| `slots` | `SlotDef[]` | yes | The shapes a deck fills. |
| `figures` | boolean | no | The pattern is built around numbers. The planner offers it only when the brief contains figures, so a model is not led to invent statistics. |
| `repeat` | `RepeatDef` | no | A row or grid of like items whose number a deck may choose. See below. |
| `sequences` | string[][][] | no | Rows of items with a fixed count, for mirroring in decks that read the other way. Each row lists its items first to last in the template's direction, and each item lists its slot names. Not needed for slots covered by `repeat`. |

## Slots: properties every kind has

| Field | Type | Required | Meaning |
| --- | --- | --- | --- |
| `name` | string | yes | What decks call the slot. Unique within the pattern. Letters and digits, starting with a lower-case letter. |
| `kind` | `"text"`, `"image"`, `"table"` or `"chart"` | yes | What the slot holds. |
| `shapeId` | string | yes | The id of the shape on the source slide (the `id` attribute of the shape's `p:cNvPr` element, as a string). `inspectTemplate` and `studioView` list them. |
| `description` | string | no | Guidance for the planning model about this slot, such as "One line". |
| `empty` | `"error"`, `"remove"`, `"keep"` or `"drop"` | no | What happens without a value. Default `error` (required). See [concepts.md](concepts.md). |

Slots with `empty` of `keep` or `drop` are never offered to the planning model.

A slot with `"empty": "keep"` and a `box` shows a part of the design in another place: the shape keeps its
content and is put at the box. Such a slot may point at any shape that is not inside a group (a drawn shape,
a line, a text, a group as a whole), with `"kind": "text"`, or `"kind": "image"` for a picture. When the box is
the shape's own size, only its position changes, so nothing inside a group is rescaled. The `derivePattern`
operation adds these slots, named `moved1`, `moved2`, ..., for shapes that a variant moves and that are not
slots themselves:

```json
{ "name": "moved1", "kind": "text", "shapeId": "14", "empty": "keep", "box": { "x": 0.52, "y": 0.3, "w": 0.4, "h": 0.05 } }
```

## Text slot (`kind: "text"`)

| Field | Type | Meaning |
| --- | --- | --- |
| `maxChars` | integer | Soft limit on the total characters of the slot. Exceeding it gives a warning. |
| `maxParagraphs` | integer | Soft limit on the number of paragraphs. When it is more than 1, the planner asks the model for a list of paragraphs. |
| `deckField` | string | Take the value from `deck.fields[deckField]` when the slide does not set it. |
| `size` | number | Font size of the slot's text on the template slide, in points. Informational: it decides whether a deck's `textScale` applies (see `scales`). |
| `scales` | boolean | Whether a deck's `textScale` applies to this slot. Default: yes for text of 12 pt and under that is not a deck field. |
| `fontSize` | number | Sets the font size outright, in points. For a derived pattern that uses the shape differently. |
| `box` | `{ x, y, w, h }` | Moves and resizes the text box, as fractions of the slide (0 to 1). Only for a shape that is not inside a group. |
| `paragraphGap` | number | Space after each paragraph, in points. |
| `alignOnDirectionChange` | `"flip"` or `"keep"` | Overrides the manifest's setting for this slot. |
| `styleRun` | `"first"` or `"last"` | Which run of the template paragraph supplies the formatting, when the template paragraph mixes styles (a bold label followed by plain text). Default `first`. |

## Image slot (`kind: "image"`)

| Field | Type | Meaning |
| --- | --- | --- |
| `imageKind` | `"photo"`, `"icon"` or `"graphic"` | Required. `photo`: a generated or supplied photo. `icon`: one of the app's icon set. `graphic`: a picture the app supplies, such as a diagram. |
| `fit` | `"cover"` or `"contain"` | `cover` crops the picture to fill the frame. `contain` shows the whole picture inside the frame's area. Default: `cover` for photos, `contain` for icons and graphics. |
| `box` | `{ x, y, w, h }` | The area the picture may use, when other than the sample picture's frame. Fractions of the slide. Only for a picture that is not inside a group. |

## Table slot (`kind: "table"`)

The shape must be a table.

| Field | Type | Meaning |
| --- | --- | --- |
| `maxColumns` | integer | Soft limit on columns. |
| `maxRows` | integer | Soft limit on body rows (the header row is not counted). |
| `maxCellChars` | integer | Soft limit on the characters in one cell. |
| `headerText` | `{ size?, bold?, color?, font? }` | Formatting for header cells, where the sample table defines none or it should differ. `color` is hex without `#`. |
| `bodyText` | `{ size?, bold?, color?, font? }` | The same for body cells. |

## Chart slot (`kind: "chart"`)

The shape is the sample picture of a chart. At render time it is replaced by a native chart in the same place.
The shape must not be inside a group.

| Field | Type | Meaning |
| --- | --- | --- |
| `box` | `{ x, y, w, h }` | The area the chart may use, when other than the sample picture's frame. |
| `maxCategories` | integer | Soft limit on categories. |
| `maxSeries` | integer | Soft limit on series. |
| `colors` | string[] | Series colours in order, hex without `#`. They repeat when there are more series. |
| `text` | `{ size?, bold?, color?, font? }` | Formatting of axis labels, data labels and the legend. |
| `lineColor` | string | Colour of gridlines and axis lines, hex without `#`. |

## `RepeatDef`: a row or grid of items

```json
"repeat": {
  "item": "card",
  "slots": ["card{n}Icon", "card{n}Heading", "card{n}Text"],
  "shapes": [["6"], ["12"], ["18"], ["24"]],
  "min": 2,
  "max": 4,
  "span": { "left": 0.04, "right": 0.96 },
  "columns": 3,
  "widen": 1.5
}
```

| Field | Type | Required | Meaning |
| --- | --- | --- | --- |
| `item` | string | yes | What one item is called in prompts, such as `card` or `figure`. |
| `slots` | string[] | yes | Slot name templates of one item. `{n}` stands for the item number, from 1. For every item drawn on the slide the pattern must have a slot with each of these names (`card1Icon`, `card2Icon`, ...). |
| `shapes` | string[][] | yes | For each item drawn on the template slide, first in reading order first: the ids of the shapes that move together as that item. Usually the group around a card. A background shape that sits outside the group belongs in the list too. |
| `min`, `max` | integer | no | Fewest and most items a deck may ask for. Both default to the number drawn. |
| `span` | `{ left, right }` | no | Left and right limits of the row, as fractions of the slide width. Defaults to the row as drawn. Needed when `max` is more than the drawn items can fit in the drawn width. Only applies to shapes placed directly on the slide. |
| `columns` | integer | no | For a grid: how many items one line holds on the template slide. Leave it out for a single row. A grid cannot hold more items than it draws. |
| `widen` | number | no | Lets items grow wider when a line holds fewer of them than the template draws, so that they fill the row and do not stand in the middle of it. The value is the most an item may grow, as a factor on its drawn width: `1.5` lets a card become half as wide again. Backgrounds and text boxes widen; pictures keep their size. The budgets of the row's texts grow to match. Left out, items keep their drawn size. |

How the items are laid out for a given count is described in [layout-freedom.md](layout-freedom.md).

Slots of copied items do not appear in the manifest. They are derived: item 5 of a row that draws 4 gets copies
of item 4's slots, named with the next number.

## Rules a manifest must satisfy

`checkManifest(template)` reports violations of these:

- Every `sourceSlide` exists in the pptx.
- Every `shapeId` is a shape on that slide, and is the right kind of shape for the slot (a text slot points at
  a shape with text, an image slot at a picture, a table slot at a table). A slot with `"empty": "drop"` may
  point at any shape, because the shape is only ever removed; so may a slot with `"empty": "keep"` and a `box`,
  because the shape is only ever moved.
- Pattern ids are unique. Slot names are unique within a pattern.
- A `box` is not set on a shape inside a group. A chart slot's shape is not inside a group.
- For a `repeat`: a slot exists for every name template and every drawn item; every shape id exists; all items
  sit directly on the slide or all in the same group; `min` is at least 1 and not more than `max`; `widen`,
  when set, is a number of 1 or more.

It also warns about text slots without `maxChars` and about two slots that point at the same shape.

## Finding shape ids

```ts
import { inspectTemplate, PptxPackage } from 'slotdeck';

const info = inspectTemplate(PptxPackage.load(pptxBytes));
for (const shape of info.slides[4].shapes) {
  // slide 5: index 4
  console.log(shape.id, shape.kind, shape.box, shape.groupId ?? '', shape.paragraphs?.join(' / ') ?? '');
}
```

`shape.box` is `{ x, y, w, h }` as fractions of the slide, already worked out through any groups. `shape.kind`
is `text`, `picture`, `shape` (a drawn shape without text), `line`, `group`, `table` or `other`.

## Writing good descriptions

The planning model chooses a pattern by its `description` and its slots' names and limits. A description should
say what the slide holds and when to use it, in one or two sentences, without numbers that the deck may change:

- Good: "A row of cards, each with an icon, a heading and a short text: features, steps or categories."
- Poor: "Four cards." (the count can vary, and it does not say what the cards are for)
