# Deck format

A deck is a JSON object of type `DeckSpec`. It describes one presentation and contains no binary data.

```ts
import type { DeckSpec, SlideSpec, SlotValue, ImageRef, TableValue, ChartValue } from 'slotdeck';
```

## Complete example

```json
{
  "version": 1,
  "title": "A year of Riverbend Bike Share",
  "language": "en",
  "direction": "ltr",
  "fields": { "organizationDepartment": "Riverbend Bike Share  |  Season review" },
  "slides": [
    {
      "pattern": "cover",
      "slots": { "title": "A year of Riverbend Bike Share", "subtitle": "What the 2025 season taught us" }
    },
    {
      "pattern": "text-and-photo",
      "slots": {
        "title": "Bikes became part of the daily commute",
        "body": ["Most rides used to be weekend outings.", "This year the busiest hour is eight on a weekday morning."],
        "photo": { "asset": "photo-commute", "alt": "A commuter on a river path" }
      }
    },
    {
      "pattern": "icon-cards",
      "items": 3,
      "slots": {
        "title": "Three changes for next season",
        "card1Icon": { "asset": "clock" },
        "card1Heading": "Longer hours",
        "card1Text": "Bikes unlock from five in the morning.",
        "card2Icon": { "asset": "shield" },
        "card2Heading": "Safer docks",
        "card2Text": "Every station gets lighting and a camera.",
        "card3Icon": { "asset": "people" },
        "card3Heading": "Group passes",
        "card3Text": "One pass for up to five riders."
      }
    },
    {
      "pattern": "table",
      "slots": {
        "title": "The busiest stations",
        "table": {
          "header": ["Station", "Docks", "Rides a day"],
          "rows": [["Tram Square", "24", "410"], ["River Gate", "18", "365"]]
        }
      }
    },
    {
      "pattern": "chart-native",
      "slots": {
        "title": "Rides grew through the summer",
        "chart": {
          "type": "column",
          "categories": ["Apr-May", "Jun-Jul", "Aug-Sep"],
          "series": [
            { "name": "2024", "values": [5, 9, 11] },
            { "name": "2025", "values": [7, 15, 19] }
          ]
        },
        "chartDescription": "Rides in thousands for each part of the season."
      }
    },
    { "pattern": "closing", "slots": { "title": "Thank you" } }
  ]
}
```

The pattern ids and slot names in this example come from the demo template's manifest. Yours come from your own
manifest.

## `DeckSpec`

| Field | Type | Required | Meaning |
| --- | --- | --- | --- |
| `version` | `1` | yes | Always the number 1. |
| `slides` | `SlideSpec[]` | yes | The slides in order. Must not be empty. |
| `title` | string | no | A name for the deck. Not placed on any slide by itself. |
| `brief` | string | no | The request the deck was written from. `editDeck` uses it as background. |
| `language` | string | no | BCP 47 tag of the deck's text, such as `en` or `he`. |
| `direction` | `"ltr"` or `"rtl"` | no | Reading direction. Defaults to the template's. See [languages.md](languages.md). |
| `fields` | object of strings | no | Values shared by every slide, by field name. See "Deck fields" below. |
| `textScale` | number | no | Size of body text as a factor on the template's: `1` as designed, `1.25` large, `1.5` extra large. See [layout-freedom.md](layout-freedom.md). |
| `layouts` | `PatternDesign[]` | no | Layouts made for this deck (see "New layouts for a deck" in [generating.md](generating.md)). Slides use them by their `id`, like the template's patterns. |

## `SlideSpec`

| Field | Type | Required | Meaning |
| --- | --- | --- | --- |
| `pattern` | string | yes | The `id` of a pattern in the manifest. |
| `slots` | object | yes | Slot name to value. |
| `items` | integer | no | Only for a pattern with a row or grid of items: how many items this slide has. Defaults to the number drawn on the template slide. Must be within the pattern's range. |

## Slot values

The value depends on the slot's kind in the manifest.

### Text slot

A string, or an array of strings with one entry per paragraph.

```json
"title": "Deliveries are faster"
"body": ["First paragraph.", "Second paragraph."]
"body": "First paragraph.\nSecond paragraph."
```

- A newline inside a string starts a new paragraph, so the last two lines above mean the same.
- Empty paragraphs are dropped.
- Each paragraph takes its formatting (font, size, colour, bullet) from the template paragraph in the same
  position. When there are more paragraphs than the template slot has, the last template paragraph's formatting
  is reused. This is how a bulleted list works: the template shape has bulleted paragraphs, and the deck gives
  one string per bullet.
- Do not put bullet characters, numbering or Markdown in the text. The text is used literally.
- Limits: `maxChars` counts all characters of all paragraphs together; `maxParagraphs` counts paragraphs. Both
  are soft (a warning, not an error).

### Image slot

An `ImageRef`:

```json
"photo": { "asset": "photo-commute", "alt": "A commuter on a river path" }
```

| Field | Type | Required | Meaning |
| --- | --- | --- | --- |
| `asset` | string | yes | The id of the picture. The bytes are passed to `renderDeck` in `options.assets` under this id. |
| `alt` | string | no | Alternative text for screen readers, stored on the picture. |
| `prompt` | string | no | A description the picture can be generated from. Written by the planner for photos; lets the picture be made again later. |

How the picture is placed depends on the slot (`fit` in the manifest):

- `cover` (the default for photos): the picture fills the frame and is cropped at the edges, keeping its
  proportions. The frame's shape (rounded corners, a circle) is kept.
- `contain` (the default for icons and graphics): the whole picture is shown, centred, inside the frame's
  area, keeping its proportions.

Supported formats are PNG, JPEG and GIF.

### Table slot

A `TableValue`:

```json
"table": { "header": ["Station", "Docks"], "rows": [["Tram Square", "24"], ["River Gate", "18"]] }
```

- `header` sets the number of columns. Every row must have exactly that many cells, or the deck is invalid.
- Cells are strings. Numbers must be written as strings (`"24"`).
- The template table's first row supplies the header style and its following rows supply the body styles in
  turn (so banding is kept). Columns share the table's width equally.
- Limits (`maxColumns`, `maxRows`, `maxCellChars`) are soft.

### Chart slot

A `ChartValue`, which becomes a native PowerPoint chart that can be changed with Edit Data:

```json
"chart": {
  "type": "column",
  "categories": ["Q1", "Q2", "Q3"],
  "series": [{ "name": "Revenue", "values": [12, 15, 19] }],
  "format": "number"
}
```

| Field | Type | Required | Meaning |
| --- | --- | --- | --- |
| `type` | `"column"`, `"bar"`, `"line"` or `"pie"` | yes | `column` has vertical bars, `bar` horizontal ones. |
| `categories` | string[] | yes | Labels along the axis, or the names of the pie slices. At least one. |
| `series` | `{ name, values }[]` | yes | At least one. Each `values` array must have exactly one number per category. A pie takes exactly one series. |
| `format` | `"number"` or `"percent"` | no | `percent` shows values with a percent sign (45 is shown as 45%). Any other value is ignored with a warning. |

Bars start at zero. Colours and fonts come from the manifest.

### Leaving a slot out

- Omit the slot, or give an empty string: the slot has no value, and its `empty` policy decides what happens
  (see [concepts.md](concepts.md)).
- Give `null`: the slot is explicitly left out and its shape is removed. This needs `layout: 'adaptive'` and a
  slot that is not required.

A name in `slots` that the pattern does not have is ignored with a warning.

## Deck fields

A text slot whose manifest entry has `"deckField": "x"` takes its value from `deck.fields.x` when the slide does
not give that slot a value itself. Use it for text that repeats on every slide.

```json
{ "fields": { "organizationDepartment": "Acme  |  Operations" }, "slides": [ ... ] }
```

If a required slot depends on a field that the deck does not give, validation reports `this slot is required`
for that slot on every slide that has it.

## Items of a row or grid

For a pattern with a row or grid of items, `items` says how many. The slots of each item are numbered:
`card1Heading`, `card2Heading`, and so on. With `"items": 3`, give slots for items 1 to 3 only. With more items
than the template slide draws, the extra items are copies of the last one and have the next numbers
(`card5Heading`, ...).

The first item is the first in the deck's reading direction. In a right-to-left deck on a left-to-right
template, `card1...` appears at the right.

## A deck's own layouts

A deck planned with `composeLayouts` carries the designs of the layouts that were made for it in `layouts`: the
same `PatternDesign` objects that `composePattern` returns (see [template-studio.md](template-studio.md)). Its
slides name them in `pattern` and fill their slots like any other.

`renderDeck`, `editDeck`, the planner and the reviews build these layouts onto the template themselves, the same
way every time, so the stored deck JSON is all that is needed. To check such a deck yourself, check it against
`withLayouts(template, deck.layouts).manifest`: the bare manifest does not have the deck's layouts, and
`validateDeck` says so. When the template gains a layout of the same id, as when the design was added to it in a
studio, the template's own is used. An edit that removes the last slide using a layout removes its design too.

## Checking a deck

```ts
import { validateDeck } from 'slotdeck';

const issues = validateDeck(manifest, deck, { layout: 'adaptive', assets: ['photo-commute', 'clock'] });
// [{ severity: 'error' | 'warning', slide?: number, slot?: string, message: string }]
```

Errors make `renderDeck` throw. Warnings do not. Every message is listed in
[troubleshooting.md](troubleshooting.md).

## Another pattern for a slide

`slideAlternatives(manifest, deck, slideNumber)` lists the other patterns that can hold what a slide says, and
returns the slide written for each: cards as a list, a slide as its mirror image. It changes no words and
calls no model. See "Another layout for the same slide" in [layout-freedom.md](layout-freedom.md).
