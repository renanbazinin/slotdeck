# Layout freedom

A deck can depart from the template slide in four bounded ways. In each, the choice is a small value in the deck
JSON, the manifest says what is allowed, and the code moves the shapes. Nothing is ever positioned by a model.

## 1. A different number of items

A pattern with a `repeat` in the manifest has a row or grid of like items: cards, figures, steps. The deck says
how many with `items`.

Manifest:

```json
"repeat": {
  "item": "card",
  "slots": ["card{n}Icon", "card{n}Heading", "card{n}Text"],
  "shapes": [["6"], ["12"], ["18"], ["24"]],
  "min": 2,
  "max": 4
}
```

Deck:

```json
{ "pattern": "icon-cards", "items": 3, "slots": { "title": "...", "card1Icon": {"asset": "clock"}, "card1Heading": "...", "card1Text": "...", "card2Icon": {"asset": "shield"}, "card2Heading": "...", "card2Text": "...", "card3Icon": {"asset": "people"}, "card3Heading": "...", "card3Text": "..." } }
```

What the renderer does:

- **Fewer items than drawn:** the surplus items' shapes are removed.
- **More items than drawn:** the last item is copied. A copy's slots have the next numbers (`card5Icon`,
  `card5Heading`, ...) and take the last drawn item's limits.
- **Spacing:** the items are laid out again around the centre of the area the row occupied. They keep the
  template's spacing when there is room and move closer together when there is not. With a `span`, the row may
  spread out to those limits.
- **Size:** every item keeps its drawn size, unless the row has `widen` (see "Fewer items that fill the row"
  below).
- Without `items`, the slide has the number of items drawn on the template slide.

An `items` value outside `min` to `max` is an error:
`pattern "icon-cards" takes 2 to 4 cards, not 5`.

How many items fit is a property of the template: `changeManifest` refuses a `max` that does not fit side by
side and says the most that do (see [template-studio.md](template-studio.md)).

### Fewer items that fill the row

By default an item keeps its drawn size, so three cards in a row drawn for four stand in the middle with room
left at both ends. With `widen` on the `repeat`, a line that holds fewer items than the template draws has its
items made wider, so that they fill the row:

```json
"repeat": {
  "item": "card",
  "slots": ["card{n}Icon", "card{n}Heading", "card{n}Text"],
  "shapes": [["6"], ["12"], ["18"], ["24"]],
  "min": 2,
  "max": 4,
  "widen": 2.2
}
```

`widen` is the most an item may grow, as a factor on its drawn width: `1.5` lets a card become half as wide
again, `2.2` lets two cards fill a row drawn for four. With `changeManifest`:
`{ op: 'setRepeat', pattern: 'icon-cards', widen: 2.2 }`; `null` or `1` turns it off.

- **How wide.** The items share the row's span (`span`, or the row as drawn) and keep the template's own gap
  between them. Where that is more than the limit allows, they stop at the limit and the row is centred. In
  the demo template a card is 21% of the slide wide; three become 29% each and two become 45%.
- **What decides.** Only the number of items. Long text does not widen a card and short text does not narrow
  one, so the same deck always looks the same.
- **Inside an item** each shape is treated on its own, whether the item is a group or not:
  - A background, a rule or a text box that spans the item grows with it and keeps its distance to both
    sides. The text boxes of the row's own slots always grow.
  - A picture or an icon keeps its size. So does a circle or any other outline that is not a rectangle, and
    a slanted line. It keeps its distance to the nearer side of the item: an icon at the start of a card
    stays at the start, and one in the middle stays in the middle.
  - Rounded and cut corners keep the size they were drawn with.
  - Heights do not change.
- **Budgets.** A wider text box holds more. The character budget of the row's text slots grows by the factor
  the items are at least wider by: the number drawn on a line divided by the number held, up to `widen`. In a
  row drawn for four with `widen: 2.2`, a 117-character text holds 156 with three cards and 234 with two.
  `validateDeck` uses these budgets, the planner gives them to the model for the number of items it chose,
  and `sampleDeck` with `items: 'min'` fills them. Paragraph budgets do not change. A deck's text size is
  applied on top.
- **Grids.** A grid has one item width, decided by its fullest line. In a three-column grid, four figures make
  two lines of two, which widen; five make three and two, and keep their width.
- **Both reading directions.** Positions come from the template: a right-to-left template with its icons at
  the right edge of each card keeps them there. A deck that reads the other way gets the same shapes, with
  its first item at the other end.
- **Nothing else changes.** At the drawn number of items and above, and for every row without `widen`, the
  output is byte for byte what it was.

## 2. Grids

With `columns` in the `repeat`, the items form a grid. `columns` is how many items one line holds on the
template slide.

- Items fill lines in reading order and use as few lines as the count needs.
- They are spread evenly over those lines: in a three-column grid, four items make two lines of two (not three
  and one), five make three and two, three make one line.
- Each line is centred, and a single line moves to the middle of the grid's area.
- A grid cannot hold more items than it draws.

## 3. Larger text

A deck can ask for larger body text:

```json
{ "version": 1, "textScale": 1.25, "slides": [ ... ] }
```

or, on a brief for the planner, `textSize: 'large'` (1.25) or `'x-large'` (1.5). The named sizes are in
`TEXT_SIZES`: `{ normal: 1, large: 1.25, 'x-large': 1.5 }`.

- It applies to **body text**: by default, text slots whose template size is 12 pt or under and that are not
  deck fields. Titles and large figures keep their size. A slot can opt in or out with `scales` in the
  manifest.
- Boxes do not grow, so larger text fits less. The character budget of a scaled slot is divided by the square
  of the scale (a 300-character slot holds 192 at 1.25), and the paragraph budget by the scale. The planner
  shows the model the reduced budgets, so it writes less.
- Through `editDeck`, "make the text larger" sets the scale and rewrites only the slides whose text no longer
  fits.
- Table text and chart labels follow the scale where the manifest states their size (`bodyText.size`,
  `headerText.size`, `text.size`), and a table cell then holds less.

### Larger text for a whole template

The engine keeps the template's own font sizes. A template designed with small body text, such as 9 pt on a
10-inch-wide slide, therefore gives decks with small body text. To change that once, for every deck, set
`textScale` on the manifest:

```json
{ "version": 1, "name": "corporate", "direction": "rtl", "textScale": 1.3, "patterns": [ ... ] }
```

With `changeManifest`: `{ op: 'setManifest', textScale: 1.3 }`.

- It works exactly like a deck's `textScale` and applies to the same slots (body text: 12 pt and under by
  default), but for every deck made from the template.
- Budgets shrink to match, in validation, in what the planner shows the model, and in `sampleDeck`: a slot
  that held 240 characters holds 142 at 1.3. Render the sample deck after setting it to see that everything
  still fits.
- A deck's own text size goes on top. With a template at 1.3, a deck at `large` (1.25) shows body text at about
  1.6 times the designed size.
- A slot that should keep its size can opt out with `"scales": false`.

## 4. Derived patterns

Several patterns may point at the same template slide and use it differently. A derived pattern is made by
copying a pattern and changing some of its slots:

| Change | Manifest property | Example |
| --- | --- | --- |
| Remove a shape always | `"empty": "drop"` on a slot that points at it | A text-only slide from a slide with a photo |
| Move or resize a text box | `box` on the text slot | Widen the text over the area the photo left free |
| Set a font size | `fontSize` | Larger body text for a slide with little on it |
| Space paragraphs | `paragraphGap` | Air between two or three short paragraphs |
| Give a picture another area | `box` on the image slot | A wide diagram in place of a small sample chart |
| A native chart in place of a picture | a `chart` slot on the picture's shape | The chart slide with editable data |

Example: a text-only pattern derived from a photo pattern.

```json
{
  "id": "text-only",
  "sourceSlide": 2,
  "description": "A title, an optional supporting line and two or three short paragraphs, with no picture.",
  "slots": [
    { "name": "photo", "kind": "image", "imageKind": "graphic", "shapeId": "2", "empty": "drop" },
    { "name": "title", "kind": "text", "shapeId": "5", "size": 34, "maxChars": 83, "box": { "x": 0.0375, "y": 0.2, "w": 0.925, "h": 0.187 } },
    { "name": "body", "kind": "text", "shapeId": "7", "size": 18, "maxChars": 331, "maxParagraphs": 3, "fontSize": 20, "paragraphGap": 12, "box": { "x": 0.0375, "y": 0.4933, "w": 0.72, "h": 0.3867 } }
  ]
}
```

By hand, with `changeManifest`, this is three operations (`copyPattern`, `addSlot` with role `decoration`,
`setSlot` with a `box`), and the budgets are worked out for you: see [template-studio.md](template-studio.md).
For the three common variants below, the boxes are worked out as well.

### Variants the code works out

Three kinds of derived pattern need no box typed in. One `changeManifest` operation makes each, and
`editManifest` lets a model ask for one by name:

```ts
changeManifest(template, [
  { op: 'derivePattern', pattern: 'text-and-photo', variant: 'mirror' },      // text-and-photo-mirrored
  { op: 'derivePattern', pattern: 'text-and-photo', variant: 'text-only' },   // text-and-photo-text-only
  { op: 'derivePattern', pattern: 'text-and-photo', variant: 'without', slot: 'subtitle' }, // text-and-photo-no-subtitle
]);
```

Each makes a new pattern next to its source. A deck uses it by its id like any other pattern, and the planner
chooses it by its description; nothing changes in the deck format. `patternVariants(template, id)` says which
variants a pattern allows before you ask.

**How the shapes of a slide are sorted.** The code looks at the shapes that sit directly on the slide (a group
counts as one shape) and puts each in one of these:

| What | Which shapes | In a variant |
| --- | --- | --- |
| Locked | The logo, rules and other elements that repeat across the template's slides; page-number, date and footer placeholders; texts that are deck fields. | Never moved. |
| The picture | The pattern's one photo, graphic or chart slot. Icons, pictures inside a row of items and pictures kept as decoration do not count. | Changes sides, or is left out. |
| Beside the picture | Every shape that shares no width with the picture: titles, texts, panels. | Moves as one block, so edges the shapes share stay shared. |
| With the picture | Shapes in the picture's own column: a sticker on it, a caption under it. | Goes where the picture goes. |
| Around both | A title above the picture and its text, a caption under both, a background behind both. | Stays. |

#### `mirror`

The picture and the block beside it change sides.

- They change places inside the room the two share, so the slide keeps its margins. A picture that runs off
  one edge of the slide runs off the other edge instead, and the block keeps the margin it had.
- Nothing changes size, so no budget changes. The block is moved, not flipped: text boxes that shared a left
  edge still share one.
- Locked shapes stay. When the picture would cover one near the top or the bottom of the slide (a full-height
  photo landing on the logo), the picture is cut back to start a little under it, or end a little above it,
  and the photo is cropped to the new frame.
- Shapes that are not slots, and groups, are moved too. The operation adds a slot for each, named `moved1`,
  `moved2`, ..., with `"empty": "keep"` and a `box`: the shape keeps its content and is shown at the box.
- A mirrored pattern is a layout of its own that a deck or the planner picks. A deck's reading direction still
  never moves a shape: a right-to-left deck made from `photo-stat` has its photo where the template drew it.

In the demo template, `text-and-photo-mirrored`, `photo-stat-mirrored` and `list-with-photo-mirrored` were made
this way.

#### `text-only`

The picture is left out and the text beside it widens.

- The picture's slot stays under its name with `"empty": "drop"`, so its shape is always removed. Decoration
  drawn on the picture is removed with it. Texts the deck writes are never removed.
- Every text slot beside the picture grows towards the side the picture was on: up to the picture's far edge,
  or, for a picture that ran off the slide, up to the margin the text keeps on its own side. All of them end
  on the same edge.
- The column is no wider than about 90 characters a line of the slot with the largest budget, because lines
  that run the width of a slide are hard to read. It keeps the edge where the template's reading starts (the
  left edge of a left-to-right template) and gives up room at the other.
- A text stops short of any other shape that is in its way.
- The budget of each widened text is measured again for its new width.
- To enlarge the text of the new pattern, follow with `setSlot` and a `fontSize`; the budget follows.

#### `without`

One optional text is left out for good, and the room it took is used.

- `slot` names the text. It must be optional (`"empty": "remove"`). In the variant it has `"empty": "drop"`.
- The shapes under it in its column move up, so that the first of them starts where the text started.
- The lowest text among them keeps its lower edge, so it becomes taller and its budget grows. This is done
  only when its text starts at the top of its box; a text that is centred or sits at the bottom of its box is
  moved and keeps its size, so that the text does not shift inside it.
- When nothing sits under the text, the text above it is extended down to where the text ended, on the same
  condition.
- When there is nothing to move or extend, the variant is the same slide with the slot gone.

Why make one when the slot is optional anyway: in `strict` layout an optional slot has to be filled, and in
`adaptive` layout leaving it empty leaves its room empty. The variant closes the gap and lets the text under it
say more.

#### When a variant is refused

`derivePattern` refuses, with the reason, when the slide does not allow the variant:

| Reason | Example |
| --- | --- |
| The pattern has no picture of its own. | `"cover" has no photo, graphic or chart of its own, so there is nothing to change sides with` |
| It has more than one. | `"gallery" has 2 pictures (photo1, photo2); this needs a slide with one picture beside its text` |
| Nothing sits beside the picture. | A wide diagram under a title: `nothing sits beside the picture "diagramGraphic", ...` |
| The picture has content on both of its sides. | A picture in the middle of the slide. |
| A shape that keeps its place is where something would go. | `shape #9 ("Source: annual report"), which keeps its place, is where shape #7 would go` |
| A row of items would have to move. | The renderer lays a row out again for every item count, so a box cannot place it. |
| A table would have to move. | A table slot has no `box`. |
| The text to widen is inside a group. | `the text beside the picture is inside a group, and a shape inside a group cannot take a box` |
| The slot to leave out is not an optional text. | `"title" is required; only an optional text can be left out; "text-and-photo" has "subtitle"` |

Render the sample deck after deriving a variant (`sampleDeck`, then `renderDeck`): it holds every pattern,
variants included, filled to its budgets.

## Another layout for the same slide

The sections above are about what a deck may ask for. `slideAlternatives` answers another question: a slide
already has its content; which other patterns could show it? It puts the content into every pattern that has a
place for all of it and returns the slide written for each. It uses no model and changes no words, so an
alternative says exactly what the slide said.

```ts
import { slideAlternatives } from 'slotdeck';

const alternatives = slideAlternatives(manifest, deck, 5, { layout: 'adaptive' }); // slide 5 of the deck, counted from 1
// [
//   {
//     pattern: 'text-only',
//     description: 'A title, an optional supporting line and two or three short paragraphs, with no picture. ...',
//     slide: { pattern: 'text-only', slots: { title: 'Three changes for next season', body: ['Longer hours: Bikes unlock ...', '...', '...'] } },
//     changes: ['The 3 cards became a list of 3 points, each written as "heading: text".'],
//     dropped: ['card1Icon', 'card2Icon', 'card3Icon'],
//     warnings: [],
//   },
// ]

deck.slides[4] = alternatives[0].slide; // take one, then validate and render as usual
```

| Field | Meaning |
| --- | --- |
| `pattern`, `description` | The pattern's id and what it is for, from the manifest. |
| `slide` | The slide written for that pattern: put it in the deck in place of the original. |
| `changes` | What was done to the content where it was not copied slot for slot, a sentence each. Empty for a plain copy. |
| `dropped` | Slots of the original slide whose photo or icon the pattern has no place for. |
| `warnings` | What `validateDeck` says about the new slide, such as `text is 30 characters; the slot fits about 20`. Empty when everything fits. |

Options: `layout` (`'strict'` by default, as in `renderDeck`) and `patterns` (consider only these ids). The list
is ordered: alternatives that fit and leave nothing out come first, then those that leave out a photo or
icons, then those with a warning. A slide with no alternative gives `[]`.

### How the content of a slide is read

A manifest does not say what a slot is for, so the code tells from names and limits:

| Part | Which slot |
| --- | --- |
| Title | The text slot whose name ends in `title` (`title`, `slideTitle`) and is not a subtitle. |
| Subtitle | The text slot whose name ends in `subtitle` or `subheading`. |
| Figure | A big, short text: a budget of at most 10 characters at 24 pt or more, such as `87%`. |
| Body | Of the text slots that are left, the one with the largest budget. It counts as a list when it takes several paragraphs (`maxParagraphs`). |
| Picture | The pattern's one photo or graphic outside its row. |
| Items | The row or grid (`repeat`). In each item: an icon, a picture, a figure, and of its other texts the one with the smaller budget is the heading and the one with the larger is the text. |
| Everything else | A takeaway, a footnote, a table, a chart. It keeps its name and fits only a slot of the same name and kind. |

Deck fields and slots with `"empty": "keep"` or `"drop"` take no part.

### What can become what

- **The same part.** A title goes to the title, a subtitle to the subtitle, a figure to the figure, the body to
  the body, a picture to a picture of the same kind.
- **Variants of one slide.** Between a pattern and its mirror image, or any two patterns whose slots have the
  same names, every value goes to the slot of the same name.
- **Items to items.** Cards become columns, a row becomes a grid, when the target takes that many items.
  Headings go to headings and texts to texts. Items with a heading and a text fit items with one text, written
  `heading: text`; single texts written that way fit items with a heading and a text, split at the first colon
  (or dash between spaces).
- **Items to a list.** One point for each item, written `heading: text`, in a body that takes several
  paragraphs. Each part of an item must be one paragraph.
- **A list to items.** One item for each point, from two points up. For items with a heading and a text, only
  when every point is written `heading: text`.
- **A list to paragraphs and back.** The words are the same; whether they show as points is the new slot's own
  formatting. `changes` says so.

### When a pattern is an alternative

- **Everything the slide says has a place**: every text, table, chart and supplied graphic. Nothing is cut and
  nothing is run together, apart from `heading: text`.
- **Everything the pattern requires can be filled** from the slide. In `strict` layout that includes its
  optional slots; in `adaptive` layout those may stay empty. Code does not invent a subtitle, a heading, a
  number, an icon or a photo, so a pattern that needs one the slide lacks is not offered.
- **A photo or an icon may be left out**, because it illustrates. `dropped` names the slots.
- **A figure stays a figure.** Figures do not become headings of cards, and headings do not become figures.
- **A text that is too long for its new slot does not rule a pattern out.** It is reported in `warnings`, and
  the alternative is listed after those that fit.

### Limits

- It judges what fits, not what suits. A section slide is offered for a cover, because both hold a title and
  a subtitle. `description` is there for a person or a model to choose by.
- The parts are told from slot names and limits. A title slot called `heading`, or an item with three texts,
  is not recognised; such a slide still gets the alternatives whose slots have the same names.
- It is not offered by the chat tools yet. Call it from your own code.

## What is not possible

- Moving a shape to an arbitrary place from a deck. Positions come only from the template and the manifest.
- Mirroring a slide because of a deck's reading direction. A mirrored layout is a separate pattern
  (`derivePattern` with `mirror`) that a deck names or the planner chooses.
- Deriving a variant that would move a row or grid of items, or a table.
- Widening a picture inside an item, or making an item taller. A widened item keeps its height, and its
  pictures keep their size.
- Letting the length of the text decide the size of an item. Only the number of items does.
- An alternative that needs words the slide does not have. `slideAlternatives` never writes text; ask a model
  for that with `editDeck`.
- A mirror or text-only variant of a slide with several pictures, or with a picture between two texts.
- A `box` for a shape inside a group. A variant moves a group as a whole and never resizes what is in it.
- A grid with more items than the template slide draws.
- New kinds of slides that the template has no sample for.
