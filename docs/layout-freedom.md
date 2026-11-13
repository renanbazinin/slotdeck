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
- Without `items`, the slide has the number of items drawn on the template slide.

An `items` value outside `min` to `max` is an error:
`pattern "icon-cards" takes 2 to 4 cards, not 5`.

How many items fit is a property of the template: `changeManifest` refuses a `max` that does not fit side by
side and says the most that do (see [template-studio.md](template-studio.md)).

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

With `changeManifest` this is three operations (`copyPattern`, `addSlot` with role `decoration`, `setSlot` with
a `box`), and the budgets are worked out for you: see [template-studio.md](template-studio.md).

## What is not possible

- Moving a shape to an arbitrary place from a deck. Positions come only from the template and the manifest.
- Swapping the side of a photo.
- A grid with more items than the template slide draws.
- New kinds of slides that the template has no sample for.
