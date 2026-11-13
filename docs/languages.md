# Languages and reading direction

A template is designed for one reading direction (`direction` in the manifest: `ltr` or `rtl`). A deck states
its own (`direction` in the deck, defaulting to the template's). When they differ, the text flips and the layout
stays.

This works both ways: a Hebrew or Arabic deck on a left-to-right template, and an English deck on a
right-to-left template.

## Setting language and direction

In a deck:

```json
{ "version": 1, "language": "he", "direction": "rtl", "slides": [ ... ] }
```

In a brief for the planner, `language` is enough: the direction follows from it.

```ts
await generateDeck(template, { prompt: '...', language: 'he' }, options);
```

`directionOf('he')` returns `'rtl'`; `directionOf('en')` returns `'ltr'`.

## What changes in a deck that reads the other way

| Thing | What happens |
| --- | --- |
| Paragraph direction | Every paragraph gets the direction of its own text, from its first letter. A Hebrew deck may hold an English line, and that line is left to right. Text with no letters ("78%") keeps the template paragraph's direction. |
| Alignment | Left and right alignment are swapped (`alignOnDirectionChange: "flip"`, the default). A paragraph that states no alignment counts as left-aligned. Centred text stays centred. |
| Rows and grids of items | Filled mirror-wise within each line, so the deck's first item is where its readers start: at the right in a right-to-left deck. Shapes do not move; the values go to the mirrored positions. |
| Tables | Run in the deck's direction, with text aligned to its start. |
| Charts | Labels are set right to left. |
| Run language | Right-to-left text is marked with a right-to-left language. |
| Shapes | Do not move. A photo on the right stays on the right. Bullets and their indent follow the paragraph direction automatically. |

## Keeping alignment

Mirroring alignment suits templates whose text boxes are as wide as the area they belong to. It does not suit a
template whose text boxes share one edge but not the other: flipped text would then start at uneven positions or
run into a neighbouring shape.

- For the whole template: `"alignOnDirectionChange": "keep"` in the manifest.
- For one slot: `"alignOnDirectionChange": "keep"` (or `"flip"`) on that text slot.

A common case for a per-slot `keep` is a line next to a logo: it should stay beside the logo in either
direction.

## Why runs are marked with a language

PowerPoint decides where punctuation and spaces go by the language of the run they are in, not only by the
paragraph's direction. In a run marked as English, the text

```
שאלות: name@example.org
```

is displayed with the colon on the far side of the address. slotdeck therefore marks right-to-left runs with a
right-to-left language (`he-IL` for Hebrew, `ar-SA` for Arabic) unless the template already gave the run one.
You do not need to do anything for this.

## Fonts

Text keeps the template's fonts. If the template's font has no glyphs for the deck's script, PowerPoint
substitutes another font. For a template that will be used in two scripts, set a complex-script font on its text
(in PowerPoint: the font for "complex scripts"), or choose fonts that cover both.

## Mirrored item order in the deck JSON

The deck always numbers items in its own reading order: `card1` is the first card for the reader. You do not
reverse anything yourself. Validation messages name the slot as the deck wrote it.

## Language of generated text

The planner writes in `brief.language`. Slot names and pattern ids stay as the manifest defines them (they are
identifiers, not content). Translating an existing deck is an edit:

```ts
await editDeck(template, deck, 'Translate the deck to English', { llm });
```

This rewrites every slide, sets the direction, and translates deck-level fields.
