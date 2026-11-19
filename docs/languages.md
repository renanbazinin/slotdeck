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
| Paragraph direction | Every paragraph gets the direction of its own text. In a deck that reads right to left, one right-to-left letter is enough, so "AI בשירות הלקוחות" reads right to left; in a deck that reads left to right, the first letter decides. A line with no right-to-left letter is left to right in either, so a Hebrew deck may hold an English line. A figure without letters ("+12%", "4.6 / 5") is left to right in a deck that reads left to right, where the template reads right to left; other text without letters keeps the template paragraph's direction. |
| Alignment | Left and right alignment are swapped (`alignOnDirectionChange: "flip"`, the default). A paragraph that states no alignment counts as left-aligned. Centred text stays centred. |
| Rows and grids of items | Filled mirror-wise within each line, so the deck's first item is where its readers start: at the right in a right-to-left deck. Shapes do not move; the values go to the mirrored positions. |
| Tables | Run in the deck's direction, with text aligned to its start. |
| Charts | Labels are set right to left. |
| Run language | Each run is marked with the language of its text: right-to-left text with a right-to-left language, and English (or a figure, in a deck that reads left to right) in a run the template marks right to left with a left-to-right one. See "Why runs are marked with a language" below. |
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

is displayed with the colon on the far side of the address. PowerPoint also checks spelling in the run's
language, so Hebrew words in a run marked as English are all underlined as misspelled.

In a run marked as Hebrew, English is reordered the other way round: "reached 74%, landing just shy of our
ambitious 75% target" is displayed as "reached landing just ,74% 75% shy of our ambitioustarget".

slotdeck therefore writes each run in the language of its text. You do not need to do anything for this.

- **Right-to-left text** gets a right-to-left language (`he-IL` for Hebrew, `ar-SA` for Arabic), unless the
  template already gave the run one. Text counts as right to left when it starts with a right-to-left letter or
  most of its letters are right to left, as in "AI בשירות הלקוחות". In a deck that reads right to left, one
  right-to-left letter is enough: "Microsoft Teams ככלי עבודה" is a Hebrew sentence that opens with a name.
- **Text with letters that is not right to left**, written into a run the template marks with a right-to-left
  language, gets a left-to-right one: the deck's language when it reads left to right (`en` gives `en-US`,
  `en-GB` stays `en-GB`), and `en-US` otherwise. The end of its paragraph gets it too, when it states a
  right-to-left language.
- **A figure without letters** ("+12%", "4.6 / 5") in a deck that reads left to right is treated the same way,
  and its paragraph reads left to right: a right-to-left paragraph shows "+12%" as "12%+" and "4.6 / 5" as
  "5 / 4.6". In a deck that reads right to left, a figure keeps what the template gave it.
- **Anything else** is left as the template has it: text with neither letters nor digits, and a run that states
  no language of its own but inherits one from the template's styles. Templates saved by PowerPoint state the
  language on every run.

Text is measured in the language and direction it is written in, so the checks and the renderer agree.

PowerPoint also saves the result of its spelling check on each run. A text that slotdeck writes does not keep the
template's result: PowerPoint checks it again when the deck is opened.

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
