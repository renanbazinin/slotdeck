# Colours

A deck normally has exactly the colours of its template. From 0.6.0 a deck may ask for other colours, for the
whole deck or for single slides: "make it blue", "slide 3 in red". The template file is never changed; the
colours are applied to the copy that `renderDeck` writes.

```ts
import { renderDeck, validateDeck, recolorDeck, templatePalette } from 'slotdeck';
import type { DeckColors, TemplatePalette } from 'slotdeck';
```

## In one minute

```json
{
  "version": 1,
  "colors": { "accent": "1565C0" },
  "slides": [
    { "pattern": "cover", "slots": { "title": "A year of rides" } },
    { "pattern": "stat-cards", "colors": { "accent": "B3261E" }, "slots": { "title": "Where we lost ground" } }
  ]
}
```

The JSON shows only where `colors` goes; the slides' other slots are left out.

- `colors` on the deck: every slide takes the new colour in place of the template's main colour.
- `colors` on a slide: that slide alone, in place of the deck's.
- A deck without `colors` renders byte for byte as it did before 0.6.0.

## `DeckColors`

| Field | Type | Meaning |
| --- | --- | --- |
| `accent` | six hex digits | Takes the place of the template's main colour. The tints and shades of that colour follow it: a pale teal panel becomes a pale blue panel. |
| `accent2` | six hex digits | The same for the template's second colour. |
| `replace` | object of hex to hex | Exact replacements of single colours, such as `{ "0E7C7B": "1565C0" }`. For a colour the two accents do not reach. |

A colour is written as six hex digits, such as `1565C0`. A leading `#` and lower case are read too. A value
that is no colour (a name such as `"blue"`, three digits) is left out, and `validateDeck` returns a warning
that names it; the deck still renders. The chat tools are more forgiving with what a chat model sends: see
`accent_color` in [chat-tools.md](chat-tools.md).

## What follows the new colour, and what stays

| Follows | Stays as the template has it |
| --- | --- |
| Shapes, lines and text set in the main colour, or in a tint or shade of it | Black, white and greys |
| The same colours in the slide layouts, the masters and the theme (deck colours) | Colours of other families: an amber highlight stays amber when the teal becomes blue |
| Chart series, table header fills and table text in those colours | Photos, and pictures that are not icons |
| Icons: PNG pictures placed in an icon slot are recoloured pixel by pixel | The logo and every other picture that is part of the template's frame |
| Gradient stops in those colours | Text colours that would become unreadable (see below) |

How a family follows: the template's colours are sorted into families by hue. The main colour of a family is
replaced by the colour asked for; each other member keeps its distance from the main colour in lightness and
strength, so a tint stays a tint and a shade stays a shade. Each colour of the template is mapped once, to one
new colour, everywhere it appears.

**Text stays readable.** A light colour asked for where the template had a dark one can leave white text on
a pale fill, or coloured text too pale for its background. Where the new colours bring a text's contrast to its
background under 3 to 1, and the template had it above that, the text is set in a darker or lighter shade of
its own colour that reads, or in black or white when no shade does. Text the template itself set at low
contrast is left as it is.

## Which colour is "the main colour"

`templatePalette(template)` returns what the software found:

```ts
const palette = templatePalette(template);
// For the demo template:
// {
//   accent: '0E7C7B',              // what "accent" in a deck's colours takes the place of
//   accent2: 'F4A259',             // what "accent2" takes the place of
//   ink: '1B2A41',                 // the dark colour text is set in, when it is not a grey
//   families: [                    // the most used first
//     { base: '0E7C7B', members: ['E8F1F0', '0E7C7B'], share: 0.5186 },
//     { base: '1B2A41', members: ['1B2A41'], share: 0.3056 },
//     { base: 'F4A259', members: ['F4A259', 'FCE9D2'], share: 0.0316 },
//   ],
//   colors: [                      // every colour, the most used first
//     { hex: 'E8F1F0', share: 0.3575 },
//     { hex: '1B2A41', share: 0.3056 },
//     { hex: '0E7C7B', share: 0.1611 },
//     ...
//   ],
// }
```

A family's `members` are the most used first, so its `base`, the strong colour it is named by, need not come
first: in the demo template the pale tint `E8F1F0` covers more of the slides than the teal itself. For the same
reason the first of `colors` need not be the main colour.

The main colour is the strong colour that the template uses most, counted over the slides its patterns stand
on. A background and a large filled shape count more than a thin line, and a long text more than a word. Where
the software's choice is not the brand's, the manifest says so:

```json
{ "colors": { "accent": "0E7C7B", "accent2": "F4A259" } }
```

and a template that must never change colour says:

```json
{ "colors": false }
```

A deck's `colors` are then left out when it is rendered, and `validateDeck` returns the warning `the template
keeps its own colours in every deck; the colours the deck asks for are left out`. A studio sets both with
`setManifest` (see [template-studio.md](template-studio.md)).

## Deck colours and slide colours

They are applied differently, and the difference shows in PowerPoint.

- **Deck colours** change the theme, the masters and the layouts in the output file as well as the slides.
  A person who adds a slide in PowerPoint afterwards gets the new colours, and the colour picker shows them.
- **Slide colours** change that slide only. The colours the slide names, as they are or through the theme, are
  changed on the slide itself (a theme colour that changes becomes a plain colour there), and a background
  the slide inherits is copied to it, so the other slides keep theirs. Text that states no colour is not
  reached: see "Limits".

A slide with its own `colors` does not take the deck's as well: its colours are worked out from the template's,
not from the deck's. A slide that states `"colors": {}` in a deck with colours keeps the template's own, except
for text that states no colour, which follows the deck.

## Setting colours from code

Colours are plain fields, so they can be set directly. `recolorDeck` does it with every colour stored in its
kept form (six hex digits, upper case, no `#`), and returns a new deck. A value that is no colour is kept as
given, and `validateDeck` names it:

```ts
const blue = recolorDeck(deck, { accent: '1565C0' });            // the whole deck
const mixed = recolorDeck(blue, { accent: 'B3261E' }, [3, 4]);   // slides 3 and 4 only, as their own
const again = recolorDeck(mixed, null, [3, 4]);                  // slides 3 and 4 follow the deck again
const plain = recolorDeck(again, null);                          // the deck back to the template's colours
```

Setting the deck's colours leaves a slide's own colours alone, and removing the deck's does too: name the
slides to remove theirs. No model is called and no text changes. Render the result with `renderDeck` as usual.

## Colours when a deck is planned or edited

No model reads a colour out of a brief. A colour reaches a deck only where it is stated outright:

| Where | How |
| --- | --- |
| A new deck | `Brief.colors`, copied to `DeckSpec.colors` when the deck is written |
| A plan that is changed | `colors` in the options of `editPlan` |
| A deck that is changed | `colors` and `colorSlides` in the options of `planEdit` and `editDeck` |
| The chat tools | the `accent_color` and `color_slides` arguments, which the chat model fills when the person asks for a colour |

This is deliberate. A phrase such as "Blue Ocean strategy" in a brief must not turn a brand template blue. In
a chat, the chat model hears the wish and passes the colour as an argument; see [chat-tools.md](chat-tools.md).

## Pictures taken back out of a file

`extractAssets` returns the pictures of a rendered deck by asset id. In a deck with colours, icons were
recoloured when it was rendered, so the icons in the file are no longer the ones supplied. `extractAssets`
leaves the icons of such a deck out: supply the original icons again when the deck is rendered again.

## Limits

- Icons are recoloured when they are PNG files of 8 bits per channel, with a palette or in RGB or RGBA, not
  interlaced, and no larger than four million pixels. Other icons are placed as supplied.
- Pictures inside the template (a photo used as a background, a drawn illustration) keep their colours.
- A gradient follows when its stops are colours of a family; the kind of gradient does not change.
- Colours are not checked for accessibility beyond text on its own background: two neighbouring fills may be
  close in lightness when the colour asked for is close to another of the template's.
- Charts supplied as pictures keep their colours. Native charts (a slot of kind `chart`) follow.
- Text that states no colour of its own takes it from the master. A slide's own colours do not reach such
  text, and on a slide with `"colors": {}` it follows the deck's colours. This shows only where the colours
  change what the master gives text: a `replace` of black, or `accent` on a template whose master sets text in
  its main colour.
