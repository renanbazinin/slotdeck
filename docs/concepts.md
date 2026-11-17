# Concepts

## Template

An ordinary `.pptx` file: a deck of sample slides with filler text and sample pictures, as a designer hands it
over. slotdeck never draws anything. It copies slides of the template and replaces the content of some of their
shapes, so fonts, colours, backgrounds, logos and decorations are exactly the template's.

A loaded template is `{ pkg, manifest }` (type `Template`), made by `loadTemplate(pptxBytes, manifest)`. Loading
unzips the file once; keep the object and reuse it for every render.

## Pattern

One reusable slide layout. A pattern points at one slide of the template (`sourceSlide`, counted from 1) and has
an `id` that decks use to ask for it. Several patterns may point at the same slide (see "derived patterns" in
[layout-freedom.md](layout-freedom.md)).

A slide of the template that is not named by any pattern is simply never used.

## Slot

One shape on a pattern's slide whose content a deck replaces. A slot has a `name` (what decks call it), the id of
its shape, and a `kind`:

| Kind | The shape on the template | A deck gives it |
| --- | --- | --- |
| `text` | A text box or a shape with text | A string, or an array with one string per paragraph |
| `image` | A picture | `{ "asset": "<id>" }` naming picture bytes passed at render time |
| `table` | A table | `{ "header": [...], "rows": [[...], ...] }` |
| `chart` | A picture that stands where a chart goes | `{ "type", "categories", "series" }`, which becomes a native PowerPoint chart |

Shapes that are not slots are copied untouched.

## Manifest

The JSON that lists a template's patterns and slots, with limits for each slot (how many characters fit, how
many table columns). It is written once per template. Full reference: [manifest-format.md](manifest-format.md).

The manifest is the contract between the template and everything else:

- The renderer uses it to find shapes.
- The validator uses it to check a deck.
- The planner shows it to the language model, so the model can only choose patterns and slots that exist and
  is told how much text each slot holds.

## Deck

The JSON that describes one presentation: an ordered list of slides, each naming a pattern and giving values for
its slots. Full reference: [deck-format.md](deck-format.md).

A deck is the source of truth. It holds no binary data, so it is small (a few kilobytes), safe to send to a
model, and can be stored, compared, edited and rendered again at any time.

## Free slides

A slide need not come from a pattern. A **free slide** lists **blocks** (texts, cards, figures, charts,
tables, pictures, in any number) and the software lays them out itself, on the frame of one of the template's
layouts and in the template's own text styles. It is for content that no pattern holds. Its `pattern` is
`free`, and its content sits in `slots` under generated names (`b1Text`, `b2Chart`). See
[free-slides.md](free-slides.md).

## Colours

A deck has the template's colours unless it asks for others: `colors` on the deck for every slide, or on a
slide for that slide. The colour asked for takes the place of the template's main colour, and that colour's
tints and shades follow. See [colors.md](colors.md).

## Deck fields

Some text is the same on every slide, such as an organization or department line next to the logo. In the
manifest such a slot has `deckField: "<field name>"`, and a deck gives the value once:

```json
{ "fields": { "organizationDepartment": "Acme  |  Operations" } }
```

A slide may still set that slot itself, which overrides the field for that slide.

## Assets

Pictures are passed to the renderer as a map from **asset id** to bytes (PNG, JPEG or GIF). A deck refers to a
picture only by its id. There are three sources:

| Source | Who provides it | Typical use |
| --- | --- | --- |
| Icons (`IconDef`) | Your app registers a set; the model picks by description | Pictograms on cards |
| Supplied pictures (`AssetDef`) | Your app passes pictures it already has | A rendered diagram, a product photo, a chart image |
| Generated pictures | An `ImageProvider` you plug in | Photos described by the model |

A generated picture's reference keeps its description (`prompt`) in the deck, so the picture can be made again
later, and its id is derived from that description, so the same description always maps to the same id.

## Layout modes

| Mode | Meaning |
| --- | --- |
| `strict` (default) | Every slot must have a value unless its policy says otherwise. Nothing on a slide disappears. |
| `adaptive` | A slot marked optional (`"empty": "remove"`) may be left out, and its shape is then removed from the slide. |

## What happens to a slot without a value

Each slot has an `empty` policy in the manifest:

| `empty` | Without a value |
| --- | --- |
| `error` (default) | The deck is invalid: the slot is required. |
| `remove` | In `adaptive` layout the shape is removed. In `strict` layout the slot is required. |
| `keep` | The template's own content stays. Use it for decoration that suits any content. |
| `drop` | The shape is always removed, in any layout. Use it for decoration drawn for the sample content, such as a sticker on a sample photo. |

A deck can also give `null` for a slot to leave it out explicitly; the shape is removed (allowed where the
policy is `remove`, in `adaptive` layout).

## Character budgets

A text slot has `maxChars` (and optionally `maxParagraphs`). These are soft limits: text over the budget is
reported as a **warning**, not an error, and is still rendered. The planner shows budgets to the model and sends
a draft that exceeds them back for repair.

A budget is an estimate of what fits the box. It cannot see line wrapping, so slots with several paragraphs
should have budgets with room to spare.

## Rows and grids of items

A pattern can declare that some of its shapes form a row or a grid of like items (cards, figures, steps) and
that a deck may choose how many there are. See [layout-freedom.md](layout-freedom.md).

## Reading direction

A template has a direction (`ltr` or `rtl`), and so has a deck. When they differ, text direction and alignment
flip while shapes stay where they are. See [languages.md](languages.md).

## Models are plugged in

slotdeck defines two small interfaces and calls nothing else:

- `LlmProvider.generateJson(request)` answers a prompt with JSON.
- `ImageProvider.generateImage(request)` returns picture bytes.

Adapters for Gemini are included. Anything else is a few lines: see [providers.md](providers.md).
