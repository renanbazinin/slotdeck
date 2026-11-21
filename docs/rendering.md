# Rendering

Rendering turns a deck (JSON) into a `.pptx` using a template. It uses no model and no network, and it is
deterministic: the same template, deck and assets give the same bytes.

```ts
import { loadTemplate, renderDeck, validateDeck, extractAssets, DeckValidationError } from 'slotdeck';
```

## `loadTemplate(pptx, manifest)`

```ts
function loadTemplate(pptx: Uint8Array, manifest: TemplateManifest): Template;
```

Unzips the template once. The result is `{ pkg, manifest }`. Keep it and reuse it for every render; rendering
does not change it. Loading reads the whole file, so do it once at start-up rather than per request. A render
is fast: a 15-slide deck from the demo template takes about 100 ms.

It does not verify the manifest. Call `checkManifest(template)` when a template or manifest is loaded if you
want to catch a manifest that no longer matches its file (see [template-studio.md](template-studio.md)).

To get a picture of each slide of the result without PowerPoint, see [preview.md](preview.md).

A third argument, `{ fonts }`, gives the template the metrics of its fonts. Text is then measured as it will
be set: warnings say how many lines a text runs over, and text can be shrunk to fit. See
[text-fit.md](text-fit.md).

## `renderDeck(template, deck, options?)`

```ts
function renderDeck(template: Template, deck: DeckSpec, options?: RenderOptions): RenderResult;

interface RenderOptions {
  assets?: Record<string, Uint8Array>; // picture bytes by asset id
  layout?: 'strict' | 'adaptive';      // default 'strict'
  properties?: DeckProperties;         // from 0.7.0: what the file says of itself; see "File properties" below
}

interface RenderResult {
  pptx: Uint8Array;
  warnings: Issue[]; // non-blocking findings, such as text over a slot's budget
  shrunk: Array<{ slide: number; slot: string; size: number; scale: number }>; // see text-fit.md; empty unless the manifest has shrinkText
}
```

What it does, for each slide of the deck:

1. Clones the template slide of the slide's pattern, with its layout, background and every shape. For a free
   slide (see [free-slides.md](free-slides.md)) it clones the slide of the frame, takes that layout's own
   content off, and draws the slide's blocks as copies of the template's text boxes, panels and tables.
2. For a pattern with a row or grid of items, removes or copies items to match `items` and spaces them out.
3. Fills each slot: text into text shapes, pictures into picture frames, rows into tables, a native chart in
   place of a chart slot's sample picture.
4. Removes shapes whose slot is left out (optional slots in `adaptive` layout, and slots with `empty: "drop"`).
5. Sets the page number where the template has one.
6. Applies the colours the deck or the slide asks for, when it asks for any (see [colors.md](colors.md)).

Then it writes a new package that contains only what the new slides use: template slides, pictures and other
parts that are not referenced any more are left out, so the output is not bloated by the template's sample
content. Speaker notes, named sections and revision history of the template are not carried over, because they
describe the template's own slides.

### File properties

Outlook, Explorer and PowerPoint's File > Info show a file's properties without opening it. Left to itself a
rendered deck keeps its template's: the template's title, author and dates, "Slides: 33" for a deck of ten, and
the template's own picture of its sample cover as the file's thumbnail. With `properties` (from 0.7.0) the file
says what it is as a deck:

```ts
renderDeck(template, deck, { assets, properties: { author: 'Dana Levi', date: new Date() } });
```

| Field | Left out |
| --- | --- |
| `title` | The deck's `title`, else the title of its first slide that has one. |
| `language` | The deck's `language`, else the template's. |
| `author` | No one is named. |
| `date` | No date is written, so that the same deck always gives the same file. |
| `embedDeck` | Off. `true` keeps the deck's JSON in the file, without its `brief` and `document`; `readEmbeddedDeck(pptx)` takes it out again with its pictures, also after PowerPoint saved the file. |

The file also gets its real number of slides, and the version of slotdeck that made it as the custom property
`slotdeck` (File > Info > Properties > Advanced > Custom). The template's thumbnail, its sample slide titles and
its counts of words are taken out. Its other custom properties stay: a template whose files carry a label of how
confidential they are passes it on. The slides themselves are the same with the option or without it, and
without it the file is the same as before 0.7.0. `generateDeck`, `editDeck` and the chat tools set it unless
their own `properties` is `false`.

### Assets

Every `ImageRef` in the deck names an asset id. `options.assets` must have bytes under that id, or the deck is
invalid (`asset "x" was not provided`). The same bytes may be used by several slides; they are stored in the
file once.

```ts
renderDeck(template, deck, {
  assets: {
    clock: clockPngBytes,
    'photo-commute': photoJpegBytes,
    'gen-68db256a51fa6c83': generatedPhotoBytes,
  },
});
```

When a deck was made by the planner, `resolveAssets` or `generateDeck` builds this map for you
(see [generating.md](generating.md)).

### Errors

`renderDeck` validates first. If the deck has any error-level issue it throws:

```ts
try {
  renderDeck(template, deck, { assets });
} catch (error) {
  if (error instanceof DeckValidationError) {
    for (const issue of error.issues) console.log(issue.severity, issue.slide, issue.slot, issue.message);
  } else {
    throw error; // the manifest does not match the template, or an image format is not supported
  }
}
```

Other exceptions mean the manifest and the pptx disagree (for example
`Pattern "x": slot "y" points at shape #12, which is not on template slide 3`). `checkManifest` finds these
ahead of time.

## `validateDeck(template, deck, options?)`

```ts
function validateDeck(target: Template | TemplateManifest, deck: DeckSpec, options?: ValidateOptions): Issue[];

interface ValidateOptions {
  layout?: 'strict' | 'adaptive';
  assets?: Iterable<string>; // ids that will be available at render time; when given, image references are checked
  measure?: TextMeasurer;    // measures text with the template's fonts: see text-fit.md
}

interface Issue {
  severity: 'error' | 'warning';
  slide?: number; // 1-based; absent for deck-level issues
  slot?: string;
  message: string;
}
```

It uses exactly the same rules as the renderer, and takes either of two things:

- **The loaded template.** Everything is checked: a deck's own layouts are built in, its free slides are laid
  out and their text is checked against the room it gets, and with fonts the text is measured (no `measure`
  option is needed). Use this wherever the template is at hand.
- **The manifest alone.** For where the pptx is not available: in a browser, or before a job is queued. A deck
  of the template's own patterns is checked in full. A free slide is checked for what needs no template (its
  blocks, their required parts, the kinds of their values) and not for fit; a deck with layouts of its own
  needs `withLayouts`, as before.

- **Errors** block rendering: unknown pattern, a required slot without a value, a value of the wrong kind, a
  table whose rows do not match its header, a chart whose series do not match its categories, an item count
  outside the pattern's range, an asset that was not provided.
- **Warnings** do not: text, tables or charts over a slot's soft limits; a slot name the pattern does not have;
  `items` on a pattern with a fixed number of items; an unknown chart `format`; a colour that is no colour; a
  free slide so crowded that its text is very small or its charts are drawn bare.

All messages are listed in [troubleshooting.md](troubleshooting.md).

## `extractAssets(template, deck, pptx)`

```ts
function extractAssets(template: Template, deck: DeckSpec, pptx: Uint8Array): Record<string, Uint8Array>;
```

Reads the pictures of a deck back out of a pptx that was rendered from that deck, keyed by asset id. With it,
the deck JSON and its pptx are all you have to store between edits: generated photos are read back instead of
being generated, and paid for, again.

```ts
const cache = new Map(Object.entries(extractAssets(template, deck, previousPptx)));
const edited = await editDeck(template, deck, 'Shorten slide 3', { llm, images, cache });
```

Only slides that still match the deck are used. If the file was changed by hand (slides added, removed or
reordered, or a title retyped), those slides are skipped.

The pictures of free slides are found like any others. In a deck with colours (`colors` on the deck or on a
slide) the icons in the file were recoloured when it was rendered, so they are left out: supply the app's own
icons again, as for every render.

## Layout modes in practice

```ts
// The deck leaves out the optional "subtitle" slot.
renderDeck(template, deck, { layout: 'strict' });   // throws: leaving this slot empty needs layout: 'adaptive'
renderDeck(template, deck, { layout: 'adaptive' }); // renders; the subtitle shape is removed
```

Use `adaptive` unless you need a guarantee that every slide looks exactly like the template slide with
different words.

## What the output is

- A `.pptx` that PowerPoint opens without a repair prompt. Output was checked by rendering in PowerPoint on
  Windows; other programs were not tested.
- Charts are native charts with their data in an embedded workbook.
- Text keeps the template's formatting. Text boxes are not resized: text that is longer than its box overflows
  it, which is why budgets exist.
- The output has no dependency on slotdeck. It is an ordinary file.

## Serving a file from a web server

```ts
app.post('/decks', async (request, response) => {
  const deck = request.body as DeckSpec;
  const issues = validateDeck(template.manifest, deck, { layout: 'adaptive', assets: Object.keys(assets) });
  if (issues.some((issue) => issue.severity === 'error')) return response.status(400).json({ issues });
  const { pptx } = renderDeck(template, deck, { assets, layout: 'adaptive' });
  response
    .setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.presentationml.presentation')
    .setHeader('Content-Disposition', 'attachment; filename="deck.pptx"')
    .send(Buffer.from(pptx));
});
```
