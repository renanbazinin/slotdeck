# Slide previews

A preview is a picture of each slide of a pptx, for showing a deck in an app without opening PowerPoint.

slotdeck draws nothing itself. It has a small interface, `PreviewProvider`, that you fill with a renderer of
your choice, and it calls that interface wherever a picture is wanted: for a deck you rendered, for the
thumbnails of a template's patterns, and for every presentation the chat tools create.

An adapter for one renderer, the Pagus library, comes with the package. Pagus is not a dependency of slotdeck:
you install it and pass it in.

**Read "How close the pictures are" below before relying on a preview.** No renderer that runs without
PowerPoint was found to reproduce PowerPoint exactly. A preview shows what a deck looks like. To know whether
a text fits its box, measure it: see [text-fit.md](text-fit.md).

## Setup with Pagus

```
npm install @pagus-kit/core @pagus-kit/renderer
```

```ts
import { readFileSync, writeFileSync } from 'node:fs';
import * as pagusCore from '@pagus-kit/core';
import * as pagusRenderer from '@pagus-kit/renderer';
import { loadTemplate, pagusPreviews, previewDeck } from 'slotdeck';

const preview = pagusPreviews(pagusCore, pagusRenderer);

const template = loadTemplate(readFileSync('templates/demo.pptx'), manifest);
const { pptx, previews } = await previewDeck(template, deck, preview, { assets });

for (const picture of previews) {
  writeFileSync(`out/slide-${picture.slide}.svg`, picture.data); // picture.mimeType is 'image/svg+xml'
}
```

`previewDeck` is `renderDeck` followed by the provider. To draw a pptx you already have, call the provider
directly:

```ts
const previews = await preview.renderSlides({ pptx });               // every slide
const some = await preview.renderSlides({ pptx, slides: [1, 4] });   // 1-based slide numbers
```

A 15-slide deck takes about 0.2 seconds.

### What a Pagus preview is

- One SVG for each slide, 1280 pixels wide for a 13.33 inch slide (96 pixels to the inch). Shapes, pictures,
  tables and native charts are drawn. Pictures are embedded, so each SVG stands on its own.
- **The text is real text, set by the browser that shows the SVG.** The browser uses its own fonts and breaks
  the lines itself. The picture therefore needs a browser to be seen, and it looks right only where the
  template's fonts are installed, or where you supply substitutes (see the options).
- It can be shown with `<img src="...">`, with a `data:` address, or by putting the SVG markup into the page.
  This was checked in Chromium. Do not put the markup of several slides into one page: their internal names
  collide. Use one `<img>` for each.
- It cannot be turned into a PNG by an SVG library that has no HTML engine (the text is HTML inside the SVG).
  For PNG, take a screenshot of the SVG with a headless browser, or use another provider.

### Options

```ts
pagusPreviews(pagusCore, pagusRenderer, {
  scale: 0.5,                                  // 640 pixels wide
  fontSubstitutes: { Calibri: 'Carlito', Cambria: 'Caladea' },
  fontCss: "@import url('/fonts/carlito.css');",
});
```

| Option | Default | Meaning |
| --- | --- | --- |
| `scale` | `1` | Scale of the drawing. `1` gives 96 pixels to the inch. |
| `fontSubstitutes` | none | For each of the template's typefaces, a typeface to fall back on where the viewer does not have it. |
| `fontCss` | none | `@font-face` or `@import` rules to embed in each SVG, so that it brings its own fonts. A browser does not fetch fonts by address for an SVG shown through `<img>`: embed them as data in the rule, or put the markup into the page. |

### Right-to-left text

Pagus leaves the direction of a paragraph to the browser, which takes a paragraph to run left to right unless
told otherwise. The adapter therefore marks every paragraph that starts with a right-to-left letter as running
right to left, the same rule slotdeck's renderer uses when it writes the pptx. Without this, a Hebrew sentence
gets its full stop on the wrong side and a line that mixes Hebrew with an address or a number comes out in the
wrong order.

## Thumbnails of a template's patterns

```ts
import { patternPreviews, previewExtension } from 'slotdeck';

const pictures = await patternPreviews(template, preview);
// { cover: SlidePreview, 'text-and-photo': SlidePreview, ... }  one for each pattern id

for (const [id, picture] of Object.entries(pictures)) {
  writeFileSync(`thumbnails/${id}.${previewExtension(picture)}`, picture.data);
}
```

Each picture is the pattern's slide of the template's sample deck (`sampleDeck`), so a pattern derived from
another pattern's template slide is shown as it will look, with sample words in its slots. The options of
`sampleDeck` (`fill`, `items`, `fit`) can be passed as a third argument. Make the thumbnails once, when a
template or its manifest changes, and serve them as files: they are what the `thumbnail` option of the chat
tools points at (see [chat-tools.md](chat-tools.md)).

The same call gives the studio a picture of a manifest's sample deck, to judge budgets by eye:

```ts
const { deck, assets } = sampleDeck(template, { items: 'max' });
const { previews } = await previewDeck(template, deck, preview, { assets, layout: 'adaptive' });
```

## Previews from the chat tools

Give `presentationTools` a provider, and every presentation that `create_presentation` or `edit_presentation`
makes is drawn:

```ts
const tools = presentationTools({ template, llm, store, preview });
```

- The pictures are kept in the store under `preview/<presentation id>/slide-01.svg`, `slide-02.svg`, and so on
  (`storeKeys.preview(id, slide, extension)`; the extension follows the provider's picture type).
- The tool's `output` for the chat model lists the keys in `previews`. The pictures themselves are not given to
  the model.
- The `ToolResult` for your app carries them in `previews`, next to `deck` and `pptx`.
- If the provider fails, the presentation is still created and stored. `output.preview_error` then holds the
  message, and there is no `previews`.

## Writing your own provider

```ts
import type { PreviewProvider } from 'slotdeck';

const preview: PreviewProvider = {
  async renderSlides({ pptx, slides, signal }) {
    // Turn the pptx into pictures in any way: a service that runs PowerPoint or LibreOffice,
    // a headless browser, another library.
    const pngs: Uint8Array[] = await myRenderer.toPng(pptx, { signal });
    return pngs
      .map((data, index) => ({ slide: index + 1, mimeType: 'image/png', data }))
      .filter((picture) => !slides || slides.includes(picture.slide));
  },
};
```

The contract:

| Field | Meaning |
| --- | --- |
| `request.pptx` | The presentation's bytes. |
| `request.slides` | 1-based numbers of the slides wanted; every slide when left out. |
| `request.signal` | Stop when it is aborted. |
| result | One `SlidePreview` for each slide asked for, in slide order: `{ slide, mimeType, data, width?, height? }`. `mimeType` is `image/svg+xml` or a bitmap type such as `image/png`; `data` is the SVG as UTF-8 text or the bitmap's bytes. |

A service that runs PowerPoint itself, or LibreOffice, is the way to pictures that match what people will see
in PowerPoint, and to PNG. A PNG preview also fits `draftManifest`'s `slideImage` option, which takes the same
`{ mimeType, data }` shape.

## How close the pictures are

Three renderers were tried on this project's demo template and its two example decks (`examples/decks`, in
English and in Hebrew: 22 slides with text, cards, a table, picture charts and a native chart), and compared
with PowerPoint 16's own export of the same files. For the two that produce SVG with real text, the SVG was
drawn by Chromium on a machine that has the template's fonts.

| Renderer | Needs a browser | English | Hebrew (right to left) | Tables | Native charts |
| --- | --- | --- | --- | --- | --- |
| Pagus 0.1.1 (`@pagus-kit/renderer`), through the adapter | yes | Close: same typefaces and weights, same line breaks in running text | Readable, in the right order; punctuation and mixed lines right with the adapter's marking | Styling close; a right-to-left table is drawn left to right | Close: bars, values, legend, gridlines; the axis steps differ |
| pptx-svg 0.6.8 | yes | Positions right; a title fell back to another typeface; line breaks depend on a width function you supply | Readable; full stops on the wrong side and mixed lines in the wrong order | Grid lines drawn that the template does not have; not mirrored | Drawn with another layout (legend at the top right, a frame) |
| pptx-glimpse 5.3.1 | no (draws text as outlines, makes PNG) | Shapes and pictures right; bold and italic are lost, so lines break elsewhere | **Unreadable: the letters of every word come out in reverse order** | Styling lost; not mirrored | Bars, axis and legend; no values on the bars, no gridlines |

Measured as the mean difference between the two pictures of a slide (greyscale, 760 pixels wide), Pagus was
closest: 1.2% for the English deck and 1.3% for the Hebrew one, against 1.6% and 1.8% for pptx-glimpse and
2.0% and 1.7% for pptx-svg. The figure says little about text: the Hebrew slides of pptx-glimpse score well
and cannot be read.

What is wrong in a Pagus preview, as seen on these decks:

- **Text anchored to the bottom of its box is drawn from the top.** A cover title that sits low on the slide
  in PowerPoint sits higher in the preview. Text anchored to the middle is placed right.
- **A line that just fits in PowerPoint may wrap in the browser**, and the reverse. The browser measures text
  a little differently. In running text the breaks were the same on the slides checked; a title that fills
  its line to the last point wrapped.
- **A right-to-left table is not mirrored:** its first column is on the left.
- **Rounded corners of a picture are not drawn**, and a bullet sits closer to its text.
- **A native chart's axis** has other steps, and the bars stand a little differently.
- **Fonts.** The viewer's browser needs the template's fonts. Without them it falls back, and every line can
  break elsewhere.

Other findings worth knowing:

- pptx-glimpse 5.3.1 fails outright when it is given Calibri or Cambria as font files
  (`lookupType: 7 - substFormat: 1 is not yet supported`). Those are the fonts it finds first on a Windows
  machine, so there it fails with its default settings. It runs with open fonts such as Carlito and Caladea,
  and then uses one face for each family: the bold and italic files are not told apart.
- pptx-glimpse draws a shadow under shapes for which the template turns shadows off.
- pptx-svg in Node needs a function that gives the width of a text. It is told the typeface and the size, not
  the weight. Given real widths from the font files (see [text-fit.md](text-fit.md)), its lines for regular
  text were the ones PowerPoint made.
- pptx-browser 4.1.5 draws on a canvas. In Node that needs the `canvas` package, which is native code, so it
  was not tried.

In short: Pagus in a browser is good enough to show a person what a deck looks like, in English and in Hebrew.
None of the three can stand in for PowerPoint when the question is whether a slide is right to the line. That
question is answered by measurement ([text-fit.md](text-fit.md)), which agrees with PowerPoint line for line.

### Checking a renderer yourself

The repository has what the comparison was made with. It needs Windows with PowerPoint and is not part of the
package.

```
node examples/render-demo.ts                                    # writes out/patterns-en.pptx and out/patterns-he.pptx
powershell -File scripts/render-pptx.ps1 -Path out/patterns-he.pptx -OutDir out/ppt-he -Width 1280
```

`render-pptx.ps1` exports PowerPoint's own picture of every slide. Draw the same file with your provider and
put the two side by side.
