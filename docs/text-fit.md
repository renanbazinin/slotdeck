# Measuring text with fonts

Without fonts, slotdeck judges whether a text fits its slot by counting characters against a budget
(`maxChars`). A budget cannot see where lines break, and letters differ in width from one typeface and one
language to another, so it is a rough guard.

Given the template's font files, slotdeck measures a text the way PowerPoint sets it: it breaks the text into
lines at the slot's typeface and size and compares them with the room the slot's box gives. Validation then
says "takes 7 lines; the box holds 5", the planner repairs by that, and text can be set a little smaller to fit.

This is optional. A template loaded without fonts behaves exactly as before.

## Setup

Measuring needs a library that reads font files. slotdeck has an adapter for `fontkit` and does not depend on
it: install it yourself and pass the module in.

```
npm install fontkit
```

```ts
import { readFileSync } from 'node:fs';
import * as fontkit from 'fontkit';
import { fontkitMetrics, loadTemplate, renderDeck } from 'slotdeck';

const fonts = fontkitMetrics(fontkit, {
  // typeface name as the template states it: one file, or several for its weights and slants
  Calibri: [readFileSync('fonts/calibri.ttf'), readFileSync('fonts/calibrib.ttf')],
  Cambria: [readFileSync('fonts/cambria.ttc'), readFileSync('fonts/cambriab.ttf')],
});

const template = loadTemplate(readFileSync('templates/demo.pptx'), manifest, { fonts });
const { pptx, warnings } = renderDeck(template, deck, { assets });
```

That is all: everything that takes the template now measures text. Nothing else in your code changes.

`fontkit` ships no TypeScript declarations. In a TypeScript project add a file such as `fontkit.d.ts` with one
line: `declare module 'fontkit';`.

### Which fonts a template needs

`templateFonts(template)` lists the typefaces the template's text slots are set in, with the weights and
slants used:

```ts
import { loadTemplate, templateFonts } from 'slotdeck';

templateFonts(loadTemplate(pptx, manifest));
// [ { typeface: 'Calibri', bold: false, italic: false },
//   { typeface: 'Calibri', bold: true,  italic: false },
//   { typeface: 'Cambria', bold: true,  italic: false }, ... ]
```

Pass a language as the second argument (`templateFonts(template, 'he')`) when decks are written in a language
other than the template's: Hebrew or Arabic text is often set in another typeface than Latin text (see
"Which typeface a character is set in" below), and that typeface is then listed too.

On a machine that has the fonts installed, `fontFiles` from `slotdeck/node` finds the files in a folder by the
family name each file states:

```ts
import * as fontkit from 'fontkit';
import { fontkitMetrics, loadTemplate, templateFonts } from 'slotdeck';
import { fontFiles } from 'slotdeck/node';

const names = templateFonts(loadTemplate(pptx, manifest)).map((font) => font.typeface);
const files = fontFiles(fontkit, 'C:/Windows/Fonts', names); // or '/Library/Fonts', '/usr/share/fonts/truetype/...'
const template = loadTemplate(pptx, manifest, { fonts: fontkitMetrics(fontkit, files) });
```

`fontFiles` reads one folder and does not look into subfolders. Scanning a system font folder takes about a
second, so do it once at startup.

### What `fontkitMetrics` accepts

- TrueType (`.ttf`), OpenType (`.otf`), collections (`.ttc`), WOFF and WOFF2, as `Uint8Array` or `Buffer`.
- A key is the typeface name **as the template states it**, matched without regard to case. The value is one
  file or an array of files. Regular, bold, italic and bold italic are told apart by what each file says of
  itself, so the order does not matter.
- From a collection, the font whose family name equals the key is taken; when none does, the first.
- A weight or slant that was not supplied is measured with the nearest face that was (a bold from a regular
  file is made slightly wider, as a synthesized bold is). A variable font supplies its own bold. Results
  measured this way are marked `approximate`.
- A character the typeface has no glyph for is measured in the first supplied typeface that has one, and the
  result is marked `approximate`.
- A typeface that was **not supplied at all** is not guessed: slots set in it are not measured and keep their
  character budget. `checkManifest` and `textMeasurer(template).missingFonts()` name such typefaces.

### Fonts you may not copy

Calibri, Cambria, Arial and most fonts that come with Windows or Office are licensed with that software.
Check your licence before copying their files to a server.

A file may stand in for a typeface of another name. Some open fonts were drawn to the exact widths of a
licensed one, so text measures the same:

| Template typeface | Open font with the same widths | Note |
| --- | --- | --- |
| Calibri | Carlito | Checked against PowerPoint for English text: the same line counts at 12 pt and up, nearly always below. Carlito has no Hebrew or Arabic letters. |
| Cambria | Caladea | Checked together with Carlito. Latin only. |
| Arial, Times New Roman, Courier New | Liberation Sans, Liberation Serif, Liberation Mono | Drawn for this purpose; not checked here. |

```ts
const fonts = fontkitMetrics(fontkit, {
  Calibri: [readFileSync('fonts/Carlito-Regular.ttf'), readFileSync('fonts/Carlito-Bold.ttf')],
  Cambria: [readFileSync('fonts/Caladea-Regular.ttf'), readFileSync('fonts/Caladea-Bold.ttf')],
});
```

Any other substitute gives lines that are close but not exact. For a Hebrew or Arabic template, supply the
template's own font files: line counts depend on the real letter widths.

## What changes once the template has fonts

| Where | Without fonts | With fonts |
| --- | --- | --- |
| `validateDeck` (pass `measure`), `renderDeck` warnings | `text is 410 characters; the slot fits about 352` | `text takes 7 lines at 18 pt; the box holds 5 (2 lines too many). Text of about 290 characters fits.` |
| Text longer than its budget that does fit | warning | no warning: the measurement decides |
| Text within its budget that does not fit | no warning | warning |
| `planDeck`, `writeDeck`, `generateDeck`, `planEdit`, `editDeck`, the chat tools | a draft over a character budget goes back to the model | a draft that runs over its box goes back to the model, with the lines it takes and the length that fits |
| `renderDeck` with `shrinkText` in the manifest | nothing | text that does not fit is set smaller, within the limit you state |
| `checkManifest` | checks that slots point at real shapes | also tries every budget out and reports one that lets text run over |
| `sampleDeck(template, { fit: 'box' })` | same as the default | fills every text slot to what its box really holds |

What does **not** change: the budgets the planning model is told (`maxChars`), the paragraph limit
(`maxParagraphs`, still reported as before), and tables and charts, which are not measured. To give the model
budgets that match the boxes, see "Budgets from measurement" below.

`validateDeck` takes a manifest, not a template, so it gets the measurer as an option:

```ts
import { textMeasurer, validateDeck } from 'slotdeck';

const issues = validateDeck(template.manifest, deck, { layout: 'adaptive', measure: textMeasurer(template) });
```

`renderDeck`, the planner and the studio functions take the template and need nothing extra.

### The warnings

```
text takes 7 lines at 18 pt; the box holds 5 (2 lines too many). Text of about 290 characters fits.
```

- `7 lines` is what the text takes at the size it will be rendered at, over all its paragraphs.
- `the box holds 5` is how many lines of this text the slot's box has room for.
- `Text of about 290 characters fits` is the length to cut to. It is worked out from this text, so it is in
  the text's own language and letter widths. The planner sends this sentence to the model as it is.
- `about 7 lines` (with "about") means a typeface or weight was stood in for, so the count is close, not exact.
- `even at 15.5 pt` means the text was already set as small as `shrinkText` allows and still does not fit.

For a box that does not wrap its text, the warning reads
`text is 30% wider than its box at 12 pt, and the box does not wrap. Text of about 40 characters fits.`

## Shrinking text to fit

Set `shrinkText` on the manifest to let text that does not fit be set smaller:

```json
{
  "version": 1,
  "name": "demo",
  "direction": "ltr",
  "shrinkText": { "to": 0.85, "minSize": 10 },
  "patterns": []
}
```

| Property | Type | Meaning |
| --- | --- | --- |
| `to` | number | The smallest share of its size a text may be set at. `0.85` lets 20 pt text go down to 17 pt. |
| `minSize` | number, optional | A size in points under which no text is set, whatever `to` allows. Use it to keep small body text readable. |

How it works:

- Only text that does not fit is touched, and only as far as needed: the size goes down in half points until
  the text fits, so one long title becomes 32 pt and the others stay 34 pt. A shrunk text must fit with a
  quarter of a point of width to spare, so that it does not sit exactly on the edge.
- The code writes the new size into the slide. It does not rely on PowerPoint's own "shrink text on overflow",
  so the file looks the same in every viewer.
- A text that fits once it is shrunk produces no warning. One that does not fit even at the limit is set at the
  limit and reported with `even at`.
- `renderDeck` returns what it shrank:

```ts
const { pptx, warnings, shrunk } = renderDeck(template, deck, { assets });
// shrunk: [ { slide: 3, slot: 'title', size: 32, scale: 0.941 } ]
```

- The planner knows the setting too: a draft that fits after shrinking is not sent back to the model.
- Without fonts the setting does nothing, because nothing can be measured.

`shrinkText` works together with `textScale`: body text is first enlarged by the template's and the deck's
scale, and a text that then does not fit is reduced from that size.

## The studio: checking and setting budgets

### `checkManifest` tries budgets out

For a template with fonts, `checkManifest(template)` fills each text slot with its own sample words up to the
slot's budget and measures the result. A budget that lets text run over is reported as a warning:

```
{ severity: 'warning', pattern: 'text-only', slot: 'body',
  message: 'has a budget of 900 characters, but sample text of that length takes 14 lines where 9 fit; about 560 characters fit' }
```

Typefaces that were not supplied are reported once, without a pattern:

```
{ severity: 'warning',
  message: 'no font was supplied for Cambria bold; text set in it is checked by its character budget' }
```

The check is made at the designed size: `shrinkText` is not taken into account, since a budget that only fits
when shrunk would shrink every full slide.

### Budgets from measurement

`measuredBudgets(template)` measures how many characters each text slot's box holds, and suggests a budget a
little under that:

```ts
import { changeManifest, measuredBudgets } from 'slotdeck';

const budgets = measuredBudgets(template);
// [ { pattern: 'cover', slot: 'title', maxChars: 36, fits: 52, suggested: 46 }, ... ]

const { manifest } = changeManifest(
  template,
  budgets.map((budget) => ({ op: 'setSlot', pattern: budget.pattern, slot: budget.slot, maxChars: budget.suggested })),
);
```

| Field | Meaning |
| --- | --- |
| `pattern`, `slot` | Which slot. |
| `maxChars` | The budget the manifest states now, when it states one. |
| `fits` | The characters of running text the box holds, as measured. |
| `suggested` | `fits` less a margin: a budget to set. |

Options:

| Option | Default | Meaning |
| --- | --- | --- |
| `sample` | each slot's own sample words | Running text to measure with. Give a paragraph in the language your decks are written in when that is not the template's language: Hebrew letters are not as wide as Latin ones, so the same box holds a different number of them. |
| `margin` | `0.9` | The share of `fits` to suggest. Lines break between words, so two texts of the same length do not always take the same number of lines. |

A slot with several paragraphs is measured with as many paragraphs as it allows (`maxParagraphs`), which is
the case that leaves least room. Slots set in a typeface that was not supplied are left out. Budgets drafted
without fonts are deliberately cautious, so measured ones are usually larger: on the demo template by about a
half.

The budget is what the planning model is told, so setting measured budgets is what lets the model use the room
there is. Validation by measurement works with any budgets.

### A sample deck filled to the boxes

```ts
const { deck, assets } = sampleDeck(template, { fit: 'box' });
```

With `fit: 'box'`, each text slot is filled with as much sample text as its box holds by measurement, whatever
its budget says. Rendering it shows every box at its true capacity. `fill` still applies (0.95 by default) as a
share of what the box holds. A slot that cannot be measured is filled to its budget, as is every slot when the
template has no fonts.

## Measuring a text yourself

```ts
import { textMeasurer } from 'slotdeck';

const measure = textMeasurer(template);            // or textMeasurer(template, fonts)
const fit = measure.fit('text-and-photo', 'body', ['First paragraph.', 'Second paragraph.']);
```

`fit(pattern, slot, text, options?)` takes a pattern id and a slot name (or the `PatternDef` and `TextSlotDef`
themselves), and the text as a string or one string per paragraph. It returns `undefined` when the slot cannot
be measured: it is not a text slot, its shape holds no text, or a typeface it needs was not supplied.

| Field of the result | Meaning |
| --- | --- |
| `fits` | Whether the text stays inside its box. |
| `lines` | Lines the text takes, over all its paragraphs. |
| `capacity` | Lines of this text the box holds. |
| `overflow` | Lines that do not fit; `0` when the text fits. |
| `height`, `room` | Height the text takes, and the height the box gives it, in points. |
| `widthRatio` | For a box that does not wrap: the longest line's width divided by the box's. `1` or less fits. Always `1` for a box that wraps. |
| `size` | Font size of the first paragraph as measured, in points. |
| `sizes` | The size of each paragraph, in points. |
| `scale` | The factor the text was reduced by to make it fit; `1` when it was not. |
| `fitsChars` | About how many characters of text like this the box holds. |
| `paragraphs` | The lines of each paragraph, as they break. |
| `widths` | The width of each of those lines in points, without the space that hangs at its end. |
| `lineRoom` | The width a line of each paragraph may take, in points. |
| `sampleLines` | Lines the template's own sample text takes in the slot as drawn; `0` for a slot that resizes its box or resets its size. |
| `approximate` | A typeface, a weight or a shape outline was stood in for, so the figures are close but not exact. |

Options:

| Option | Meaning |
| --- | --- |
| `textScale` | The factor on body text the deck is rendered with: the manifest's `textScale` times the deck's. Defaults to `1`. |
| `shrink` | `{ to, minSize? }`: allow the text to be set smaller, as `shrinkText` does. |
| `items` | For a slot of a row of items: how many items the slide has. In a row with `widen`, fewer items are wider, and the text is measured in the wider box. |

`measure.missingFonts()` returns the typefaces the template's text slots need that were not supplied.

## How a text is measured

The rules below were read off PowerPoint's own behaviour.

**The box.** The room is the shape's size less its text insets (0.1 inch left and right, 0.05 inch top and
bottom unless the shape says otherwise). A rounded rectangle, an ellipse or a diamond keeps text clear of its
outline, as PowerPoint does. A box that was drawn around its sample text holds at least what the sample takes.
A slot with its own `box` or `fontSize` (a derived pattern) is measured in that box at that size. A `box` of
the shape's own size, such as a mirrored pattern's, only moves the shape, and the sample still counts.

**Rows that widen.** A row with `widen` makes its items wider when a slide has fewer of them than the template
draws (see [layout-freedom.md](layout-freedom.md)). With `items` in the options, the text of such an item is
measured in the box the renderer gives it: the same step that lays the row out for a deck is run on a copy of
the slide, and the width is read from there. Only the width changes. The height of the box, and how full the
designer drew it (`sampleLines`), are those of the item as drawn. `validateDeck`, `renderDeck`, shrinking and
`checkDesign` pass the slide's `items` themselves. On the demo template, 22 texts in widened cards and figures
broke every line where PowerPoint did, in English and in Hebrew; measured in the boxes as drawn, 3 of the 22
English ones would have had PowerPoint's line count.

**The size.** The size the text will be rendered at: the template run's size or the slot's `fontSize`, times
the body text scale where it applies (`scales`), rounded to half a point as the renderer rounds it.

**The typeface.** As the template states it on the run that supplies the formatting (`styleRun`), or as the
run inherits it: from the shape's list style, its layout and master placeholders, the master's text styles,
the presentation's defaults, and finally the theme's fonts.

**Which typeface a character is set in.** A run names up to three typefaces: Latin, complex script (Hebrew,
Arabic, Thai and others) and East Asian. Hebrew letters are set in the complex-script typeface, not the Latin
one. When the template names none, the theme decides, and a theme often sends Hebrew to Arial or Times New
Roman whatever the Latin font is. In a paragraph that starts with a right-to-left letter, digits are set in
the complex-script typeface too, and a space or punctuation mark in the typeface of the character before it.
In other paragraphs digits, spaces and punctuation are set in the Latin typeface.

**Width.** The sum of the glyph widths from the font file, after the font's own shaping (Arabic joining,
ligatures), with each glyph rounded to an eighth of a point. PowerPoint lays text out on that grid, and at
small sizes the rounding changes a line's width by a percent or two, which is enough to move a word to the
next line. Character spacing set on the run is added. Text set in capitals by the template is measured in
capitals.

**Kerning.** Hebrew and other complex-script letters are always kerned with each other, at any size, and
never with the punctuation, digits or spaces next to them. Latin text is kerned when the template turns
kerning on for its size (commonly from 12 pt up), but not inside a paragraph that starts with a right-to-left
letter.

**Line breaks.** Words go on a line while they fit. A line may break at a space and after a hyphen or a dash;
not at a slash, a full stop, a colon or a Hebrew maqaf. A word longer than a whole line is cut where the line
ends. Chinese and Japanese text breaks between any two characters. A hanging indent or a first-line indent
changes the first line's width. The space at the end of a line hangs outside the box and needs no room,
except when it sits between two words that both run against the paragraph (two Hebrew words in an English
paragraph, or the reverse): then the space stays inside the line.

**Height.** Lines are 1.2 times the font size apart, whatever the font: this is how PowerPoint spaces lines.
Line spacing set on the paragraph (a percentage, or exact points) is applied, and so is the space between
paragraphs, including a slot's `paragraphGap`. The space before the first paragraph and after the last is not
drawn and not counted.

**Fitting.** A text fits when its height is within the room, with a slack of three tenths of a line: a last
line that hangs that little under the box is not reported. A box that does not wrap fits when its longest
line is no wider than the box.

## How close it is to PowerPoint

Checked on this project's demo template against PowerPoint 16 (build 17932) on Windows, which was asked through
its automation interface for the lines of every text box:

- **At the designed sizes (12 to 84 pt): 1,153 text slots in 17 decks.** The two example decks, ten decks that
  fill every slot with running text from half its budget to 1.7 times it in English and in Hebrew, four of
  those at text scale 1.25 and 1.5, and the studio's sample deck. In every slot, every line broke at the same
  word as in PowerPoint, and the text height was the same to within 0.6 pt.
- **With every font size halved (body text of 6 to 10 pt): 700 slots in 10 decks.** 699 had PowerPoint's
  number of lines and 697 broke every line at the same word. The three others are lines whose width equals the
  width of their box to within an eighth of a point, where PowerPoint's own rounding of the box decides.
- **Shrinking:** in three decks rendered with `shrinkText` (210 slots, 108 of them shrunk), PowerPoint showed
  every text at the size the measurer chose, in the number of lines it predicted.
- **Inherited styles:** a slide whose title and bulleted body state nothing themselves and take size,
  typeface, position, bullets and indents from the layout and master (44 pt, 32 pt and a second level at
  28 pt) gave the same lines and the same heights as PowerPoint.
- **With the open stand-ins** (Carlito and Caladea) the English decks gave PowerPoint's line counts in every
  slot at the designed sizes and in 278 of 280 with halved sizes. About one slot in ten had a word on the
  other side of a break: the stand-ins match the widths closely, not exactly.

Six of these decks are kept in the repository with the line breaks PowerPoint made, and the test suite
compares against them on every run: exactly with the template's own fonts where they are installed, and by
line count with the open stand-ins everywhere.

The demo template is set in Calibri and Cambria. Expect agreement of the same kind for other templates when
the font files are the ones PowerPoint uses, and treat the following as less certain:

- **Fonts that differ.** The measurement is only as good as the files. PowerPoint on a machine that lacks the
  template's font substitutes another, and lines then break elsewhere, whatever was measured.
- **Tables and charts** are not measured; they keep their character budgets.
- **Vertical text, text in columns, and outlines other than rectangles, rounded rectangles, ellipses and
  diamonds** are measured roughly and marked `approximate`.
- **PowerPoint's own autofit.** A box set to "shrink text on overflow" stores the scale PowerPoint last
  applied; the measurer uses that stored scale. PowerPoint works a new one out only when the text is edited
  there.
- **Several styles in one paragraph.** A slot's text is rendered in one run's style (`styleRun`), and that is
  what is measured.
- **Line breaks by dictionary** (Thai, Khmer) and East Asian punctuation rules are simplified.

## Using another font library

`fontkitMetrics` is one implementation of a small interface. Implement it over any library:

```ts
import type { FontFace, FontMetrics } from 'slotdeck';

const fonts: FontMetrics = {
  // the face for a typeface name in a weight and slant, or undefined when you do not have it
  face(typeface, style): FontFace | undefined {
    const font = myLibrary.find(typeface, style.bold, style.italic);
    if (!font) return undefined;
    return {
      has: (codePoint) => font.hasGlyph(codePoint),
      // the advance of each glyph of the text, in ems: the library's units divided by the font's units per em
      advances: (text, kerning) => font.shape(text, { kerning }).map((glyph) => glyph.advance / font.unitsPerEm),
    };
  },
  // optional: a face for a character the chosen face lacks
  fallback(codePoint, style) {
    return undefined;
  },
};
```

`advances` is called with a piece of text in one script and one typeface (a word, the letters of a Hebrew
word, or the space between two words), so the library should shape it: for Arabic, the joined forms have other
widths than the letters alone. It returns one number for each glyph rather than a total, because slotdeck
rounds each glyph the way PowerPoint does. With `kerning` true, the kerning between two glyphs belongs in the
advance of the first. Set `substitute: true` on a face that stands in for another weight or slant; results
that use it are marked `approximate`.

## Checking against PowerPoint yourself

The repository has the scripts the figures above were made with. They need Windows with PowerPoint and are not
part of the package.

```
node scripts/text-fit-stress.ts                      # writes out/stress-*.pptx and their deck JSON
powershell -File scripts/measure-pptx.ps1 -Path out/stress-he-3.pptx -Out out/stress-he-3.lines.json
node scripts/compare-text-fit.ts out/stress-he-3.deck.json out/stress-he-3.lines.json
```

`measure-pptx.ps1` writes, for every text box of a pptx, the lines PowerPoint breaks its text into, the width
of each line and the height the text takes. `compare-text-fit.ts` prints every slot where the measurer
disagrees. `text-fit-stress.ts --small` writes the decks with halved font sizes.
