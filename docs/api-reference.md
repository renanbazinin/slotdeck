# API reference

Every export of the package, grouped by what it is for. Types are written as TypeScript. Everything is imported
from `slotdeck` unless noted.

```ts
import { renderDeck /* ... */ } from 'slotdeck';
import { directoryStore } from 'slotdeck/node';
```

The package ships its own type declarations (`dist/*.d.ts`), which carry the same information with comments on
every field.

## Rendering

### `loadTemplate`

```ts
function loadTemplate(pptx: Uint8Array, manifest: TemplateManifest, options?: TemplateOptions): Template;
```

Unzips a template once. Reuse the result. `options.fonts` gives it font metrics, which turns on text measuring
everywhere the template is used (see "Measuring text" below).

### `renderDeck`

```ts
function renderDeck(template: Template, deck: DeckSpec, options?: RenderOptions): RenderResult;
```

Deck JSON to pptx bytes. Synchronous, no model. Throws `DeckValidationError` when the deck has errors.

With `options.properties` (from 0.7.0) the file says what it is as a deck, not as its template: a title (the
given one, else the deck's `title`, else its first slide's title), a language, an author and a date where they
are given, its real number of slides, and the version of slotdeck that made it as the custom property
`slotdeck`. The template's own title, author, dates, its lists of sample slide titles and its picture of its
cover, which Outlook and Explorer show for the file, are taken out; its other custom properties, such as a
label of how confidential it is, stay. A date is written only where it is given, so that the same deck gives
the same file. Without the option the file keeps the template's properties and is the same file as before
0.7.0. `generateDeck`, `editDeck` and the chat tools set it unless their `properties` is `false`.

### `readEmbeddedDeck`

```ts
function readEmbeddedDeck(pptx: Uint8Array): EmbeddedDeck | undefined;
interface EmbeddedDeck { deck: DeckSpec; version: string; assets: Record<string, Uint8Array> }
```

Takes out of a pptx the deck that `renderDeck` kept in it with `properties: { embedDeck: true }`, the version
of slotdeck that made the file, and the bytes of its pictures by asset id: `renderDeck(template, found.deck,
{ assets: found.assets })` renders it again. The deck kept has no `brief` and no `document`. `undefined` for a
file that keeps none; PowerPoint keeps it when the file is saved again. Throws for bytes that are no pptx.

### `SLOTDECK_VERSION`

```ts
const SLOTDECK_VERSION: string; // '0.7.1'
```

The version of the package, as a file rendered with `properties` states it.

### `validateDeck`

```ts
function validateDeck(target: Template | TemplateManifest, deck: DeckSpec, options?: ValidateOptions): Issue[];
```

Checks a deck with the renderer's own rules, against a loaded template, or a manifest alone.

### `extractAssets`

```ts
function extractAssets(template: Template, deck: DeckSpec, pptx: Uint8Array): Record<string, Uint8Array>;
```

Reads a deck's pictures back from a pptx rendered from it, by asset id.

### `slideAlternatives`

```ts
function slideAlternatives(manifest: TemplateManifest, deck: DeckSpec, slide: number, options?: AlternativesOptions): SlideAlternative[];

interface AlternativesOptions { layout?: LayoutMode; patterns?: string[] }
interface SlideAlternative {
  pattern: string;      // a pattern that has a place for everything the slide says
  description: string;  // what that pattern is for
  slide: SlideSpec;     // the slide written for that pattern
  changes: string[];    // what was done where the content was not copied slot for slot
  dropped: string[];    // slots of the original whose photo or icon has no place
  warnings: Issue[];    // what validateDeck says about the new slide
}
```

The other patterns that can hold what slide number `slide` (from 1) of a deck says, with the slide written for
each. Synchronous, no model, no words changed. Throws a `RangeError` when the deck has no such slide. See
"Another layout for the same slide" in [layout-freedom.md](layout-freedom.md).

### `DeckValidationError`

```ts
class DeckValidationError extends Error {
  readonly issues: Issue[];
}
```

### `isImageRef`, `isTableValue`, `isChartValue`

```ts
function isImageRef(value: unknown): value is ImageRef;
function isTableValue(value: unknown): value is TableValue;
function isChartValue(value: unknown): value is ChartValue;
```

Tell the kinds of slot value apart when walking a deck.

### Types

```ts
interface Template { pkg: PptxPackage; manifest: TemplateManifest; fonts?: FontMetrics }
interface TemplateOptions { fonts?: FontMetrics }

interface RenderOptions { assets?: Record<string, Uint8Array>; layout?: LayoutMode; properties?: DeckProperties }
interface DeckProperties {
  title?: string;        // left out: the deck's title, else its first slide's
  language?: string;     // left out: the deck's, else the template's
  author?: string;       // left out: no one is named
  date?: Date | string;  // left out: no date, so the same deck gives the same file
  embedDeck?: boolean;   // keeps the deck in the file, for readEmbeddedDeck
}
interface RenderResult {
  pptx: Uint8Array;
  warnings: Issue[];
  shrunk: Array<{ slide: number; slot: string; size: number; scale: number }>; // texts set smaller to fit; empty without shrinkText
}

type LayoutMode = 'strict' | 'adaptive';
interface ValidateOptions { layout?: LayoutMode; assets?: Iterable<string>; measure?: TextMeasurer }
interface Issue {
  severity: 'error' | 'warning';
  slide?: number;                                   // from 1; absent for an issue about the deck as a whole,
                                                    // and for one whose sentence names its slide (values.slide)
  slot?: string;
  message: string;                                  // the English sentence
  code?: IssueCode;                                 // from 0.6.4, on every warning: what it is (see below)
  values?: Record<string, string | number | boolean>; // with code: the numbers and labels its sentence states
}

interface DeckSpec {
  version: 1;
  title?: string;
  brief?: string;
  language?: string;
  direction?: Direction;
  chartDirection?: Direction;   // the direction its charts read; left out, the manifest's for a deck that reads right to left
  fields?: Record<string, string>;
  textScale?: number;
  layouts?: PatternDesign[];    // layouts made for this deck, built onto the template when it is rendered
  colors?: DeckColors;          // colours in place of the template's, for every slide
  patterns?: string[];          // the only patterns the deck was asked to use; the planner keeps to it
  slides: SlideSpec[];
  documentLanguage?: string;    // the language that document is written in, as the planner read it (BCP 47)
  document?: string;            // the document the deck was made from (Brief.document), word for word
}
interface SlideSpec {
  pattern: string;
  items?: number;
  slots: Record<string, SlotValue>;
  blocks?: FreeBlockKind[];     // makes it a free slide (pattern 'free'): its parts in reading order
  arrange?: FreeArrange;        // a free slide: how its parts are set out
  base?: string;                // a free slide: the pattern whose frame it stands on
  colors?: DeckColors;          // colours for this slide alone, in place of the deck's
  passages?: number[];          // a deck made from a document: the passages of it this slide carries
  headings?: string[];          // a deck made from a document: its headings in the slide's text, set bold, run in (from 0.6.6)
}
interface DeckColors {
  accent?: string;                  // takes the place of the template's main colour; six hex digits
  accent2?: string;                 // the same for its second colour
  replace?: Record<string, string>; // exact replacements of single colours, from one hex to another
}
type SlotValue = string | string[] | ImageRef | TableValue | ChartValue | null;
interface ImageRef { asset: string; alt?: string; prompt?: string }
interface TableValue { header: string[]; rows: string[][] }
type ChartType = 'column' | 'bar' | 'line' | 'pie';
interface ChartValue {
  type: ChartType;
  categories: string[];
  series: Array<{ name: string; values: Array<number | null> }>; // null: a value that is not known, drawn as a gap
  format?: 'number' | 'percent';
}
```

## Warnings in a person's language

From 0.6.4 every warning has a `code` and `values`: what it is, and the numbers and labels its sentence
states. So does every note the chat tools give a person (`ToolResult.issues`). The English `message` is as
it was. An app shows the issue in its own language with `formatIssue`.

### `formatIssue`

```ts
function formatIssue(issue: Issue, language: string, options?: FormatIssueOptions): string;

interface FormatIssueOptions { deck?: Pick<DeckSpec, 'slides'> }
```

The issue as a person reads it. `language` is a BCP 47 tag: `en` gives the English sentence and `he` (or
`he-IL`) the Hebrew one. Any other language, an issue without a code, and an issue whose values cannot make
the sentence give the English `message`.

With `options.deck` it is the whole line, as the chat tools write it in `output.warnings`: an issue with a
slide begins with it and its slot (`slide 4 body: `), or with the part of a free slide
(`slide 3, part 7 (note): `); one without a slide is its sentence alone. In English that line is exactly the
tool's. In Hebrew the slide, the part and the slot are named in Hebrew, each after a comma
(`שקף 3, חלק 7 (הערה): `, `שקף 1, כותרת: `, `שקף 2, כותרת של כרטיס 1: `). A slot has a Hebrew name where most
templates name it alike: `title`, `subtitle`, `body`, and the `Heading` and `Text` of `column`, `col`, `card` and
`point` with a number (`card1Heading`). Any other slot keeps its name. Without `deck` it gives the sentence
alone.

A Hebrew sentence begins with a Hebrew word, so a line set to the direction of its first letter
(`dir="auto"`) reads from the right. A label, a name or a value from the deck stands in an isolate of its own
(U+2068 and U+2069, invisible), so that `p1-p14`, `3-cards` or `-5, 12.5` reads as written. Strip them only if
you also give the text a direction of its own. A Hebrew sentence states the numbers of the English one in
digits, but for a count of one, which it says with the word after the noun, as Hebrew does (`שורה אחת`).

```ts
const lines = (result.issues ?? []).map((issue) => formatIssue(issue, uiLanguage, { deck: result.deck }));
```

### `ISSUE_CODES`

```ts
const ISSUE_CODES: readonly IssueCode[];
```

Every code, in the order of the table below.

### Types

```ts
type IssueCode = keyof IssueValues; // 'text-cut' | 'words-cut' | ... : the codes below
interface IssueValues { 'text-cut': { label: string; chars: number; holds: number; freeOff: boolean }; /* ... one entry for each code */ }
```

A code is never renamed once it has shipped. A new version may add codes; treat one you do not know as its
`message`. `slide` in the values is the slide a sentence names itself, where the issue's own `slide` is left
out on purpose (no slide is sent back to its writer for it); `formatIssue` then adds no prefix. `freeOff` says
the sentence ends that free slides are off, `size` (points) that the text is set small, `label` names the
passages of a document (`p3-p5`). Text in `given`, `format`, `from`, `to` and `word` is as the deck gave it,
in JSON.

| Code | Values | Said when |
| --- | --- | --- |
| `text-too-wide` | `approximate`, `percent`, `shrunk`, `size`, `fitsChars` | A one-line box's text is wider than the box (with fonts). |
| `text-overflow` | `approximate`, `lines`, `shrunk`, `size`, `capacity`, `over`, `fitsChars` | Measured text runs over its box by `over` lines (with fonts). |
| `colors-not-object` | `slide?` | `colors` is no object. |
| `color-not-hex` | `slide?`, `name`, `given` | A colour is not six hex digits. |
| `colors-replace-not-map` | `slide?` | `replace` is no map of colours. |
| `colors-replace-not-pair` | `slide?`, `from`, `to` | A replacement is not a pair of colours. |
| `colors-unknown-setting` | `slide?`, `name` | `colors` has a setting it does not know. |
| `colors-kept-by-template` | none | The manifest says `"colors": false`, and the deck asks for colours. |
| `columns-not-whole` | `slide`, `given` | A free slide's `arrange.columns` is no whole number of 1 or more. |
| `columns-relaxed` | `slide`, `asked`, `columns` | A free slide holds fewer or more side by side than asked. |
| `free-text-small` | `slide`, `blocks`, `size` | A free slide's smallest text is under 9 pt (on a slide 13.33 inches wide), or under the template's `leastTextSize` where that is lower. |
| `free-larger-text` | `slide`, `blocks`, `half` | A free slide has room for half of the deck's larger text, or none. |
| `free-charts-bare` | `slide`, `blocks`, `charts` | Charts of a free slide are drawn without axes, legend or labels. |
| `free-blocks-low` | `slide`, `blocks`, `count` | Blocks of a free slide reach below its room. |
| `free-base-refused` | `slide`, `base`, `why`, `stands` | A free slide's `base` cannot be a frame; `why` is the reason in English (the Hebrew says the usual ones in Hebrew). |
| `free-base-gone` | `slide`, `base`, `stands` | A free slide's `base` is no longer in the template. |
| `items-ignored` | `pattern` | `items` on a pattern with a fixed number of items. |
| `slot-unknown` | `pattern` | A slot the pattern does not have. |
| `slot-over-chars` | `chars`, `most` | Text over the slot's characters. |
| `slot-over-paragraphs` | `paragraphs`, `most` | Text over the slot's paragraphs. |
| `table-over-columns` | `columns`, `most` | A table over the slot's columns. |
| `table-over-rows` | `rows`, `most` | A table over the slot's rows. |
| `table-cell-long` | `chars`, `most` | A table cell over what a cell holds. |
| `chart-over-categories` | `categories`, `most` | A chart over the slot's categories. |
| `chart-over-series` | `series`, `most` | A chart over the slot's series. |
| `chart-format-unknown` | `format` | A chart format that is not `number` or `percent`. |
| `chart-values-unstated` | `numbers`, `count` | Chart values the brief does not state (planner only). |
| `text-details-unstated` | `details`, `count` | From 0.7.0: email or web addresses, phone numbers, "24/7" or percentages in a slide's text that the brief, the document and the deck's fields do not give (planner only). |
| `word-wider-than-line` | `word`, `chars` | From 0.7.0, without fonts: a word no line of its box holds (an address, a long number), which PowerPoint breaks inside it. |
| `text-stand-in` | `found` | From 0.7.0: text that stands where a value belongs, such as "XX%", "[name]" or "TBD" (planner only). |
| `text-about-slide` | `found` | From 0.7.0: text that speaks about the slide ("this slide", "בשקף זה") where it should say its content (planner only). |
| `text-orphan` | `lines`, `word`, `chars` | A heading ends with one word on its last line (`checkDesign`). |
| `text-empty` | `lines`, `sampleLines?` or `capacity?`, `chars` | Text leaves its box almost empty (`checkDesign`). |
| `row-uneven` | `lines`, `beside`, `besideLines`, `chars` | A card holds far less than the one beside it (`checkDesign`). |
| `seen-in-picture` | `problem` | A review model saw it in a picture of the slide; `problem` is its own sentence. |
| `accent2-missing` | `slide?` | A second colour asked of a template that has none (`renderDeck`). |
| `color-near-page` | `hex` | A colour close to the page it lands on (`renderDeck`). |
| `part-unwritten` | none | Sent to the writer only: a part came back empty. |
| `photos-over` | `most` | Sent to the writer only: more photos described than are made for a slide. |
| `parts-left-out-empty` | `slide`, `count` | Parts of a free slide had no content and were left out. |
| `parts-left-out-unusable` | `slide`, `count` | Parts of a free slide could not be shown and were left out. |
| `photos-left-out` | `slide`, `count`, `most` | Photos over the most made for one slide were left out. |
| `parts-unchanged` | `slide`, `count` | What a change wrote for parts could not be shown; they stay as they were. |
| `parts-unwritten` | `slide`, `count` | A change wrote nothing for parts; they stay as they were. |
| `written-under-retry` | `chars`, `words`, `least`, `most` | Sent to the writer only: a text well short of the words asked. |
| `written-under` | `words`, `hasWords` | A slide is still well short of the words asked. |
| `words-cut` | `words`, `holdsWords`, `freeOff` | A slide holds fewer words than were asked. |
| `words-cut-dense` | `words`, `holdsWords`, `size` | The same, with its text set small. |
| `photo-left-out` | `words`, `asked` | A photo gave way to a text (edits only). |
| `words-kept-low-retry` | `percent`, `label` | Sent to the writer only: a slide keeps too few words of its passages. |
| `words-kept-low` | `percent`, `label` | A slide still keeps too few words of its passages. |
| `text-cut` | `label`, `chars`, `holds`, `freeOff` | A slide holds less of its document's text than it carries. |
| `paragraphs-joined` | `label`, `paragraphs`, `holdsParagraphs` | A slide holds its document's text only with its paragraphs joined. |
| `text-cut-dense` | `label`, `chars`, `holds`, `size` | A slide holds less of its document's text even set small. |
| `nothing-changed` | `slide` | A change named no part of a free slide to write again. |
| `recolor-slide-removed` | `slide` | A change removed a slide that was to be recoloured. |
| `change-would-cut` | `slide`, `words`, `holdsWords`, `size?`, `freeOff` | Note of `edit_presentation` that stops before a cut, one for each slide (`TextCutError.issues`). |
| `slide-holds-less` | `slide`, `words`, `holdsWords`, `size?`, `freeOff` | Note of a plan: a slide whose text is cut. |
| `slide-set-small` | `slide`, `size`, `templateSize` | Note of a plan: a slide that holds all its text, in type well below the template's. |
| `text-retold` | `name` | Note of a plan: the topic retells a listed text, which no slide carries word for word. |

The sentences of every code, in English and in Hebrew, are in `dist/messages/en.js` and `dist/messages/he.js`
of the package. "Deck warnings" in [troubleshooting.md](troubleshooting.md) lists the warnings with what to do.

## Generating and editing

### `generateDeck`

```ts
function generateDeck(template: Template, brief: Brief, options: GenerateOptions): Promise<GenerateResult>;
```

Brief to pptx: plans, gathers pictures, renders.

### `editDeck`

```ts
function editDeck(template: Template, deck: DeckSpec, instruction: string, options: EditDeckOptions): Promise<EditDeckResult>;
```

Applies an instruction to a deck, gathers pictures, renders.

### `planDeck`

```ts
function planDeck(template: Template, brief: Brief, options: PlanOptions): Promise<PlanResult>;
```

Brief to deck JSON: `planOutline` then `writeDeck`.

### `planOutline`

```ts
function planOutline(template: Template, brief: Brief, options: PlanOptions): Promise<DeckPlan>;
```

Decides which pattern each slide uses and the point it makes. One model call. Where free slides are on offer
(the `free` option, on unless `false`) a slide may be a free slide: its entry then has `pattern: 'free'` and lists
its parts in `blocks`. With `composeLayouts` set and free slides not on offer, it then makes new layouts for
slides that none of the template's layouts shows well, as `composeLayouts` below does. See "Free slides" in
[generating.md](generating.md).

### `composeLayouts`

```ts
function composeLayouts(template: Template, plan: DeckPlan, options: PlanOptions & { max?: number }): Promise<ComposeLayoutsResult>;
```

Looks at a plan for slides whose content none of the template's layouts shows well, such as three goals on a
slide of running text, and makes up to `max` (default 2) new layouts for them in the template's own look. One
light model call decides, then each layout is composed with `composePattern`, side by side, on the frame of the
layout its first slide had. The returned plan has the designs in `layouts` and its slides use them; the
template is not changed. A layout that cannot be made is listed in `failed`, and its slides keep their layout.
`planOutline` does the same when `composeLayouts` is set in its options. Nothing is made for a plan that is
limited to a list of patterns, nor while free slides are on offer: pass `free: false` to have layouts made. See
"New layouts for a deck" in [generating.md](generating.md).

### `withLayouts`

```ts
function withLayouts(template: Template, layouts: readonly PatternDesign[] | undefined): Template;
```

The template with a deck's own layouts (`DeckSpec.layouts` or `DeckPlan.layouts`) built onto it, in order: the
template to check or show such a deck with. `renderDeck`, the planner, the edits and the reviews call it
themselves. A layout whose id the template already has is not built: the template's own is used. Built once
for a template object and a list of layouts. Throws `DesignError` when a layout no longer fits the template.

### `editPlan`

```ts
function editPlan(template: Template, plan: DeckPlan, instruction: string, options: PlanOptions & PlanSettings): Promise<{ plan: DeckPlan; summary: string }>;
```

Changes a plan by instruction. One model call. A free slide of the plan keeps its parts unless the instruction
changes what that slide shows, and stays set out as it was wherever the answer does not say otherwise. From
0.6.3 a slide whose text from a document would be cut beside the parts the plan had, or whose parts leave it
under half of the words asked, is its text alone, set dense, where free slides are on offer and that keeps more
of it. A part a change asks for stays, in that change and in the changes after it: the plan marks the slide
(`OutlineEntry.partsAsked`, internal), in a plan of any kind. That is so for the parts of a free slide and for a
layout's chart or table. A layout of cards, figures, a list or a quote gives way to an amount of text asked
for it, as in 0.6.2. The answer's slides are matched with the plan's by their parts in any order, of the
same purpose first, then in the same place, then in order.

### `writeDeck`

```ts
function writeDeck(template: Template, plan: DeckPlan, options: PlanOptions): Promise<PlanResult>;
```

Writes the content of a plan's slides. One model call plus repairs. The counts of the plan are checked before
any call: `items` that is not a whole number the pattern's row takes, or a free slide of more than 400 parts,
is a `PlanError` that names the slide. A free slide of more than nine parts is
written by calls of its own, three at a time, before that call. A part of a free slide that comes back without
content is taken off the slide and named in `warnings`; it does not make the function throw. Of the photos a
writer describes for one free slide, six are kept, and the warnings name the rest. Every number a chart shows
must appear in the brief, or `warnings` names it; the check does not tell whether a number stands in the right
chart. A slide asked for words whose slot was told the whole length (`Write 500 words, as was asked`) and came
back under 0.85 of the least it was told is sent back once; still short, `warnings` says so.

### `planTextFit`

```ts
function planTextFit(template: Template, plan: DeckPlan, options?: Pick<PlanOptions, 'patterns' | 'free' | 'icons' | 'assets' | 'layout' | 'numberPatterns' | 'countSpacing'>): SlideTextFit[];
```

What each slide of a plan that has asked words or passages of a document will carry, worked out by code
alone before anything is written: the words it holds, the size its running text is set at on a free slide, and
whether it is set dense. A slide that holds less than its words is cut to what it holds when the deck is
written. `[]` for a plan with neither.

### `documentPassages`, `MOST_DOCUMENT_CHARS`

```ts
function documentPassages(document: string): DocumentPassage[];
const MOST_DOCUMENT_CHARS = 12000; // the default of maxDocumentChars
```

The passages of a document (`Brief.document`), numbered as the planner numbers them: each paragraph, a long one
cut at the ends of its sentences, short lines of one section together, a list, a table. A short line on its own,
or one that starts with `#`, is a heading over the passages after it; a run of three short lines or more with no
stops, one under another, is text, all of it. A tab at the start of a line, as text pasted from a word processor
has it, sets the line in and makes no table. Pure: the same text always gives the same passages. A new brief's
document is at most `maxDocumentChars` characters of `PlanOptions`; a longer one is a `PlanError` from
`planOutline`. `MOST_DOCUMENT_CHARS` is the default of that option: 12000 characters, about 1,800 words. From
0.7.1 the option may be set from 1000 to 30000, and the later stages read a stored plan's document up to 30000.

### `planEdit`

```ts
function planEdit(template: Template, deck: DeckSpec, instruction: string, options: EditOptions): Promise<EditResult>;
```

Applies an instruction to a deck and returns the new deck JSON, without pictures or rendering. A free slide is
changed part by part: the parts the instruction does not touch are carried over by code, and only parts that
are new or change are written. The edit can also make a free slide, new or out of a slide of a pattern, where
free slides are on offer. `rewritten` lists the slides that were written, and the free slides whose parts or
arrangement code changed with nothing to write. A part that was to be written again and came back with nothing
keeps what it had, and `warnings` says that the change was not made. It says so too for a slide the plan of the
change had words for while it marked no part to be written, when the slide's title came back as it was
(`slide 3: nothing on it was changed: ...`, with no slide number in the issue; the slide is not in
`rewritten`). How a slide is set out is
changed by a value the plan states; the plan's `auto` and 0 hand a slide back to the software only where the
plan changes that slide. A title the writer leaves out stays, and the title of a slide of ten parts or more is
changed, not removed. See "Free slides in an edit" in [generating.md](generating.md). A slide asked for words
keeps them on that slide: the edit adds no slide for a text. With `stopBeforeCut`, a slide that would hold less
than was asked for it stops the edit before anything is written, with a `TextCutError`; what it holds is counted
as it will be written, with the heading or caption it keeps. A part added beside a dense text, or a text from the
document, leaves the text less room: a text that still holds there stays as it is, and one that does not is
written again for that room, and cut. Of a deck made from a document (`DeckSpec.document`), the plan of the
change names the passages each slide carries; a slide whose passages change is written again, its text set by
code as the document has it where it holds it and the document is in the deck's language
(`DeckSpec.documentLanguage`), unless the plan says the instruction asks for that text in other words.

### `recolorDeck`

```ts
function recolorDeck(deck: DeckSpec, colors: DeckColors | null, slides?: number[]): DeckSpec;
```

Sets colours on a deck, such as `{ accent: '1E6B3A' }`, and returns the new deck: no model is asked, and the
deck given is not changed. Without `slides` they are the deck's colours; with `slides` (numbers from 1) those
slides get them as their own; `null` removes them. An empty `slides` names no slide and is read as if it were
left out.

The edits set the same by code, from their options, with no model reading a colour out of the instruction:

```ts
// editPlan takes these beside PlanOptions, and writes them to the plan's brief.
interface PlanSettings {
  colors?: DeckColors | null;      // null returns to the template's colours
  deckPatterns?: string[] | null;  // the person's list of patterns; null or [] lifts the limit
}

// planEdit takes these in EditOptions.
//   colors?: DeckColors | null
//   colorSlides?: number[]         numbers of the deck before the edit; empty means the whole deck
//   deckPatterns?: string[] | null
```

`deckPatterns` is the person's list, kept in the plan and the deck (`Brief.patterns`, `DeckSpec.patterns`).
`patterns` in the same options is still the app's own list, which narrows what is offered and is not kept.
With `colorSlides`, `EditResult.recolored` holds the numbers those slides have in the new deck, and a slide
the edit removed is named in `warnings`.

### `resolveAssets`

```ts
function resolveAssets(template: Template, deck: DeckSpec, options?: ResolveOptions): Promise<Record<string, Uint8Array>>;
```

Collects the bytes of every picture a deck refers to.

### `directionOf`

```ts
function directionOf(language: string): Direction;
```

The reading direction of a BCP 47 language tag.

### `isPlainText`

```ts
function isPlainText(pattern: PatternDef): boolean;
```

Whether a slide made from the pattern is running text and nothing else: the patterns a text-only deck may use.

### `PlanError`

```ts
class PlanError extends Error {}
```

Thrown when no usable plan can be made, and by `writeDeck` for a plan whose counts cannot be written (a
number of items the pattern does not take, a free slide of more than 400 parts). The message says why.

### `TextCutError`

```ts
class TextCutError extends PlanError {
  readonly fits: SlideTextFit[]; // the slides that would hold less than they carry, numbered as after the change
  readonly freeOff: boolean;     // from 0.6.4: free slides are off, so no text is set smaller
  readonly issues: Issue[];      // from 0.6.4: a note for each slide (`change-would-cut`), for formatIssue
}
```

Thrown by `planEdit` with `stopBeforeCut` when a slide it writes would hold less text than was asked for it, or
than it carries.
Nothing was written. The message names each slide and the words it would hold, counted as the slide would be
written. A slide whose parts and passages stay, asked for no length (other words, a new title, another
language), does not throw it.

### `TEXT_SIZES`

```ts
const TEXT_SIZES: { normal: 1; large: 1.25; 'x-large': 1.5 };
type TextSize = 'normal' | 'large' | 'x-large';
```

### Types

```ts
interface Brief {
  prompt: string;
  slideCount?: number | { min: number; max: number };
  language?: string;
  direction?: Direction;
  imageDensity?: ImageDensity;
  instructions?: string;
  fields?: Record<string, string>;
  textSize?: TextSize;
  textOnly?: boolean; // every slide plain text: only the template's plain text patterns are offered
  patterns?: string[]; // the only patterns the person asked for, `free` among them for free slides; kept with the plan and the deck
  colors?: DeckColors; // colours in place of the template's; copied to the deck
  document?: string; // a document to make the deck from, word for word; at most maxDocumentChars characters (12000, MOST_DOCUMENT_CHARS, unless set) when the plan is made
}
type ImageDensity = 'none' | 'low' | 'medium' | 'high';

interface PlanOptions {
  llm: LlmProvider;
  delivery?: 'present' | 'read'; // from 0.7.0: 'read' for decks read on their own: full sentences, no thank-you slide
  layout?: LayoutMode;
  icons?: IconDef[];
  assets?: AssetDef[];
  patterns?: string[];
  numberPatterns?: 'from-brief' | 'always' | 'never';
  maxRepairs?: number;
  maxDocumentChars?: number; // from 0.7.1: the most characters of a new brief's document; 12000 unless set, a whole number from 1000 to 30000, else a RangeError
  review?: DeckReview; // a second look at every draft; see "Design checks and reviews"
  composeLayouts?: number; // up to this many new layouts for the deck, in the template's own look; off (0) by default
  free?: boolean; // whether a slide may be a free slide, laid out by the software from the parts the plan names; on unless false
  countSpacing?: boolean; // a writer of a layout's slot of paragraphs is told what its box holds, the template's spacing counted; on unless false
  onProgress?: (event: PlanProgress) => void;
  signal?: AbortSignal;
}
type PlanProgress =
  | { stage: 'outline' }
  | { stage: 'layouts' }
  | { stage: 'write'; slides: number }
  | { stage: 'repair'; attempt: number; issues: number };

interface IconDef { id: string; description: string; data: Uint8Array }
interface AssetDef { id: string; kind: 'photo' | 'graphic'; description: string; data: Uint8Array }

interface OutlineEntry {
  pattern: string;
  purpose: string;
  items?: number;
  blocks?: FreeBlockKind[]; // a free slide (pattern 'free'): its parts in reading order
  arrange?: FreeArrange;    // a free slide: how its parts are set out, where the plan says so
  words?: number;           // the words of running text asked for this slide ("200 words"), 1 to 3000; never spread over more slides
  passages?: number[];      // a deck made from a document: the passages this slide carries
  partsAsked?: true;        // internal, set by editPlan: a free slide's parts, or a layout's chart or table, asked for in a change; they stay
}
interface DocumentPassage {
  number: number;           // from 1, in the document's order
  kind: 'text' | 'list' | 'table';
  text: string;             // as the document has it: a list one item to a line, a table one row to a line
  section?: string;         // the heading it stands under
  continues?: true;         // goes on from the passage before it
  leads?: true;             // ends with a colon, and goes on a slide with the passage after it
  chars: number;            // the characters a slide shows of it
  paragraphs: number;       // the paragraphs a slide shows of it
}
interface SlideTextFit {
  slide: number;            // from 1
  words: number;            // the words asked for the slide, or those of its passages
  holds: number;            // the words it holds: `words` when it holds them all
  size?: number;            // the size its running text is set at, in points, on a free slide
  templateSize?: number;    // the size running text has on a free slide of the template when nothing makes it smaller
  dense?: true;             // set dense (FreeArrange.dense)
  most?: number;            // set dense: the most words of a text of up to 12 paragraphs it holds at the least size that reads; `holds` counts its own words in its own paragraphs, and can be more or less
  passages?: string;        // a slide of a document: its passages, as "p3-p5"
}
interface DeckPlan {
  version: 1;
  title: string;
  brief: Brief;
  slides: OutlineEntry[];
  layouts?: PatternDesign[];
  documentLanguage?: string; // a deck made from a document: the language the document is written in, as the planner read it (BCP 47)
}
interface PlanResult { deck: DeckSpec; warnings: Issue[]; outline: OutlineEntry[] }

interface ComposeLayoutsResult {
  plan: DeckPlan;               // with layouts, and the slides they were made for using them
  made: MadeLayout[];
  failed: FailedLayout[];
}
interface MadeLayout { pattern: string; slides: number[]; summary: string }
interface FailedLayout {
  instruction: string;          // what the layout was to be; '' when the deciding call itself failed
  slides: number[];             // they keep the layout they had
  reason: string;
}

interface EditOptions extends PlanOptions {
  fields?: Record<string, string>;
  direction?: Direction;
  imageDensity?: ImageDensity;
  colors?: DeckColors | null;       // sets the deck's colours by code, as recolorDeck; null returns to the template's
  colorSlides?: number[];           // gives `colors` to these slides alone, numbered as before the edit
  deckPatterns?: string[] | null;   // the person's list of patterns from now on; null or [] lifts the limit
  stopBeforeCut?: boolean;          // a slide that would hold less than was asked for it throws TextCutError before writing
}
interface EditResult {
  deck: DeckSpec;
  summary: string;
  rewritten: number[];
  recolored?: number[];             // with colorSlides: the numbers those slides have in the new deck
  warnings: Issue[];
}

// generateDeck and editDeck add picture handling to the options above:
interface GenerateOptions extends Omit<PlanOptions, 'onProgress'> {
  images?: ImageProvider;
  cache?: Map<string, Uint8Array>;
  imageConcurrency?: number;
  onProgress?: (event: GenerateProgress) => void;
  properties?: DeckProperties | false; // what the file says of itself (RenderOptions.properties); set unless false
}
interface EditDeckOptions extends Omit<EditOptions, 'onProgress'> {
  images?: ImageProvider;
  cache?: Map<string, Uint8Array>;
  imageConcurrency?: number;
  onProgress?: (event: GenerateProgress) => void;
  properties?: DeckProperties | false; // what the file says of itself (RenderOptions.properties); set unless false
}
type GenerateProgress = PlanProgress | { stage: 'images'; done: number; total: number } | { stage: 'render' };

interface GenerateResult { pptx: Uint8Array; deck: DeckSpec; assets: Record<string, Uint8Array>; warnings: Issue[]; outline: OutlineEntry[] }
interface EditDeckResult { pptx: Uint8Array; deck: DeckSpec; assets: Record<string, Uint8Array>; warnings: Issue[]; summary: string; rewritten: number[]; recolored?: number[] }

interface ResolveOptions {
  images?: ImageProvider;
  icons?: IconDef[];
  assets?: AssetDef[];
  cache?: Map<string, Uint8Array>;
  concurrency?: number;
  onProgress?: (event: { stage: 'images'; done: number; total: number }) => void;
  signal?: AbortSignal;
}
```

## Free slides and colours

### `freeSlide`, `freeContent`

```ts
function freeSlide(content: FreeSlideContent): SlideSpec;
function freeContent(slide: SlideSpec): FreeSlideContent;

interface FreeSlideContent {
  title?: string;
  blocks: FreeBlockContent[];
  columns?: number;                           // how many blocks stand side by side on one line
  lead?: 'left' | 'right' | 'top' | 'none';   // where the first block stands when it is the main one
  dense?: boolean;                            // a long text a person asked to have whole (FreeArrange.dense); kept both ways
  base?: string;                              // the pattern whose frame the slide stands on
  colors?: DeckColors;                        // colours for this slide alone
}
type FreeBlockContent =
  | { text: string | string[]; heading?: string }
  | { list: string[]; heading?: string }
  | { card: { heading: string; text: string; icon?: ImageRef | string } }
  | { figure: { value: string; label: string } }
  | { step: { heading: string; text: string } }
  | { quote: string; source?: string }
  | { note: string }
  | { chart: ChartValue; caption?: string }
  | { table: TableValue; caption?: string }
  | { photo: ImageRef | string; caption?: string }
  | { graphic: ImageRef | string; caption?: string };
```

`freeSlide` writes a free slide: one the software lays out itself from the blocks given, any number of them,
in reading order. It needs no template and returns the slide as a deck holds it: the kinds in `blocks`, the
content in `slots` under `b1Text`, `b2Chart`, and so on. `freeContent` is its reverse, for code that adds,
removes or reorders blocks. `freeSlide` throws an `Error` for a block that is none of the kinds. See
[free-slides.md](free-slides.md).

### `isFreeSlide`, `FREE_PATTERN`, `FREE_BLOCK_KINDS`

```ts
function isFreeSlide(slide: Pick<SlideSpec, 'blocks'>): boolean;
const FREE_PATTERN: 'free';
const FREE_BLOCK_KINDS: readonly ['text', 'list', 'card', 'figure', 'step', 'quote', 'note', 'chart', 'table', 'photo', 'graphic'];
type FreeBlockKind = (typeof FREE_BLOCK_KINDS)[number];
interface FreeArrange { columns?: number; lead?: 'left' | 'right' | 'top' | 'none'; dense?: boolean }
```

A slide is a free slide when it lists `blocks`; its `pattern` is then `FREE_PATTERN`. `FreeArrange` is the type
of `SlideSpec.arrange`. `dense` is set by the planner on a free slide of one text that carries an amount of text
a person asked for, or their own text: where that text is too long for four fifths of the template's size, it
is set smaller, down to the least size that reads (9 pt on a slide 13.33 inches wide, or the template's
`leastTextSize`), in two or three columns
in the most room the frame has. A text that fits is set as large as it fits, as without the flag, down to four
fifths of the template's size. From 0.6.3 one rule says whether such a text fits and what it holds, when it is
planned, written and checked: its paragraphs counted by length one by one, its columns filled line by line as
PowerPoint fills them, each column keeping a line free, so it can be set a little smaller than the same text
without the flag; a text in columns gets a box as high as its columns and a line more. See "A long text a
person asked for" in [free-slides.md](free-slides.md).

### `freeBaseRefusal`

```ts
function freeBaseRefusal(template: Template, base: string): string | undefined;
```

Why a layout of the template cannot be the frame that free slides stand on, as a sentence, or `undefined` when
it can: `the template has no layout "gone"`, `the layout "two-columns" has a row of columns, which the software
lays out itself`, `the layout "cover" has its title low on the slide: ...`. `base` is the id of a pattern, as
`free.base` in a manifest or `base` on a slide names it. A studio asks it before it offers a layout as the
frame. A slide that names a layout that cannot be one stands on the template's usual frame, and the check of the
deck says why. See "What a template gives a free slide" in [free-slides.md](free-slides.md).

### `templatePalette`

```ts
function templatePalette(template: Template): TemplatePalette;

interface TemplatePalette {
  accent?: string;      // the template's main colour: what "accent" in a deck's colours takes the place of
  accent2?: string;     // its second colour
  ink?: string;         // the dark colour its text is set in, when that is not black or a grey
  families: ColorFamily[];                         // the most used first; greys, black and white belong to none
  colors: Array<{ hex: string; share: number }>;   // every colour found, the most used first
}
interface ColorFamily { base: string; members: string[]; share: number }
```

The colours of a template as its patterns use them, read from the pptx with no model. Colours are six hex
digits in upper case. `manifest.colors` states the main and second colour outright. See [colors.md](colors.md).

## Chat tools and storage

### `presentationTools`

```ts
function presentationTools(options: PresentationToolsOptions): PresentationTools;
```

Builds the four tools for a chat model. `call` runs one of the four declared names and answers
`{ output: { error } }` for any other name. An error is told to the chat model as its message, with one
exception: a `TypeError`, `RangeError` or `SyntaxError` of the software itself, which a stored plan or
presentation that is no longer what the tools saved brings about, is told as one sentence that says so. What
the `llm`, the `images` provider, the `store` or the `review` throws is always passed on as its message. See
"Errors" in [chat-tools.md](chat-tools.md).

With `documents: true` the plan tool takes `documents`, the names of texts the person gave in the
conversation. The app passes those texts with every call, as `call(name, args, { signal, documents })`, and
the slides carry the texts the chat model names word for word. From 0.7.1 named texts longer than
`maxDocumentChars` (12000 unless set) are answered with a list of their parts, and the chat model names the
parts the presentation is made from in `document_parts`. A wrong `maxDocumentChars` is a `RangeError` from
`presentationTools` itself. See "Texts the person gave" in [chat-tools.md](chat-tools.md).

From 0.6.4 a result that tells a person something has `issues`: every line of `output.warnings`, in its
order, and a note for each slide the `note` of a plan, or of a change that stops before a cut, speaks of, and
for each text of `not_on_slides`. `formatIssue(issue, language, { deck: result.deck })` gives each line in the
person's language; in English it is the line of `output.warnings`. Such a result's `output.next` also tells
the chat model to say it in the language the user writes in. See "Warnings and notes in the person's
language" in [chat-tools.md](chat-tools.md).

### `describeDocuments`

```ts
function describeDocuments(documents: readonly ChatDocument[], options?: { maxDocumentChars?: number } | null): string;
function describeDocuments(documents: readonly ChatDocument[], index: number, ...rest: unknown[]): string; // as a callback of map: the index is passed over
```

The line for the chat model's system prompt that lists the texts a person gave in the conversation, each with
its words and first words, and says how to pass one to the plan tool. It lists the last ten that have a name
and a text, and returns `''` for none. Give it the list the tools are given with each call.

From 0.7.1 give it the options the tools were built with too, or any object with their `maxDocumentChars`. A
text longer than that (12000 when it is left out) is marked too long to make a presentation of whole. Where any
is, the line ends with two more sentences: they tell the chat model to name such a text all the same, with the
part the person wants in `document_parts` or with none, so that the plan tool answers with the parts to choose
from, and never to ask the person to paste a part of a text they gave. A second argument that is no object,
such as an index from `map`, is left out. A wrong `maxDocumentChars` is a `RangeError`.

### `TOOL_NAMES`

```ts
const TOOL_NAMES: {
  plan: 'plan_presentation';
  editPlan: 'edit_presentation_plan';
  create: 'create_presentation';
  edit: 'edit_presentation';
};
```

### `storeKeys`

```ts
const storeKeys: {
  plan(id: string): string;  // 'plan/<id>.json'
  deck(id: string): string;  // 'deck/<id>.json'
  pptx(id: string): string;  // 'pptx/<id>.pptx'
  asset(id: string): string; // 'asset/<id>'
  preview(id: string, slide: number, extension: string): string; // 'preview/<id>/slide-03.svg'
};
```

### `memoryStore`

```ts
function memoryStore(): BlobStore & { keys(): string[] };
```

### `directoryStore` (from `slotdeck/node`)

```ts
function directoryStore(directory: string): BlobStore;
```

A store backed by a folder. Keys become file paths under it.

### Types

```ts
interface BlobStore {
  get(key: string): Promise<Uint8Array | undefined>;
  set(key: string, data: Uint8Array): Promise<void>;
}

interface PresentationToolsOptions {
  template: Template;
  llm: LlmProvider;
  store: BlobStore;
  images?: ImageProvider;
  icons?: IconDef[];
  assets?: AssetDef[];
  layout?: LayoutMode;
  fields?: Record<string, string>;
  properties?: DeckProperties | false; // what a presentation's file says of itself (RenderOptions.properties); set unless false
  numberPatterns?: 'from-brief' | 'always' | 'never';
  review?: DeckReview; // a second look at every draft the tools write
  composeLayouts?: number; // new layouts for a new plan, as for planOutline; none while free slides are on offer
  free?: boolean; // whether a slide may be a free slide, as `free` of PlanOptions; on unless false
  countSpacing?: boolean; // as `countSpacing` of PlanOptions; on unless false
  patterns?: string[]; // the only layouts the tools may use: the app's own list; a person may narrow it, not widen it
  colors?: boolean; // whether a person may ask for another colour; false leaves accent_color and color_slides out
  thumbnail?: (patternId: string) => string | undefined;
  preview?: PreviewProvider; // draws every presentation the tools create or change
  newId?: (kind: 'plan' | 'deck') => string;
  onProgress?: (event: GenerateProgress, tool: string) => void;
  documents?: boolean; // the plan tool takes the names of texts the person gave; off unless true
  maxDocumentChars?: number; // from 0.7.1: the most characters of the person's texts a presentation is made from, as for planOutline; 12000 unless set, 1000 to 30000; checked when the tools are built
}
interface PresentationTools {
  definitions: ToolDefinition[];
  instructions: string; // from 0.7.0: lines for the chat model's system prompt, made from the same options; see chat-tools.md
  call(name: string, args: unknown, context?: ToolCallContext): Promise<ToolResult>;
}
interface ToolCallContext {
  signal?: AbortSignal;
  documents?: readonly ChatDocument[]; // the texts of the conversation, read by the plan tool of tools built with documents: true
}
interface ChatDocument {
  name: string; // the name the chat model sees and passes: a file name, "pasted text 1"
  text: string; // the text as the person gave it
}
interface ToolDefinition { name: string; description: string; parameters: JsonSchema }
interface ToolResult {
  output: Record<string, unknown>;
  issues?: Issue[]; // from 0.6.4: what the result tells a person, to show with formatIssue; not in output
  plan?: DeckPlan;
  deck?: DeckSpec;
  pptx?: Uint8Array;
  previews?: SlidePreview[];
}
```

## Providers

### `gemini`, `geminiImages`

```ts
function gemini(options: GeminiOptions): LlmProvider;
function geminiImages(options: GeminiImageOptions): ImageProvider;
```

### `vertexEndpoint`, `GEMINI_DEVELOPER_ENDPOINT`

```ts
function vertexEndpoint(project: string, location?: string): string; // location defaults to 'global'
const GEMINI_DEVELOPER_ENDPOINT: string;
```

### `GeminiError`

```ts
class GeminiError extends Error {
  readonly status: number | undefined;
  readonly retryAfterMs: number | undefined; // from 0.7.0: the wait the service asked for
  readonly timedOut: boolean;                // from 0.7.0: no answer within timeoutMs
}
```

### `retryDelay`

```ts
function retryDelay(error: unknown, attempt: number, random?: () => number): number | undefined;
```

From 0.7.0. How long to wait, in milliseconds, before trying a model call again that failed with `error` on try
number `attempt` (from 1), as the Gemini adapters wait; `undefined` where another try is not worth it (a bad
request or key, a spending cap or a daily quota, a stop). See `retries` in [providers.md](providers.md).

### Types

```ts
type JsonSchema = { [key: string]: unknown };
interface JsonRequest {
  system: string;
  prompt: string;
  schema: JsonSchema;
  maxOutputTokens?: number; // a ceiling on the answer's size, reasoning not counted; pass it on as an output limit
  effort?: 'low'; // a short structural request: use light reasoning if the model has a setting for it
  images?: Array<{ mimeType: string; data: Uint8Array }>;
  signal?: AbortSignal;
}
interface LlmProvider { generateJson(request: JsonRequest): Promise<unknown> }

interface ImageRequest { prompt: string; kind: ImageKind; aspectRatio: number; signal?: AbortSignal }
interface ImageProvider { generateImage(request: ImageRequest): Promise<Uint8Array> }

interface GeminiOptions {
  model: string;
  apiKey?: string;
  endpoint?: string;
  headers?: Record<string, string> | (() => Record<string, string> | Promise<Record<string, string>>);
  fetch?: typeof fetch;
  timeoutMs?: number;
  retries?: number;    // from 0.7.0: tries after a passing failure; default 2
  deadlineMs?: number; // from 0.7.0: the most a call takes, its tries and waits together
  thinkingLevel?: 'minimal' | 'low' | 'medium' | 'high';
  maxOutputTokens?: number;
}
interface GeminiImageOptions {
  model: string;
  apiKey?: string;
  endpoint?: string;
  headers?: Record<string, string> | (() => Record<string, string> | Promise<Record<string, string>>);
  fetch?: typeof fetch;
  timeoutMs?: number;
  retries?: number;    // from 0.7.0: tries after a passing failure; default 2
  deadlineMs?: number; // from 0.7.0: the most a call takes, its tries and waits together
  imageSize?: '512' | '1K' | '2K' | '4K';
  style?: string;
}
```

## Template studio

### `draftManifest`

```ts
function draftManifest(pptx: Uint8Array | PptxPackage, options: DraftOptions): Promise<DraftResult>;

interface DraftOptions {
  llm: LlmProvider;
  name?: string;
  slideImage?: (slide: number) => SlidePicture | undefined | Promise<SlidePicture | undefined>;
  slides?: number[];
  variants?: Array<'mirror' | 'text-only'>;   // also add these variants of every pattern that allows them
  rows?: 'fit' | 'drawn';   // default 'fit': a row takes as many items as the slide has room for
  widen?: number;           // above 1, at most 4; default off. Set where it can change a slide
  concurrency?: number;
  onProgress?: (event: { done: number; total: number; slide: number }) => void;
  signal?: AbortSignal;
}

interface DraftResult {
  manifest: TemplateManifest;
  skipped: Array<{ slide: number; reason: string }>;
  issues: ManifestIssue[];
  notes: string[];                    // what the code settled by measurement
  suggestions: ManifestSuggestion[];  // what the slides allow beyond the draft
}
```

Rejects, before any model call, when `rows` is neither `'fit'` nor `'drawn'`, when `widen` is not a number
above 1 and at most 4, and when `variants` names anything but `'mirror'` and `'text-only'`. See
[template-studio.md](template-studio.md).

### `checkManifest`

```ts
function checkManifest(template: Template): ManifestIssue[];
```

### `sampleDeck`

```ts
function sampleDeck(template: Template, options?: SampleOptions): { deck: DeckSpec; assets: Record<string, Uint8Array> };
```

### `changeManifest`

```ts
function changeManifest(template: Template, changes: ManifestChange[]): ChangeResult;
```

### `editManifest`

```ts
function editManifest(template: Template, instruction: string, options: ManifestEditOptions): Promise<ManifestEditResult>;
```

### `prepareManifest`

```ts
function prepareManifest(template: Template, options: PrepareOptions): PreparedManifest;

interface PrepareOptions {
  icons: boolean;                         // whether the app has an icon set for this template
  layout?: LayoutMode;                    // default 'strict': the mode the planner is given; matters with icons: false
  variants?: readonly PatternVariant[];   // default ['mirror', 'text-only']; [] adds none
  widen?: number | false;                 // default 2; above 1, at most 4; given to rows that have no factor; false gives none
  remove?: boolean;                       // default true; false for a manifest that stored decks use
  fields?: Record<string, string>;        // the deck-field values the app already has
}

interface PreparedManifest {
  manifest: TemplateManifest;
  sound: boolean;                  // no error, and the sample deck renders at the fewest and at the most items
  done: string[];                  // what was done, in sentences
  left: string[];                  // what only a person can settle, in sentences
  added: string[];                 // ids the manifest has now and did not have: variants
  removed: string[];               // ids the manifest had and no longer has: patterns that could not be used
  kept: Array<{ pattern: string; slot: string; was: 'error' | 'remove' }>;  // icon slots set to keep the template's pictogram
  fields: Record<string, string>;  // a value for each deck field that has one: the one given, or the text the template shows there
}
```

Synchronous, no model, the template passed in is not touched. Takes what `suggestChanges` offers (the facts,
widening, the variants named), repairs a row whose items hold a slot that is not the row's, keeps the
template's pictograms in the icon slots that would have to be filled when `icons` is `false`, and solves the
issues that have one answer. A pattern that cannot be used is taken out unless `remove` is `false`; then it
stays, nothing is made of it, and `sound` is `false`. Renames nothing; a second run with the same options
changes nothing. Throws when `icons` is not a boolean, and for a `layout`, a `widen` or a variant that does
not exist. See [template-studio.md](template-studio.md).

### `suggestChanges`, `SUGGESTION_KINDS`

```ts
function suggestChanges(template: Template, options?: SuggestOptions): ManifestSuggestion[];
const SUGGESTION_KINDS: readonly ['icon', 'row', 'row-max', 'row-widen', 'row-emphasis', 'variant'];
type SuggestionKind = (typeof SUGGESTION_KINDS)[number];

interface SuggestOptions {
  kinds?: readonly SuggestionKind[];  // default: all
  widen?: number;                     // the factor `row-widen` suggests; default 2
}

interface ManifestSuggestion {
  kind: SuggestionKind;
  pattern: string;            // the pattern it is about
  slide: number;              // that pattern's slide, from 1
  variant?: PatternVariant;   // for `variant`
  slot?: string;              // for `variant` `without`: the text left out
  message: string;            // what is so now, and what the change would do
  changes: ManifestChange[];  // pass to changeManifest
}
```

Synchronous, no model, nothing changed. What a template's slides allow that its manifest does not use yet, each
with the changes that make it. Only suggestions that `changeManifest` accepts as they stand are listed. Throws
for a kind that is not in `SUGGESTION_KINDS`, and, when `row-widen` is asked for, for a `widen` that is not a
number above 1 and at most 4. See [template-studio.md](template-studio.md).

### `patternVariants`, `VARIANTS`

```ts
function patternVariants(template: Template, pattern: string): VariantOption[];
const VARIANTS: readonly ['mirror', 'text-only', 'without'];
type PatternVariant = (typeof VARIANTS)[number];
```

The variants of a pattern that `changeManifest` can derive with `{ op: 'derivePattern' }`, each with whether
the pattern's slide allows it. See [template-studio.md](template-studio.md).

### `composePattern`, `buildPattern`, `designKit`

```ts
function composePattern(template: Template, instruction: string, options: ComposeOptions): Promise<ComposedPattern>;
function buildPattern(template: Template, design: PatternDesign): ComposedPattern;
function designKit(template: Template, inspected?: TemplateInfo, base?: string): DesignKit | undefined;
const DESIGN_GRID: 12;
const DESIGN_ROLES: readonly ['title', 'subtitle', 'text', 'list', 'number', 'caption', 'decoration'];
class DesignError extends Error { readonly problems: string[] }
```

A layout the template does not have, made from its own parts. `composePattern` asks the model for a design on
a grid of `DESIGN_GRID` columns and rows over the content area, with the template's own text styles and
decorations (`designKit` lists them), sends a design with problems back once, and builds it with
`buildPattern`: a new slide after the last one, on the frame of an existing layout, and a pattern for it.
`buildPattern` is the code half alone, for a design the app keeps or a person corrected. Both throw
`DesignError` with every reason when a design cannot be built. The result carries a new pptx: store it in
place of the template's. See "composePattern" in [template-studio.md](template-studio.md). `base` (in the
options, the design and `designKit`) names the layout to stand on; without it, or when that layout has no title
or is a row of items, the code picks the one whose title sits where most titles do.

### `describeManifest`, `describeSlot`

```ts
function describeManifest(template: Template): string;
function describeSlot(slot: SlotDef): string;
```

### `studioView`

```ts
function studioView(template: Template): StudioView;
```

### `inspectTemplate`

```ts
function inspectTemplate(pkg: PptxPackage): TemplateInfo;
```

### `ROLES`

```ts
const ROLES: readonly ['title', 'subtitle', 'text', 'list', 'number', 'caption', 'photo', 'icon', 'graphic', 'chart', 'table', 'decoration', 'fixed'];
type Role = (typeof ROLES)[number];
```

### `PptxPackage`

```ts
class PptxPackage {
  static load(bytes: Uint8Array): PptxPackage;
  clone(): PptxPackage;
  has(partPath: string): boolean;
  paths(): string[];
  read(partPath: string): Uint8Array;
  readText(partPath: string): string;
  write(partPath: string, data: Uint8Array | string): void;
  delete(partPath: string): void;
  save(): Uint8Array;
}
```

The unzipped parts of a pptx. Needed directly only for `inspectTemplate` and for reading a part such as a
picture out of a template (`template.pkg.read('ppt/media/image1.png')`).

### Types

```ts
interface SlidePicture { mimeType: string; data: Uint8Array }

interface ManifestIssue { severity: 'error' | 'warning'; pattern?: string; slot?: string; message: string }
interface SampleOptions { fill?: number; items?: 'drawn' | 'max' | 'min'; fit?: 'budget' | 'box' }

type ManifestChange =
  | { op: 'setManifest'; name?: string; language?: string; direction?: Direction; alignOnDirectionChange?: AlignPolicy; chartDirection?: Direction | null; textScale?: number | null; characterWidth?: number | null; leastTextSize?: number | null; colors?: { accent?: string | null; accent2?: string | null } | false | null; free?: boolean | null | { base?: string | null; styles?: Record<string, { slide: number; shape: string } | null> } }
  | { op: 'addPattern'; slide: number; id: string; description?: string }
  | { op: 'copyPattern'; pattern: string; id: string; description?: string }
  | { op: 'derivePattern'; pattern: string; variant: PatternVariant; slot?: string; id?: string; description?: string }
  | { op: 'removePattern'; pattern: string }
  | { op: 'setPattern'; pattern: string; id?: string; description?: string; figures?: boolean }
  | { op: 'addSlot'; pattern: string; shape: string; role: Role; name?: string; optional?: boolean }
  | { op: 'removeSlot'; pattern: string; slot: string }
  | ({ op: 'setSlot'; pattern: string; slot: string } & SlotPatch)
  | { op: 'setRepeat'; pattern: string; min?: number; max?: number; item?: string; span?: { left: number; right: number } | null; widen?: number | null; emphasis?: number | null; spread?: number | null }
  | { op: 'removeRepeat'; pattern: string };

interface SlotPatch {
  name?: string;
  description?: string | null;
  empty?: EmptyPolicy;
  maxChars?: number;
  maxParagraphs?: number | null;
  deckField?: string | null;
  alignOnDirectionChange?: AlignPolicy | null;
  styleRun?: 'first' | 'last' | null;
  fontSize?: number | null;
  paragraphGap?: number | null;
  scales?: boolean | null;
  box?: SlideBox | null;
  imageKind?: ImageKind;
  fit?: 'cover' | 'contain' | null;
  maxColumns?: number;
  maxRows?: number;
  maxCellChars?: number;
  headerText?: TableTextStyle | null;
  bodyText?: TableTextStyle | null;
  fill?: boolean | null;
  maxCategories?: number;
  maxSeries?: number;
  colors?: string[] | null;
  text?: TableTextStyle | null;
  lineColor?: string | null;
}

interface ChangeResult {
  manifest: TemplateManifest;
  applied: Array<{ index: number; change: ManifestChange; message: string }>;
  rejected: Array<{ index: number; change: ManifestChange; reason: string }>;
  issues: ManifestIssue[];
}
interface VariantOption { variant: PatternVariant; slot?: string; available: boolean; note: string }
interface ManifestEditOptions { llm: LlmProvider; signal?: AbortSignal }
interface ManifestEditResult extends ChangeResult { note: string }

interface StudioView { aspectRatio: number; slides: StudioSlide[] }
interface StudioSlide { slide: number; patterns: string[]; shapes: StudioShape[] }
interface StudioShape {
  id: string;
  kind: ShapeKind;
  box: Box;
  groupId?: string;
  text?: string;
  fontSize?: number;
  frame: boolean;
  roles: Role[];
  slots: Array<{ pattern: string; name: string; kind: 'text' | 'image' | 'table' | 'chart'; empty: EmptyPolicy; summary: string }>;
}

type ShapeKind = 'text' | 'picture' | 'shape' | 'line' | 'group' | 'table' | 'other';
interface Box { x: number; y: number; w: number; h: number } // fractions of the slide
interface ShapeInfo {
  id: string;
  name: string;
  kind: ShapeKind;
  box: Box;
  groupId?: string;
  placeholder?: string;
  paragraphs?: string[];
  fontSizes?: number[];
  image?: { part: string; width?: number; height?: number; cropped: boolean; geometry: string };
  table?: { columns: number; rows: number };
}
interface SlideInfo { number: number; part: string; layout: string; shapes: ShapeInfo[] }
interface TemplateInfo { slideWidth: number; slideHeight: number; defaultFontSize: number; slides: SlideInfo[] }
```

Composing a layout:

```ts
interface ComposeOptions {
  llm: LlmProvider;
  base?: string;                // the layout whose frame and title the new one stands on
  effort?: 'low';               // light reasoning: about 10 s instead of minutes, and enough in the runs measured
  signal?: AbortSignal;
}
interface ComposedPattern {
  pptx: Uint8Array;             // the template with the new slide after the last one
  manifest: TemplateManifest;   // with the new pattern at the end
  pattern: string;
  slide: number;
  summary: string;
  design: PatternDesign;
  warnings: string[];           // what was allowed, such as a part of the slide left empty
}
interface PatternDesign {
  id: string;
  description: string;
  keep: string[];               // slots of the base slide that stay where they are, such as 'title'
  base?: string;                // the layout it stands on; recorded, so that it is built on the same frame again
  elements: DesignElement[];    // in drawing order: a panel before the texts on it
}
interface DesignElement {
  name: string;                 // the slot's name; '' for a decoration
  style: string;                // 'T1'... a text style, 'D1'... a decoration, from designKit
  role: DesignRole;
  optional: boolean;
  sample: string;               // paragraphs separated by line breaks
  column: number; row: number; width: number; height: number;   // grid cells, from 1; column 1 is at the left
}
type DesignRole = (typeof DESIGN_ROLES)[number];
interface DesignKit {
  base: string;                 // the pattern whose slide gives the frame and the title
  baseSlide: number;
  area: Box;                    // the content area the grid covers
  keepable: Array<{ slot: string; role: string; cells: { column: number; row: number; width: number; height: number } }>;
  taken: boolean[][];           // [row][column]: cells the frame takes
  styles: KitStyle[];
}
interface KitStyle {
  key: string;
  kind: 'text' | 'decoration';
  slide: number;
  shapeId: string;
  text: string;                 // one line about it, as the model reads it
  drawn: Box;
  size?: number;
  lineRows?: number;            // grid rows one line takes
  bullet?: boolean;
  shape?: 'across' | 'down' | 'area';
}
```

## Design checks and reviews

Explained in [design-checks.md](design-checks.md). Both need a template loaded with fonts.

### `checkDesign`

```ts
function checkDesign(template: Template, deck: DeckSpec, options?: DesignOptions): DesignIssue[];
```

Lists design problems that follow from measurement alone: text that runs over its box, boxes and cards left
mostly empty, headings that end in a lone word. Synchronous, no model, no picture.

### `designReview`

```ts
function designReview(template: Template, options?: DesignOptions): DeckReview;
```

The design checks as a review for the planner's `review` option. Reports `empty`, `uneven` and `orphan`.

### `visionReview`

```ts
function visionReview(options: VisionReviewOptions): DeckReview;
```

A review by a model that looks at pictures of the draft's slides, drawn by a `PreviewProvider` that gives
bitmap pictures. Not run against a live model; see the cautions in [design-checks.md](design-checks.md).

### Types

```ts
type DesignKind = 'overflow' | 'empty' | 'uneven' | 'orphan';
interface DesignIssue { severity: 'warning'; slide: number; slot: string; kind: DesignKind; message: string } // an Issue with a kind
interface DesignOptions {
  layout?: LayoutMode;
  kinds?: DesignKind[];  // default: all four
  emptyBelow?: number;   // default 1/3
  headingSize?: number;  // points; default 20
}

// A second look at a draft deck: problems to send back to the writing model.
type DeckReview = (deck: DeckSpec, context: { signal?: AbortSignal }) => Issue[] | Promise<Issue[]>;

interface VisionReviewOptions {
  template: Template;
  llm: LlmProvider;         // a model that accepts pictures
  preview: PreviewProvider; // must give PNG, JPEG, WebP or GIF
  icons?: IconDef[];
  assets?: AssetDef[];
  layout?: LayoutMode;       // default 'adaptive'
  slidesPerRequest?: number; // default 6
}
```

## Slide previews

Explained in [preview.md](preview.md). slotdeck draws nothing itself: a `PreviewProvider` is a renderer you plug in.

### `pagusPreviews`

```ts
function pagusPreviews(core: PagusCore, renderer: PagusRenderer, options?: PagusOptions): PreviewProvider;

interface PagusCore { parse(data: ArrayBuffer): Promise<{ slides: readonly unknown[]; slideSize: unknown }> } // the `@pagus-kit/core` module
interface PagusRenderer { renderSlide(slide: never, slideSize: never, options?: never): { svg: string; width: number; height: number } } // the `@pagus-kit/renderer` module
interface PagusOptions { scale?: number; fontSubstitutes?: Record<string, string>; fontCss?: string }
```

A provider that draws slides with the Pagus library, which you install and pass in:
`pagusPreviews(pagusCore, pagusRenderer)` after `import * as pagusCore from '@pagus-kit/core'` and
`import * as pagusRenderer from '@pagus-kit/renderer'`. One SVG for each slide, with the text as real text for a
browser to set. Paragraphs that start with a right-to-left letter are marked as running right to left.

### `previewDeck`

```ts
function previewDeck(template: Template, deck: DeckSpec, preview: PreviewProvider, options?: PreviewDeckOptions): Promise<{ pptx: Uint8Array; previews: SlidePreview[] }>;
```

`renderDeck`, then the provider: the pptx and a picture of each of its slides.

### `patternPreviews`

```ts
function patternPreviews(template: Template, preview: PreviewProvider, options?: SampleOptions & { signal?: AbortSignal }): Promise<Record<string, SlidePreview>>;
```

A picture of every pattern, by pattern id: each pattern's slide of the template's sample deck.

### `previewExtension`

```ts
function previewExtension(preview: { mimeType: string }): string; // 'svg', 'png', 'jpg', ...
```

The file extension for a preview's type, for naming it in a store or on disk.

### Types

```ts
// What slotdeck needs from a slide renderer. Implement it to plug in any renderer.
interface PreviewProvider {
  renderSlides(request: PreviewRequest): Promise<SlidePreview[]>; // one picture for each slide asked for, in slide order
}
interface PreviewRequest {
  pptx: Uint8Array;
  slides?: number[]; // 1-based; every slide when left out
  direction?: Direction; // the direction the deck reads, for a renderer that sets paragraph direction itself; slotdeck's own calls fill it
  signal?: AbortSignal;
}
interface SlidePreview {
  slide: number;    // 1-based
  mimeType: string; // 'image/svg+xml', 'image/png'
  data: Uint8Array; // SVG as UTF-8 text, or the bitmap's bytes
  width?: number;   // pixels
  height?: number;
}
interface PreviewDeckOptions { assets?: Record<string, Uint8Array>; layout?: LayoutMode; signal?: AbortSignal }
```

## Measuring text

Explained in [text-fit.md](text-fit.md). All of it is optional: without fonts nothing here is used.

### `fontkitMetrics`

```ts
function fontkitMetrics(fontkit: FontkitModule, files: FontFiles): FontMetrics;

type FontFiles = Record<string, Uint8Array | Uint8Array[]>; // typeface name as the template states it -> its font files
interface FontkitModule { create(buffer: never, postscriptName?: string): unknown } // the `fontkit` module itself
```

Font metrics from font files, read with the `fontkit` library, which you install and pass in:
`fontkitMetrics(fontkit, { Calibri: [regularBytes, boldBytes] })` after `import * as fontkit from 'fontkit'`.
Reads TrueType, OpenType, collections, WOFF and WOFF2.

### `fontFiles` (from `slotdeck/node`)

```ts
function fontFiles(fontkit: FontkitModule, directory: string, typefaces: Iterable<string>): FontFiles;
```

Finds the font files of the given typefaces in a folder, such as the system's font folder, by the family name
each file states. Typefaces with no file there are left out.

### `templateFonts`

```ts
function templateFonts(template: Template, language?: string): TemplateFont[];
```

The typefaces a template's text slots are set in, with the weights and slants used: the font files to supply.
`language` (default: the manifest's) adds the typeface that text of that language is set in.

### `textMeasurer`

```ts
function textMeasurer(template: Template, fonts?: FontMetrics): TextMeasurer;
```

A measurer for a template. `fonts` defaults to the fonts the template was loaded with; it throws when there
are none. Pass the result to `validateDeck` as `measure`, or call `fit` yourself.

### `measuredBudgets`

```ts
function measuredBudgets(template: Template, options?: BudgetOptions): MeasuredBudget[];
```

How many characters each text slot's box holds, as measured, with a budget to set. The template must have
fonts.

### Types

```ts
// What slotdeck needs from a font library. Implement it to use a library other than fontkit.
interface FontMetrics {
  face(typeface: string, style: FontStyle): FontFace | undefined; // undefined: the typeface was not supplied
  fallback?(codePoint: number, style: FontStyle): FontFace | undefined; // a face for a character the chosen one lacks
}
interface FontStyle { bold: boolean; italic: boolean }
interface FontFace {
  has(codePoint: number): boolean;
  advances(text: string, kerning: boolean): number[]; // the advance of each glyph after shaping, in ems
  substitute?: boolean; // stands in for a weight or slant that was not supplied
}

interface TemplateFont { typeface: string; bold: boolean; italic: boolean }

interface TextMeasurer {
  fit(pattern: string | PatternDef, slot: string | TextSlotDef, text: string | string[], options?: FitOptions): TextFit | undefined;
  missingFonts(): TemplateFont[];
}
interface FitOptions {
  textScale?: number;   // the factor on body text the deck is rendered with; default 1
  shrink?: ShrinkLimit; // allow the text to be set smaller when it does not fit
  items?: number;       // for a slot of a row of items: how many items the slide has
  direction?: Direction; // the direction the deck reads; decides how a paragraph that mixes scripts is written and measured; the template's when left out
  bold?: number[];       // paragraphs drawn in bold, by position: a slide's document headings (SlideSpec.headings), measured in the bold face; one over a paragraph of text is measured run in at its head, where the renderer sets it so
}
interface ShrinkLimit { to: number; minSize?: number } // the shape of the manifest's shrinkText
interface TextFit {
  fits: boolean;
  lines: number;       // lines the text takes
  capacity: number;    // lines of this text the box holds
  overflow: number;    // lines that do not fit
  height: number;      // points the text takes
  room: number;        // points the box gives
  widthRatio: number;  // for a box that does not wrap: longest line / box width
  size: number;        // font size of the first paragraph, in points
  sizes: number[];     // font size of each paragraph
  scale: number;       // the factor the text was shrunk by; 1 when it was not
  fitsChars: number;   // about how many characters of such text the box holds
  paragraphs: string[][]; // the lines of each paragraph; none for a heading measured run in, whose words open the first line of the next
  widths: number[][];     // the width of each of those lines, in points
  lineRoom: number[];     // the width a line of each paragraph may take, in points
  sampleLines: number;    // lines the template's own sample text takes in the slot; 0 when there is none to compare with
  approximate: boolean;   // a typeface, weight or outline was stood in for
}

interface BudgetOptions { sample?: string; margin?: number } // margin defaults to 0.9
interface MeasuredBudget { pattern: string; slot: string; maxChars?: number; fits: number; suggested: number }
```

## Manifest types

Described field by field in [manifest-format.md](manifest-format.md).

```ts
type Direction = 'rtl' | 'ltr';
type AlignPolicy = 'flip' | 'keep';
type EmptyPolicy = 'error' | 'remove' | 'keep' | 'drop';
type ImageKind = 'photo' | 'icon' | 'graphic';
interface SlideBox { x: number; y: number; w: number; h: number }
interface TableTextStyle { size?: number; bold?: boolean; color?: string; font?: string }

interface TemplateManifest {
  version: 1;
  name: string;
  language?: string;
  direction: Direction;
  alignOnDirectionChange?: AlignPolicy;
  chartDirection?: Direction; // 'rtl' draws the charts of a deck that reads right to left from the right
  textScale?: number; // enlarges body text in every deck made from the template
  characterWidth?: number; // how wide PowerPoint sets a character of its running text, as a share of its size, measured (0.3 to 0.8); free slides count their running text with it, and from 0.6.8 a layout's slots of paragraphs what their writers are told
  leastTextSize?: number; // the least size, in points, of a long text set dense on a free slide (5 to 24, up to a half point); left out, by the slide's width
  shrinkText?: { to: number; minSize?: number }; // lets text that does not fit be set smaller; needs fonts
  colors?: { accent?: string; accent2?: string } | false; // the main and second colour; false keeps every deck in the template's own
  free?: { base?: string; styles?: Record<string, { slide: number; shape: string }> } | false; // free slides; false stops them being offered
  patterns: PatternDef[];
}
interface PatternDef {
  id: string;
  sourceSlide: number;
  description: string;
  slots: SlotDef[];
  figures?: boolean;
  sequences?: string[][][];
  repeat?: RepeatDef;
}
interface RepeatDef {
  item: string;
  slots: string[];
  shapes: string[][];
  min?: number;
  max?: number;
  span?: { left: number; right: number };
  columns?: number;
  widen?: number; // fewer items on a line than drawn grow up to this factor to fill the row
  emphasis?: number; // the item the slide draws to stand out; a deck that reads the other way moves the items with their looks
  spread?: number; // fewer lines than drawn stand up to this factor further apart, within the area they are drawn in
}
type SlotDef = TextSlotDef | ImageSlotDef | TableSlotDef | ChartSlotDef;

// Every slot kind has: name, shapeId, description?, empty?
interface TextSlotDef {
  kind: 'text';
  name: string;
  shapeId: string;
  description?: string;
  empty?: EmptyPolicy;
  maxChars?: number;
  maxParagraphs?: number;
  deckField?: string;
  alignOnDirectionChange?: AlignPolicy;
  styleRun?: 'first' | 'last';
  size?: number;
  fontSize?: number;
  box?: SlideBox;
  paragraphGap?: number;
  scales?: boolean;
}
interface ImageSlotDef {
  kind: 'image';
  name: string;
  shapeId: string;
  description?: string;
  empty?: EmptyPolicy;
  imageKind: ImageKind;
  fit?: 'cover' | 'contain';
  box?: SlideBox;
  copyOf?: string; // set by the engine on slots of copied items
}
interface TableSlotDef {
  kind: 'table';
  name: string;
  shapeId: string;
  description?: string;
  empty?: EmptyPolicy;
  maxColumns?: number;
  maxRows?: number;
  maxCellChars?: number;
  headerText?: TableTextStyle;
  bodyText?: TableTextStyle;
  fill?: boolean; // with fewer rows than the sample, the table still fills its frame; numbers end on the right
}
interface ChartSlotDef {
  kind: 'chart';
  name: string;
  shapeId: string;
  description?: string;
  empty?: EmptyPolicy;
  box?: SlideBox;
  maxCategories?: number;
  maxSeries?: number;
  colors?: string[];
  text?: TableTextStyle;
  lineColor?: string;
  detail?: 'full' | 'compact' | 'spark'; // how much a chart shows besides its data; default 'full'
  drawn?: ChartDrawn; // set by the software on the charts of a free slide; a manifest does not state it
}
```

`ChartDrawn` is not exported: it is what the software works out for a chart whose place on the slide is known
(its frame, the colour it stands on, the end of its value axis, whether it draws numbers, a value axis and a
legend, and how its numbers are written). A chart slot without it is drawn as it always was.
