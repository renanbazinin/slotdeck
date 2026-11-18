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

interface RenderOptions { assets?: Record<string, Uint8Array>; layout?: LayoutMode }
interface RenderResult {
  pptx: Uint8Array;
  warnings: Issue[];
  shrunk: Array<{ slide: number; slot: string; size: number; scale: number }>; // texts set smaller to fit; empty without shrinkText
}

type LayoutMode = 'strict' | 'adaptive';
interface ValidateOptions { layout?: LayoutMode; assets?: Iterable<string>; measure?: TextMeasurer }
interface Issue { severity: 'error' | 'warning'; slide?: number; slot?: string; message: string }

interface DeckSpec {
  version: 1;
  title?: string;
  brief?: string;
  language?: string;
  direction?: Direction;
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
  series: Array<{ name: string; values: number[] }>;
  format?: 'number' | 'percent';
}
```

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
function planTextFit(template: Template, plan: DeckPlan, options?: Pick<PlanOptions, 'patterns' | 'free' | 'icons' | 'assets' | 'layout' | 'numberPatterns'>): SlideTextFit[];
```

What each slide of a plan that has asked words or passages of a document will carry, worked out by code
alone before anything is written: the words it holds, the size its running text is set at on a free slide, and
whether it is set dense. A slide that holds less than its words is cut to what it holds when the deck is
written. `[]` for a plan with neither.

### `documentPassages`, `MOST_DOCUMENT_CHARS`

```ts
function documentPassages(document: string): DocumentPassage[];
const MOST_DOCUMENT_CHARS = 12000;
```

The passages of a document (`Brief.document`), numbered as the planner numbers them: each paragraph, a long one
cut at the ends of its sentences, short lines of one section together, a list, a table. A short line on its own,
or one that starts with `#`, is a heading over the passages after it; a run of three short lines or more with no
stops, one under another, is text, all of it. A tab at the start of a line, as text pasted from a word processor
has it, sets the line in and makes no table. Pure: the same text always gives the same passages. A brief's
document is at most `MOST_DOCUMENT_CHARS` characters; a longer one is a `PlanError`.

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
  document?: string; // a document to make the deck from, word for word; at most MOST_DOCUMENT_CHARS characters
}
type ImageDensity = 'none' | 'low' | 'medium' | 'high';

interface PlanOptions {
  llm: LlmProvider;
  layout?: LayoutMode;
  icons?: IconDef[];
  assets?: AssetDef[];
  patterns?: string[];
  numberPatterns?: 'from-brief' | 'always' | 'never';
  maxRepairs?: number;
  review?: DeckReview; // a second look at every draft; see "Design checks and reviews"
  composeLayouts?: number; // up to this many new layouts for the deck, in the template's own look; off (0) by default
  free?: boolean; // whether a slide may be a free slide, laid out by the software from the parts the plan names; on unless false
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
}
interface EditDeckOptions extends Omit<EditOptions, 'onProgress'> {
  images?: ImageProvider;
  cache?: Map<string, Uint8Array>;
  imageConcurrency?: number;
  onProgress?: (event: GenerateProgress) => void;
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
is set smaller, down to the least size that reads (9 pt on a slide 13.33 inches wide), in two or three columns
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
the slides carry the texts the chat model names word for word. See "Texts the person gave" in
[chat-tools.md](chat-tools.md).

### `describeDocuments`

```ts
function describeDocuments(documents: readonly ChatDocument[]): string;
```

The line for the chat model's system prompt that lists the texts a person gave in the conversation, each with
its words and first words, and says how to pass one to the plan tool. It lists the last ten that have a name
and a text, and returns `''` for none. Give it the list the tools are given with each call.

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
  numberPatterns?: 'from-brief' | 'always' | 'never';
  review?: DeckReview; // a second look at every draft the tools write
  composeLayouts?: number; // new layouts for a new plan, as for planOutline; none while free slides are on offer
  free?: boolean; // whether a slide may be a free slide, as `free` of PlanOptions; on unless false
  patterns?: string[]; // the only layouts the tools may use: the app's own list; a person may narrow it, not widen it
  colors?: boolean; // whether a person may ask for another colour; false leaves accent_color and color_slides out
  thumbnail?: (patternId: string) => string | undefined;
  preview?: PreviewProvider; // draws every presentation the tools create or change
  newId?: (kind: 'plan' | 'deck') => string;
  onProgress?: (event: GenerateProgress, tool: string) => void;
  documents?: boolean; // the plan tool takes the names of texts the person gave; off unless true
}
interface PresentationTools {
  definitions: ToolDefinition[];
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
interface ToolResult { output: Record<string, unknown>; plan?: DeckPlan; deck?: DeckSpec; pptx?: Uint8Array; previews?: SlidePreview[] }
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
}
```

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
const SUGGESTION_KINDS: readonly ['icon', 'row', 'row-max', 'row-widen', 'variant'];
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
  | { op: 'setManifest'; name?: string; language?: string; direction?: Direction; alignOnDirectionChange?: AlignPolicy; textScale?: number | null }
  | { op: 'addPattern'; slide: number; id: string; description?: string }
  | { op: 'copyPattern'; pattern: string; id: string; description?: string }
  | { op: 'derivePattern'; pattern: string; variant: PatternVariant; slot?: string; id?: string; description?: string }
  | { op: 'removePattern'; pattern: string }
  | { op: 'setPattern'; pattern: string; id?: string; description?: string; figures?: boolean }
  | { op: 'addSlot'; pattern: string; shape: string; role: Role; name?: string; optional?: boolean }
  | { op: 'removeSlot'; pattern: string; slot: string }
  | ({ op: 'setSlot'; pattern: string; slot: string } & SlotPatch)
  | { op: 'setRepeat'; pattern: string; min?: number; max?: number; item?: string; span?: { left: number; right: number } | null; widen?: number | null }
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
  paragraphs: string[][]; // the lines of each paragraph
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
  textScale?: number; // enlarges body text in every deck made from the template
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
