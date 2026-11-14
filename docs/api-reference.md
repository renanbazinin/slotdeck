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
function validateDeck(manifest: TemplateManifest, deck: DeckSpec, options?: ValidateOptions): Issue[];
```

Checks a deck against a manifest with the renderer's own rules.

### `extractAssets`

```ts
function extractAssets(template: Template, deck: DeckSpec, pptx: Uint8Array): Record<string, Uint8Array>;
```

Reads a deck's pictures back from a pptx rendered from it, by asset id.

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
  slides: SlideSpec[];
}
interface SlideSpec { pattern: string; items?: number; slots: Record<string, SlotValue> }
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

Decides which pattern each slide uses and the point it makes. One model call.

### `editPlan`

```ts
function editPlan(template: Template, plan: DeckPlan, instruction: string, options: PlanOptions): Promise<{ plan: DeckPlan; summary: string }>;
```

Changes a plan by instruction. One model call.

### `writeDeck`

```ts
function writeDeck(template: Template, plan: DeckPlan, options: PlanOptions): Promise<PlanResult>;
```

Writes the content of a plan's slides. One model call plus repairs.

### `planEdit`

```ts
function planEdit(template: Template, deck: DeckSpec, instruction: string, options: EditOptions): Promise<EditResult>;
```

Applies an instruction to a deck and returns the new deck JSON, without pictures or rendering.

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

Thrown when no usable plan can be made. The message says why.

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
  onProgress?: (event: PlanProgress) => void;
  signal?: AbortSignal;
}
type PlanProgress = { stage: 'outline' } | { stage: 'write'; slides: number } | { stage: 'repair'; attempt: number; issues: number };

interface IconDef { id: string; description: string; data: Uint8Array }
interface AssetDef { id: string; kind: 'photo' | 'graphic'; description: string; data: Uint8Array }

interface OutlineEntry { pattern: string; purpose: string; items?: number }
interface DeckPlan { version: 1; title: string; brief: Brief; slides: OutlineEntry[] }
interface PlanResult { deck: DeckSpec; warnings: Issue[]; outline: OutlineEntry[] }

interface EditOptions extends PlanOptions {
  fields?: Record<string, string>;
  direction?: Direction;
  imageDensity?: ImageDensity;
}
interface EditResult { deck: DeckSpec; summary: string; rewritten: number[]; warnings: Issue[] }

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
interface EditDeckResult { pptx: Uint8Array; deck: DeckSpec; assets: Record<string, Uint8Array>; warnings: Issue[]; summary: string; rewritten: number[] }

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

## Chat tools and storage

### `presentationTools`

```ts
function presentationTools(options: PresentationToolsOptions): PresentationTools;
```

Builds the four tools for a chat model.

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
  thumbnail?: (patternId: string) => string | undefined;
  preview?: PreviewProvider; // draws every presentation the tools create or change
  newId?: (kind: 'plan' | 'deck') => string;
  onProgress?: (event: GenerateProgress, tool: string) => void;
}
interface PresentationTools {
  definitions: ToolDefinition[];
  call(name: string, args: unknown, context?: { signal?: AbortSignal }): Promise<ToolResult>;
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
```

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
interface DraftOptions {
  llm: LlmProvider;
  name?: string;
  slideImage?: (slide: number) => SlidePicture | undefined | Promise<SlidePicture | undefined>;
  slides?: number[];
  concurrency?: number;
  onProgress?: (event: { done: number; total: number; slide: number }) => void;
  signal?: AbortSignal;
}
interface DraftResult { manifest: TemplateManifest; skipped: Array<{ slide: number; reason: string }>; issues: ManifestIssue[] }

interface ManifestIssue { severity: 'error' | 'warning'; pattern?: string; slot?: string; message: string }
interface SampleOptions { fill?: number; items?: 'drawn' | 'max' | 'min'; fit?: 'budget' | 'box' }

type ManifestChange =
  | { op: 'setManifest'; name?: string; language?: string; direction?: Direction; alignOnDirectionChange?: AlignPolicy; textScale?: number | null }
  | { op: 'addPattern'; slide: number; id: string; description?: string }
  | { op: 'copyPattern'; pattern: string; id: string; description?: string }
  | { op: 'removePattern'; pattern: string }
  | { op: 'setPattern'; pattern: string; id?: string; description?: string; figures?: boolean }
  | { op: 'addSlot'; pattern: string; shape: string; role: Role; name?: string; optional?: boolean }
  | { op: 'removeSlot'; pattern: string; slot: string }
  | ({ op: 'setSlot'; pattern: string; slot: string } & SlotPatch)
  | { op: 'setRepeat'; pattern: string; min?: number; max?: number; item?: string; span?: { left: number; right: number } | null }
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
}
```
