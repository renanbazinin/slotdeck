// Brief in, pptx out; and an instruction applied to an existing deck.
export { editDeck, generateDeck } from './generate.ts';
export type { EditDeckOptions, EditDeckResult, GenerateOptions, GenerateProgress, GenerateResult } from './generate.ts';

// Tools for a chat model: plan, change the plan, create, change the deck.
export { TOOL_NAMES, presentationTools, storeKeys } from './tools.ts';
export type { PresentationTools, PresentationToolsOptions, ToolDefinition, ToolResult } from './tools.ts';
export { memoryStore } from './store.ts';
export type { BlobStore } from './store.ts';

// The steps of generateDeck, usable on their own.
export { PlanError, directionOf, editPlan, planDeck, planOutline, writeDeck } from './plan/plan.ts';
export type { Brief, DeckPlan, OutlineEntry, PlanOptions, PlanProgress, PlanResult } from './plan/plan.ts';
export { planEdit } from './plan/edit.ts';
export type { EditOptions, EditResult } from './plan/edit.ts';
export { resolveAssets } from './plan/assets.ts';
export type { ResolveOptions } from './plan/assets.ts';
export { extractAssets } from './render/extract.ts';
export type { AssetDef, IconDef, ImageDensity } from './plan/catalog.ts';
export { isPlainText } from './plan/catalog.ts';
export { loadTemplate, renderDeck } from './render/render.ts';
export type { RenderOptions, RenderResult, Template, TemplateOptions } from './render/render.ts';

// Model providers.
export type { ImageProvider, ImageRequest, JsonRequest, JsonSchema, LlmProvider } from './llm/provider.ts';
export { GEMINI_DEVELOPER_ENDPOINT, GeminiError, gemini, geminiImages, vertexEndpoint } from './llm/gemini.ts';
export type { GeminiImageOptions, GeminiOptions } from './llm/gemini.ts';

// Decks.
export { DeckValidationError, validateDeck } from './deck/validate.ts';
export type { Issue, LayoutMode, ValidateOptions } from './deck/validate.ts';
export type { ChartType, ChartValue, DeckSpec, ImageRef, SlideSpec, SlotValue, TableValue } from './deck/spec.ts';
export { isChartValue, isImageRef, isTableValue } from './deck/spec.ts';
export { slideAlternatives } from './deck/alternatives.ts';
export type { AlternativesOptions, SlideAlternative } from './deck/alternatives.ts';

// Templates.
export { draftManifest } from './template/draft.ts';
export type { DraftOptions, DraftResult, SlidePicture } from './template/draft.ts';
export { checkManifest } from './template/check.ts';
export type { ManifestIssue } from './template/check.ts';
export { sampleDeck } from './template/sample.ts';
export type { SampleOptions } from './template/sample.ts';
export { changeManifest, describeSlot } from './template/change.ts';
export type { ChangeResult, ManifestChange, SlotPatch } from './template/change.ts';
export { SUGGESTION_KINDS, suggestChanges } from './template/suggest.ts';
export type { ManifestSuggestion, SuggestOptions, SuggestionKind } from './template/suggest.ts';
export { VARIANTS, patternVariants } from './template/derive.ts';
export type { PatternVariant, VariantOption } from './template/derive.ts';
export { describeManifest, editManifest } from './template/instruct.ts';
export type { ManifestEditOptions, ManifestEditResult } from './template/instruct.ts';
export { studioView } from './template/view.ts';
export type { StudioShape, StudioSlide, StudioView } from './template/view.ts';
export { ROLES } from './template/measure.ts';
export type { Role } from './template/measure.ts';
export { inspectTemplate } from './template/inspect.ts';
export type { Box, ShapeInfo, ShapeKind, SlideInfo, TemplateInfo } from './template/inspect.ts';
export type {
  AlignPolicy,
  ChartSlotDef,
  Direction,
  EmptyPolicy,
  ImageKind,
  ImageSlotDef,
  PatternDef,
  RepeatDef,
  SlideBox,
  SlotDef,
  TableSlotDef,
  TableTextStyle,
  TemplateManifest,
  TextSize,
  TextSlotDef,
} from './template/manifest.ts';
export { TEXT_SIZES } from './template/manifest.ts';

// Design checks, and a second look at a draft.
export { checkDesign } from './deck/design.ts';
export type { DesignIssue, DesignKind, DesignOptions } from './deck/design.ts';
export { designReview, visionReview } from './plan/review.ts';
export type { DeckReview, VisionReviewOptions } from './plan/review.ts';

// Pictures of slides.
export { previewExtension } from './preview/provider.ts';
export type { PreviewProvider, PreviewRequest, SlidePreview } from './preview/provider.ts';
export { pagusPreviews } from './preview/pagus.ts';
export type { PagusCore, PagusOptions, PagusRenderer } from './preview/pagus.ts';
export { patternPreviews, previewDeck } from './preview/patterns.ts';
export type { PreviewDeckOptions } from './preview/patterns.ts';

// Measuring text with real fonts.
export { fontkitMetrics } from './text/fontkit.ts';
export type { FontFiles, FontkitModule } from './text/fontkit.ts';
export type { FontFace, FontMetrics, FontStyle } from './text/fonts.ts';
export { templateFonts, textMeasurer } from './text/measure.ts';
export type { FitOptions, ShrinkLimit, TemplateFont, TextFit, TextMeasurer } from './text/measure.ts';
export { measuredBudgets } from './template/budgets.ts';
export type { BudgetOptions, MeasuredBudget } from './template/budgets.ts';

export { PptxPackage } from './pptx/package.ts';
