// The four tools a chat model needs to make presentations with a person in the loop:
// plan, change the plan, create the deck, change the deck.
import { isImageRef } from './deck/spec.ts';
import type { DeckSpec } from './deck/spec.ts';
import type { Issue, LayoutMode } from './deck/validate.ts';
import type { ImageProvider, JsonSchema, LlmProvider } from './llm/provider.ts';
import { resolveAssets } from './plan/assets.ts';
import type { ResolveOptions } from './plan/assets.ts';
import type { AssetDef, IconDef, ImageDensity } from './plan/catalog.ts';
import { planEdit } from './plan/edit.ts';
import { editPlan, planOutline, writeDeck } from './plan/plan.ts';
import type { Brief, DeckPlan, PlanOptions } from './plan/plan.ts';
import { previewExtension } from './preview/provider.ts';
import type { PreviewProvider, SlidePreview } from './preview/provider.ts';
import { renderDeck } from './render/render.ts';
import type { Template } from './render/render.ts';
import { readJson, writeJson } from './store.ts';
import type { BlobStore } from './store.ts';
import type { GenerateProgress } from './generate.ts';

export interface ToolDefinition {
  name: string;
  description: string;
  /** JSON Schema of the arguments. */
  parameters: JsonSchema;
}

export interface PresentationToolsOptions {
  template: Template;
  llm: LlmProvider;
  /** Where plans, decks, pictures and files are kept between calls. */
  store: BlobStore;
  /** Generates photos. Without it, decks are planned with no photos. */
  images?: ImageProvider;
  icons?: IconDef[];
  assets?: AssetDef[];
  layout?: LayoutMode;
  /** Deck-level values every deck gets, e.g. the unit shown in the header. */
  fields?: Record<string, string>;
  numberPatterns?: PlanOptions['numberPatterns'];
  /** A second look at every draft the tools write, as for `planDeck`. Off unless set. */
  review?: PlanOptions['review'];
  /** Address of a picture of a pattern's layout, for showing a plan visually. */
  thumbnail?: (patternId: string) => string | undefined;
  /**
   * Draws the slides of every presentation the tools create or change. The pictures are kept in
   * the store under `preview/<presentation id>/`, and returned to the app with the result.
   */
  preview?: PreviewProvider;
  /** Makes ids for plans and decks. Defaults to a short random id. */
  newId?: (kind: 'plan' | 'deck') => string;
  onProgress?: (event: GenerateProgress, tool: string) => void;
}

export interface ToolResult {
  /** What to hand back to the chat model: small, and free of binary data. */
  output: Record<string, unknown>;
  /** The same result for the app itself, to draw a preview or offer a download. */
  plan?: DeckPlan;
  deck?: DeckSpec;
  pptx?: Uint8Array;
  /** Pictures of the slides, when the tools were given a preview provider. */
  previews?: SlidePreview[];
}

export interface PresentationTools {
  /** Tool declarations to give the chat model. */
  definitions: ToolDefinition[];
  /** Runs a tool the chat model called. Expected failures come back as `{ error }` for the model to act on. */
  call(name: string, args: unknown, context?: { signal?: AbortSignal }): Promise<ToolResult>;
}

export const TOOL_NAMES = {
  plan: 'plan_presentation',
  editPlan: 'edit_presentation_plan',
  create: 'create_presentation',
  edit: 'edit_presentation',
} as const;

/** Store keys, so an app can read what the tools saved. */
export const storeKeys = {
  plan: (id: string) => `plan/${id}.json`,
  deck: (id: string) => `deck/${id}.json`,
  pptx: (id: string) => `pptx/${id}.pptx`,
  asset: (id: string) => `asset/${id}`,
  /** A slide's picture: `extension` is `svg` or `png`, as the preview provider made it. */
  preview: (id: string, slide: number, extension: string) => `preview/${id}/slide-${String(slide).padStart(2, '0')}.${extension}`,
};

const definitions: ToolDefinition[] = [
  {
    name: TOOL_NAMES.plan,
    description:
      'Plans a presentation: decides the slides, the layout of each and the point each makes. Nothing is written or rendered yet. ' +
      'Show the plan to the user and get their approval or changes before creating the presentation.',
    parameters: {
      type: 'object',
      properties: {
        topic: {
          type: 'string',
          description: 'What the presentation is about, with every fact, figure and name the user gave. Numbers not stated here will not appear on slides.',
        },
        slide_count: { type: 'integer', minimum: 1, maximum: 30, description: 'Number of slides, when the user asked for one' },
        language: { type: 'string', description: 'BCP 47 tag of the language to write in, e.g. "he" or "en". Defaults to the template language.' },
        image_density: { type: 'string', enum: ['none', 'low', 'medium', 'high'], description: 'How many slides carry a photo. Defaults to medium.' },
        text_size: { type: 'string', enum: ['normal', 'large', 'x-large'], description: 'Size of body text. Larger text means less of it per slide. Defaults to normal.' },
        text_only: {
          type: 'boolean',
          description:
            'True when the user asks for a presentation of text only. Every slide is then plain text: a title with paragraphs or a list, with no pictures, icons, cards, figures, tables or charts.',
        },
        instructions: { type: 'string', description: 'Audience, tone, and anything else to respect' },
      },
      required: ['topic'],
    },
  },
  {
    name: TOOL_NAMES.editPlan,
    description:
      'Changes a plan the user has not approved yet: add, remove, reorder or re-purpose slides, change the length or the language. Returns a new plan with a new id.',
    parameters: {
      type: 'object',
      properties: {
        plan_id: { type: 'string', description: 'Id of the plan to change' },
        instruction: { type: 'string', description: "What to change, in the user's words, with any new facts or figures they gave" },
      },
      required: ['plan_id', 'instruction'],
    },
  },
  {
    name: TOOL_NAMES.create,
    description: 'Writes and renders the presentation of an approved plan, as a PowerPoint file. Call it only after the user approved the plan.',
    parameters: {
      type: 'object',
      properties: { plan_id: { type: 'string', description: 'Id of the approved plan' } },
      required: ['plan_id'],
    },
  },
  {
    name: TOOL_NAMES.edit,
    description:
      'Changes a presentation that was already created: rewrite or shorten a slide, add or remove slides, change the number of cards on a slide, ' +
      'make the text larger or smaller, change a picture, translate. ' +
      'Only the slides the change touches are rewritten. Returns a new presentation with a new id.',
    parameters: {
      type: 'object',
      properties: {
        presentation_id: { type: 'string', description: 'Id of the presentation to change' },
        instruction: { type: 'string', description: "What to change, in the user's words, with any new facts or figures they gave" },
      },
      required: ['presentation_id', 'instruction'],
    },
  },
];

function randomId(kind: 'plan' | 'deck'): string {
  return `${kind}_${crypto.randomUUID().replace(/-/g, '').slice(0, 10)}`;
}

/** A short line per slide for the chat model: enough to talk about the deck, not the whole content. */
function slideLine(deck: DeckSpec, index: number): Record<string, unknown> {
  const slide = deck.slides[index] as DeckSpec['slides'][number];
  const title = slide.slots.title;
  const line: Record<string, unknown> = { number: index + 1, layout: slide.pattern };
  if (typeof title === 'string') line.title = title;
  return line;
}

/**
 * Builds the presentation tools for a chat model. Give `definitions` to the model and pass its
 * tool calls to `call`. Plans, decks and pictures are kept in the store under ids, so the model
 * only ever handles ids and short summaries, never deck JSON or files.
 */
export function presentationTools(options: PresentationToolsOptions): PresentationTools {
  const { template, store } = options;
  const patterns = new Map(template.manifest.patterns.map((pattern) => [pattern.id, pattern]));
  const newId = options.newId ?? randomId;
  const canMakePhotos = !!options.images || (options.assets ?? []).some((asset) => asset.kind === 'photo');

  const planOptions = (tool: string, signal: AbortSignal | undefined): PlanOptions => {
    const plan: PlanOptions = { llm: options.llm };
    if (options.layout) plan.layout = options.layout;
    if (options.icons) plan.icons = options.icons;
    if (options.assets) plan.assets = options.assets;
    if (options.numberPatterns) plan.numberPatterns = options.numberPatterns;
    if (options.review) plan.review = options.review;
    if (options.onProgress) plan.onProgress = (event) => options.onProgress?.(event, tool);
    if (signal) plan.signal = signal;
    return plan;
  };

  const describePlan = (id: string, plan: DeckPlan): Record<string, unknown> => ({
    plan_id: id,
    title: plan.title,
    language: plan.brief.language ?? template.manifest.language,
    slides: plan.slides.map((slide, index) => {
      const line: Record<string, unknown> = {
        number: index + 1,
        layout: slide.pattern,
        layout_description: patterns.get(slide.pattern)?.description,
        purpose: slide.purpose,
      };
      const thumbnail = options.thumbnail?.(slide.pattern);
      if (thumbnail) line.thumbnail = thumbnail;
      return line;
    }),
    next: `Show this plan to the user. Use ${TOOL_NAMES.editPlan} for changes and ${TOOL_NAMES.create} once they approve.`,
  });

  /** Gathers the deck's pictures (stored ones first), renders it, and saves everything under a new id. */
  const finish = async (tool: string, deck: DeckSpec, warnings: Issue[], signal: AbortSignal | undefined) => {
    const supplied = new Set([...(options.icons ?? []), ...(options.assets ?? [])].map((item) => item.id));
    const wanted = new Set<string>();
    for (const slide of deck.slides) {
      for (const value of Object.values(slide.slots)) if (isImageRef(value) && !supplied.has(value.asset)) wanted.add(value.asset);
    }
    const cache = new Map<string, Uint8Array>();
    await Promise.all(
      [...wanted].map(async (asset) => {
        const data = await store.get(storeKeys.asset(asset));
        if (data) cache.set(asset, data);
      }),
    );
    const stored = new Set(cache.keys());

    const resolve: ResolveOptions = { cache };
    if (options.images) resolve.images = options.images;
    if (options.icons) resolve.icons = options.icons;
    if (options.assets) resolve.assets = options.assets;
    if (options.onProgress) resolve.onProgress = (event) => options.onProgress?.(event, tool);
    if (signal) resolve.signal = signal;
    const assets = await resolveAssets(template, deck, resolve);
    // New pictures go to the store on their own, apart from the JSON and the file.
    await Promise.all([...cache].filter(([asset]) => !stored.has(asset)).map(([asset, data]) => store.set(storeKeys.asset(asset), data)));

    options.onProgress?.({ stage: 'render' }, tool);
    const { pptx } = renderDeck(template, deck, options.layout ? { assets, layout: options.layout } : { assets });
    const id = newId('deck');
    await writeJson(store, storeKeys.deck(id), deck);
    await store.set(storeKeys.pptx(id), pptx);

    const output: Record<string, unknown> = {
      presentation_id: id,
      title: deck.title,
      file: storeKeys.pptx(id),
      slides: deck.slides.map((_, index) => slideLine(deck, index)),
    };
    if (warnings.length > 0) {
      output.warnings = warnings.map((issue) => `slide ${issue.slide}${issue.slot ? ` ${issue.slot}` : ''}: ${issue.message}`);
    }

    // Pictures of the slides are an extra: the presentation stands when they cannot be made.
    let previews: SlidePreview[] | undefined;
    if (options.preview) {
      try {
        previews = await options.preview.renderSlides(signal ? { pptx, signal } : { pptx });
        const keys = previews.map((picture) => storeKeys.preview(id, picture.slide, previewExtension(picture)));
        await Promise.all(previews.map((picture, index) => store.set(keys[index] as string, picture.data)));
        output.previews = keys;
      } catch (error) {
        if (signal?.aborted) throw error;
        previews = undefined;
        output.preview_error = error instanceof Error ? error.message : String(error);
      }
    }
    return { id, output, pptx, previews };
  };

  const text = (args: Record<string, unknown>, name: string): string => {
    const value = args[name];
    if (typeof value !== 'string' || value.trim() === '') throw new ToolInputError(`"${name}" is required`);
    return value.trim();
  };

  const handlers: Record<string, (args: Record<string, unknown>, signal: AbortSignal | undefined) => Promise<ToolResult>> = {
    async [TOOL_NAMES.plan](args, signal) {
      const brief: Brief = { prompt: text(args, 'topic') };
      if (options.fields) brief.fields = options.fields;
      if (typeof args.language === 'string' && args.language.trim()) brief.language = args.language.trim();
      if (typeof args.instructions === 'string' && args.instructions.trim()) brief.instructions = args.instructions.trim();
      const count = Number(args.slide_count);
      if (Number.isInteger(count) && count > 0) brief.slideCount = count;
      const density = args.image_density;
      if (density === 'none' || density === 'low' || density === 'medium' || density === 'high') brief.imageDensity = density as ImageDensity;
      if (!canMakePhotos) brief.imageDensity = 'none';
      const size = args.text_size;
      if (size === 'normal' || size === 'large' || size === 'x-large') brief.textSize = size;
      if (args.text_only === true) brief.textOnly = true;

      const plan = await planOutline(template, brief, planOptions(TOOL_NAMES.plan, signal));
      const id = newId('plan');
      await writeJson(store, storeKeys.plan(id), plan);
      return { output: describePlan(id, plan), plan };
    },

    async [TOOL_NAMES.editPlan](args, signal) {
      const planId = text(args, 'plan_id');
      const current = await readJson<DeckPlan>(store, storeKeys.plan(planId));
      if (!current) throw new ToolInputError(`There is no plan with id "${planId}"`);
      const { plan, summary } = await editPlan(template, current, text(args, 'instruction'), planOptions(TOOL_NAMES.editPlan, signal));
      const id = newId('plan');
      await writeJson(store, storeKeys.plan(id), plan);
      return { output: { ...describePlan(id, plan), summary, previous_plan_id: planId }, plan };
    },

    async [TOOL_NAMES.create](args, signal) {
      const planId = text(args, 'plan_id');
      const plan = await readJson<DeckPlan>(store, storeKeys.plan(planId));
      if (!plan) throw new ToolInputError(`There is no plan with id "${planId}"`);
      const written = await writeDeck(template, plan, planOptions(TOOL_NAMES.create, signal));
      const { output, pptx, previews } = await finish(TOOL_NAMES.create, written.deck, written.warnings, signal);
      return withPreviews({ output: { ...output, plan_id: planId }, deck: written.deck, pptx }, previews);
    },

    async [TOOL_NAMES.edit](args, signal) {
      const deckId = text(args, 'presentation_id');
      const deck = await readJson<DeckSpec>(store, storeKeys.deck(deckId));
      if (!deck) throw new ToolInputError(`There is no presentation with id "${deckId}"`);
      const editOptions = { ...planOptions(TOOL_NAMES.edit, signal), ...(canMakePhotos ? {} : { imageDensity: 'none' as const }) };
      const edited = await planEdit(template, deck, text(args, 'instruction'), editOptions);
      const { output, pptx, previews } = await finish(TOOL_NAMES.edit, edited.deck, edited.warnings, signal);
      return withPreviews(
        {
          output: { ...output, summary: edited.summary, rewritten_slides: edited.rewritten, previous_presentation_id: deckId },
          deck: edited.deck,
          pptx,
        },
        previews,
      );
    },
  };

  return {
    definitions,
    async call(name, args, context) {
      const handler = handlers[name];
      if (!handler) return { output: { error: `Unknown tool "${name}"` } };
      try {
        return await handler((typeof args === 'object' && args !== null ? args : {}) as Record<string, unknown>, context?.signal);
      } catch (error) {
        if (context?.signal?.aborted) throw error;
        // Told to the model in words it can act on: a missing id, a plan it must change, a limit it broke.
        return { output: { error: error instanceof Error ? error.message : String(error) } };
      }
    },
  };
}

function withPreviews(result: ToolResult, previews: SlidePreview[] | undefined): ToolResult {
  if (previews) result.previews = previews;
  return result;
}

class ToolInputError extends Error {}
