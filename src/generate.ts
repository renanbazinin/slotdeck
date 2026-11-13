import type { Issue } from './deck/validate.ts';
import type { DeckSpec } from './deck/spec.ts';
import type { ImageProvider } from './llm/provider.ts';
import { resolveAssets } from './plan/assets.ts';
import type { ResolveOptions } from './plan/assets.ts';
import { planEdit } from './plan/edit.ts';
import type { EditOptions } from './plan/edit.ts';
import { planDeck } from './plan/plan.ts';
import type { Brief, OutlineEntry, PlanOptions, PlanProgress } from './plan/plan.ts';
import { renderDeck } from './render/render.ts';
import type { Template } from './render/render.ts';

export type GenerateProgress = PlanProgress | { stage: 'images'; done: number; total: number } | { stage: 'render' };

interface PictureOptions {
  /** Generates photos. Without it, and without supplied photos, new slides are planned with no photos. */
  images?: ImageProvider;
  /** Generated pictures by asset id; reuse it across calls to avoid generating the same picture twice. */
  cache?: Map<string, Uint8Array>;
  /** How many pictures to generate at once. Defaults to 3. */
  imageConcurrency?: number;
  onProgress?: (event: GenerateProgress) => void;
}

export interface GenerateOptions extends Omit<PlanOptions, 'onProgress'>, PictureOptions {}

export interface GenerateResult {
  pptx: Uint8Array;
  /** The deck as JSON. Store it: it is what gets edited and re-rendered. */
  deck: DeckSpec;
  /** Every picture the deck uses, by asset id. */
  assets: Record<string, Uint8Array>;
  warnings: Issue[];
  outline: OutlineEntry[];
}

export interface EditDeckOptions extends Omit<EditOptions, 'onProgress'>, PictureOptions {}

export interface EditDeckResult {
  pptx: Uint8Array;
  deck: DeckSpec;
  assets: Record<string, Uint8Array>;
  warnings: Issue[];
  /** One sentence from the model on what it changed. */
  summary: string;
  /** Slide numbers in the new deck that were written or rewritten. */
  rewritten: number[];
}

function canMakePhotos(options: PictureOptions & { assets?: Array<{ kind: string }> }): boolean {
  return !!options.images || (options.assets ?? []).some((asset) => asset.kind === 'photo');
}

async function picturesAndPptx(
  template: Template,
  deck: DeckSpec,
  options: (GenerateOptions | EditDeckOptions) & PictureOptions,
): Promise<{ pptx: Uint8Array; assets: Record<string, Uint8Array> }> {
  const resolve: ResolveOptions = {};
  if (options.images) resolve.images = options.images;
  if (options.icons) resolve.icons = options.icons;
  if (options.assets) resolve.assets = options.assets;
  if (options.cache) resolve.cache = options.cache;
  if (options.imageConcurrency) resolve.concurrency = options.imageConcurrency;
  if (options.onProgress) resolve.onProgress = options.onProgress;
  if (options.signal) resolve.signal = options.signal;
  const assets = await resolveAssets(template, deck, resolve);

  options.onProgress?.({ stage: 'render' });
  const rendered = renderDeck(template, deck, options.layout ? { assets, layout: options.layout } : { assets });
  return { pptx: rendered.pptx, assets };
}

/** Brief in, pptx out: plans the deck with the language model, gathers its pictures and renders it. */
export async function generateDeck(template: Template, brief: Brief, options: GenerateOptions): Promise<GenerateResult> {
  const plan = await planDeck(template, canMakePhotos(options) ? brief : { ...brief, imageDensity: 'none' }, options);
  const output = await picturesAndPptx(template, plan.deck, options);
  return { ...output, deck: plan.deck, warnings: plan.warnings, outline: plan.outline };
}

/**
 * Applies an instruction to an existing deck and renders the result. Pass the pictures of the
 * earlier version in `cache` (see `extractAssets`) so unchanged photos are not generated again.
 */
export async function editDeck(template: Template, deck: DeckSpec, instruction: string, options: EditDeckOptions): Promise<EditDeckResult> {
  const edit = await planEdit(template, deck, instruction, canMakePhotos(options) ? options : { ...options, imageDensity: 'none' });
  const output = await picturesAndPptx(template, edit.deck, options);
  return { ...output, deck: edit.deck, warnings: edit.warnings, summary: edit.summary, rewritten: edit.rewritten };
}
