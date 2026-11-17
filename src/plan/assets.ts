import { isFreeSlide, isImageRef } from '../deck/spec.ts';
import type { DeckSpec } from '../deck/spec.ts';
import type { ImageProvider } from '../llm/provider.ts';
import type { PptxPackage } from '../pptx/package.ts';
import type { Template } from '../render/render.ts';
import { inspectTemplate } from '../template/inspect.ts';
import type { TemplateInfo } from '../template/inspect.ts';
import { freeFor } from '../free/resolve.ts';
import { effectiveSlots, itemRange, slotSources } from '../template/manifest.ts';
import type { ImageKind } from '../template/manifest.ts';
import type { AssetDef, IconDef } from './catalog.ts';

export interface ResolveOptions {
  /** Generates the pictures the deck describes by prompt. */
  images?: ImageProvider;
  icons?: IconDef[];
  assets?: AssetDef[];
  /**
   * Generated pictures by asset id. Pass the same map on later calls and pictures whose
   * description did not change are reused instead of generated again.
   */
  cache?: Map<string, Uint8Array>;
  /** How many pictures to generate at once. Defaults to 3. */
  concurrency?: number;
  onProgress?: (event: { stage: 'images'; done: number; total: number }) => void;
  signal?: AbortSignal;
}

const inspections = new WeakMap<PptxPackage, TemplateInfo>();

function inspect(template: Template): TemplateInfo {
  let info = inspections.get(template.pkg);
  if (!info) {
    info = inspectTemplate(template.pkg);
    inspections.set(template.pkg, info);
  }
  return info;
}

interface Pending {
  asset: string;
  prompt: string;
  kind: ImageKind;
  aspectRatio: number;
}

/**
 * Collects the bytes of every picture a deck refers to: supplied icons and assets as they are,
 * described pictures from the cache or the image provider.
 */
export async function resolveAssets(template: Template, deck: DeckSpec, options: ResolveOptions = {}): Promise<Record<string, Uint8Array>> {
  const supplied = new Map<string, Uint8Array>();
  for (const icon of options.icons ?? []) supplied.set(icon.id, icon.data);
  for (const asset of options.assets ?? []) supplied.set(asset.id, asset.data);
  const cache = options.cache ?? new Map<string, Uint8Array>();

  const patterns = new Map(template.manifest.patterns.map((pattern) => [pattern.id, pattern]));
  const resolved: Record<string, Uint8Array> = {};
  const pending = new Map<string, Pending>();
  const mirrored = (deck.direction ?? template.manifest.direction) !== template.manifest.direction;

  deck.slides.forEach((slide, index) => {
    const free = isFreeSlide(slide) ? freeFor(template, deck, slide) : undefined;
    const pattern = free === undefined ? patterns.get(slide.pattern) : typeof free === 'string' ? undefined : free.plan.pattern;
    if (!pattern) return; // rendering reports unknown patterns
    // So it does a number of items the row does not take. Its slots are not worked out here: a
    // deck is read before it is checked, and the slots of a million items are never built.
    const range = itemRange(pattern);
    if (slide.items !== undefined && range && !(Number.isInteger(slide.items) && slide.items >= range.min && slide.items <= range.max)) return;
    const sources = slotSources(pattern, mirrored, slide.items);
    for (const slot of effectiveSlots(pattern, slide.items)) {
      // The frame a picture is sized for is the one it will actually be placed in.
      // A slide stored without slots asks for no picture; the deck's check says what it lacks.
      const value = slide.slots?.[sources.get(slot.name) ?? slot.name];
      if (slot.kind !== 'image' || !isImageRef(value)) continue;
      const known = supplied.get(value.asset) ?? cache.get(value.asset);
      if (known) {
        resolved[value.asset] = known;
        continue;
      }
      if (pending.has(value.asset)) continue;
      const where = `slide ${index + 1} / ${slot.name}`;
      if (!value.prompt) throw new Error(`${where}: asset "${value.asset}" was not supplied and has no description to generate it from`);
      if (!options.images) throw new Error(`${where}: the deck asks for a generated picture, but no image provider was given`);

      const info = inspect(template);
      // A copied item has no shape of its own in the template; its frame is that of the item it copies.
      const frameId = slot.copyOf ?? slot.shapeId;
      const shape = info.slides[pattern.sourceSlide - 1]?.shapes.find((candidate) => candidate.id === frameId);
      // A slot that places its picture says the frame itself; a slide the software laid out always does.
      const frame = slot.box ?? shape?.box;
      const width = (frame?.w ?? 0) * info.slideWidth;
      const height = (frame?.h ?? 0) * info.slideHeight;
      pending.set(value.asset, {
        asset: value.asset,
        prompt: value.prompt,
        kind: slot.imageKind,
        aspectRatio: width > 0 && height > 0 ? width / height : 1,
      });
    }
  });

  const queue = [...pending.values()];
  const total = queue.length;
  let done = 0;
  const images = options.images;
  const worker = async (): Promise<void> => {
    for (let job = queue.shift(); job && images; job = queue.shift()) {
      const request = { prompt: job.prompt, kind: job.kind, aspectRatio: job.aspectRatio };
      const data = await images.generateImage(options.signal ? { ...request, signal: options.signal } : request);
      cache.set(job.asset, data);
      resolved[job.asset] = data;
      options.onProgress?.({ stage: 'images', done: ++done, total });
    }
  };
  await Promise.all(Array.from({ length: Math.max(1, Math.min(options.concurrency ?? 3, total)) }, worker));

  return resolved;
}
