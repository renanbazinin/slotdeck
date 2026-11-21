// Pictures of a template's patterns and of a deck, through the preview provider.
import type { DeckSpec } from '../deck/spec.ts';
import type { LayoutMode } from '../deck/validate.ts';
import { renderDeck } from '../render/render.ts';
import type { Template } from '../render/render.ts';
import { sampleDeck } from '../template/sample.ts';
import type { SampleOptions } from '../template/sample.ts';
import type { PreviewProvider, SlidePreview } from './provider.ts';

export interface PreviewDeckOptions {
  /** Picture bytes by asset id, as for `renderDeck`. */
  assets?: Record<string, Uint8Array>;
  layout?: LayoutMode;
  signal?: AbortSignal;
}

/**
 * Renders a deck and draws its slides: `renderDeck`, then the provider. For showing a deck
 * before or instead of offering the file.
 */
export async function previewDeck(template: Template, deck: DeckSpec, preview: PreviewProvider, options: PreviewDeckOptions = {}): Promise<{ pptx: Uint8Array; previews: SlidePreview[] }> {
  const render: { assets?: Record<string, Uint8Array>; layout?: LayoutMode } = {};
  if (options.assets) render.assets = options.assets;
  if (options.layout) render.layout = options.layout;
  const { pptx } = renderDeck(template, deck, render);
  const direction = deck.direction ?? template.manifest.direction;
  const previews = await preview.renderSlides(options.signal ? { pptx, direction, signal: options.signal } : { pptx, direction });
  return { pptx, previews };
}

/**
 * A picture of every pattern of a template, by pattern id: the thumbnails to show a plan with.
 * Each is the pattern's slide of the template's sample deck, so a pattern derived from another
 * pattern's slide is shown as it will look, not as the slide it was made from.
 */
export async function patternPreviews(template: Template, preview: PreviewProvider, options: SampleOptions & { signal?: AbortSignal } = {}): Promise<Record<string, SlidePreview>> {
  const { signal, ...sample } = options;
  const { deck, assets } = sampleDeck(template, sample);
  const { previews } = await previewDeck(template, deck, preview, signal ? { assets, layout: 'adaptive', signal } : { assets, layout: 'adaptive' });
  const bySlide = new Map(previews.map((picture) => [picture.slide, picture]));
  const pictures: Record<string, SlidePreview> = {};
  template.manifest.patterns.forEach((pattern, index) => {
    const picture = bySlide.get(index + 1);
    if (picture) pictures[pattern.id] = picture;
  });
  return pictures;
}
