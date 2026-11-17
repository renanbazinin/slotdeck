// A second look at a draft, beyond the template's limits: problems found here go back to the
// writing model in the same repair round as the limits it broke. Nothing reviews a draft unless
// the caller asks for it.
import { freeFor } from '../free/resolve.ts';
import { checkDesign } from '../deck/design.ts';
import type { DesignOptions } from '../deck/design.ts';
import { isImageRef } from '../deck/spec.ts';
import type { DeckSpec } from '../deck/spec.ts';
import { DeckValidationError } from '../deck/validate.ts';
import type { Issue, LayoutMode } from '../deck/validate.ts';
import type { JsonRequest, JsonSchema, LlmProvider } from '../llm/provider.ts';
import type { PreviewProvider, SlidePreview } from '../preview/provider.ts';
import { renderDeck } from '../render/render.ts';
import type { Template } from '../render/render.ts';
import { withLayouts } from '../template/compose.ts';
import { effectiveSlots } from '../template/manifest.ts';
import { solidPng } from '../util/png.ts';
import type { AssetDef, IconDef } from './catalog.ts';

/**
 * Looks at a draft deck and returns problems to send back to the writing model. Each problem
 * names a slide, and a slot when it can. Give one to the planner as `review`.
 */
export type DeckReview = (deck: DeckSpec, context: { signal?: AbortSignal }) => Issue[] | Promise<Issue[]>;

/**
 * A review made of the code-only design checks ({@link checkDesign}): boxes left mostly empty,
 * uneven cards, headings that end in a lone word. Overflow is left out, since validation reports
 * it already. The template must be loaded with fonts. No model is called.
 */
export function designReview(template: Template, options: DesignOptions = {}): DeckReview {
  return (deck) => checkDesign(template, deck, { kinds: ['empty', 'uneven', 'orphan'], ...options });
}

export interface VisionReviewOptions {
  template: Template;
  /** A model that can look at pictures. It is asked once for every `slidesPerRequest` slides. */
  llm: LlmProvider;
  /** Draws the draft. It must give bitmap pictures (PNG, JPEG, WebP or GIF): a model cannot read SVG. */
  preview: PreviewProvider;
  /** The app's icons and supplied pictures, so that they are drawn as they are. Any other picture is drawn as a plain rectangle. */
  icons?: IconDef[];
  assets?: AssetDef[];
  /** As for `renderDeck`. Defaults to `adaptive`, so that a draft with optional slots left out can be drawn. */
  layout?: LayoutMode;
  /** How many slides are shown to the model in one request. Defaults to 6. */
  slidesPerRequest?: number;
}

const BITMAP = /^image\/(png|jpe?g|webp|gif)$/i;

const SYSTEM = [
  'You check pictures of presentation slides for layout problems, for the program that wrote their text.',
  'You report only what can be seen in a picture and fixed by changing the text of a slot. You do not judge wording, facts or the design of the template.',
].join('\n');

const INSTRUCTIONS = [
  'Report a problem only when you can see it in the picture:',
  '- text that is cut off, runs out of its box or off the slide, or lies over other text or over a picture',
  '- text too small to read',
  '- a box, a card or a whole slide left mostly empty beside others that are full',
  '- a heading that ends with one word on a line of its own',
  'Photos may be plain coloured rectangles that stand in for pictures not made yet: ignore them.',
  'Name the slot the problem is in when you can tell which it is, and say in one sentence what is wrong and how the text should change (shorter, longer, about how much).',
  'If a slide has no such problem, report nothing for it. Do not report doubts.',
].join('\n');

const short = (text: string) => (text.length > 120 ? `${text.slice(0, 119)}…` : text);

/**
 * A review by a model that looks at pictures of the draft: the deck is rendered with stand-ins
 * for pictures that do not exist yet, drawn by the preview provider, and shown to the model with
 * the text of each slot. What the model reports comes back as problems for the repair round.
 *
 * It is only as good as the pictures. A preview that differs from PowerPoint makes the model
 * report problems that are not there, so use a provider whose pictures can be trusted.
 */
export function visionReview(options: VisionReviewOptions): DeckReview {
  const { llm, preview } = options;
  const supplied = new Map<string, Uint8Array>([...(options.icons ?? []), ...(options.assets ?? [])].map((item) => [item.id, item.data]));
  const standIn = solidPng(96, 64, [176, 184, 190]);
  const perRequest = Math.max(1, options.slidesPerRequest ?? 6);

  return async (deck, { signal }) => {
    const template = withLayouts(options.template, deck.layouts);
    const assets: Record<string, Uint8Array> = {};
    for (const slide of deck.slides) {
      for (const value of Object.values(slide.slots ?? {})) if (isImageRef(value)) assets[value.asset] = supplied.get(value.asset) ?? standIn;
    }
    let pptx: Uint8Array;
    try {
      pptx = renderDeck(template, deck, { assets, layout: options.layout ?? 'adaptive' }).pptx;
    } catch (error) {
      // A draft that cannot be rendered has errors of its own, which validation reports.
      if (error instanceof DeckValidationError) return [];
      throw error;
    }
    const pictures = await preview.renderSlides(signal ? { pptx, signal } : { pptx });
    const unreadable = pictures.find((picture) => !BITMAP.test(picture.mimeType));
    if (unreadable) throw new Error(`visionReview needs bitmap pictures such as PNG, and the preview provider gave ${unreadable.mimeType}`);
    const bySlide = new Map<number, SlidePreview>(pictures.map((picture) => [picture.slide, picture]));

    const issues: Issue[] = [];
    const numbers = deck.slides.map((_, index) => index + 1).filter((number) => bySlide.has(number));
    for (let start = 0; start < numbers.length; start += perRequest) {
      const batch = numbers.slice(start, start + perRequest);
      const slotsOf = new Map<number, string[]>();
      const listing = batch.map((number) => {
        const slide = deck.slides[number - 1] as DeckSpec['slides'][number];
        const free = Array.isArray(slide.blocks) ? freeFor(options.template, deck, slide) : undefined;
        const pattern = free === undefined ? template.manifest.patterns.find((entry) => entry.id === slide.pattern) : typeof free === 'string' ? undefined : free.plan.pattern;
        const names: string[] = [];
        const texts: string[] = [];
        for (const slot of pattern ? effectiveSlots(pattern, slide.items) : []) {
          const value = slide.slots?.[slot.name];
          if (slot.kind !== 'text' || value === undefined || value === null) continue;
          names.push(slot.name);
          texts.push(`- ${slot.name}: ${JSON.stringify(short(Array.isArray(value) ? value.join(' / ') : String(value)))}`);
        }
        slotsOf.set(number, names);
        return `s${number} (${slide.pattern}):\n${texts.join('\n') || '- no text'}`;
      });
      const schema: JsonSchema = {
        type: 'object',
        properties: {
          problems: {
            type: 'array',
            items: {
              type: 'object',
              properties: {
                slide: { type: 'integer', enum: batch, description: 'Number of the slide, as in its "s" label' },
                slot: { type: 'string', description: 'Name of the slot the problem is in, or an empty string when it cannot be told' },
                problem: { type: 'string', description: 'One sentence of at most 25 words: what is wrong, and how the text should change' },
              },
              required: ['slide', 'slot', 'problem'],
            },
          },
        },
        required: ['problems'],
      };
      const request: JsonRequest = {
        system: SYSTEM,
        prompt: [
          `The pictures show slides ${batch.map((number) => `s${number}`).join(', ')} of a deck, in that order, one picture for each.`,
          `The text of each slide's slots:\n${listing.join('\n\n')}`,
          INSTRUCTIONS,
        ].join('\n\n'),
        schema,
        images: batch.map((number) => {
          const picture = bySlide.get(number) as SlidePreview;
          return { mimeType: picture.mimeType, data: picture.data };
        }),
        // A sentence for each problem, and rarely more than a few problems on a slide.
        maxOutputTokens: 300 + 200 * batch.length,
      };
      if (signal) request.signal = signal;
      const answer = (await llm.generateJson(request)) as { problems?: unknown } | null;
      for (const item of Array.isArray(answer?.problems) ? answer.problems : []) {
        const raw = item as { slide?: unknown; slot?: unknown; problem?: unknown };
        const slide = Number(raw?.slide);
        const problem = String(raw?.problem ?? '').replace(/\s+/g, ' ').trim();
        if (!batch.includes(slide) || problem === '') continue;
        const issue: Issue = { severity: 'warning', slide, message: `seen in the picture of the slide: ${problem}` };
        const slot = String(raw?.slot ?? '').trim();
        if (slotsOf.get(slide)?.includes(slot)) issue.slot = slot;
        issues.push(issue);
      }
    }
    return issues;
  };
}
