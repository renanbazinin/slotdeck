// Reads a manifest against its pptx and lists what the slides allow that the manifest does not
// use yet: a row with room for more items, a picture that is a pictogram, a variant of a layout.
// No model is asked. Each suggestion carries the changes that make it, ready for changeManifest,
// and has been tried: one that the pptx would refuse is not listed.
import type { Template } from '../render/render.ts';
import { changeWith } from './change.ts';
import type { ManifestChange } from './change.ts';
import { deriveVariant } from './derive.ts';
import type { PatternVariant, VariantContext, VariantRequest } from './derive.ts';
import { inspectTemplate } from './inspect.ts';
import type { TemplateInfo } from './inspect.ts';
import type { PatternDef, TemplateManifest } from './manifest.ts';
import { isPictogram } from './measure.ts';
import { detectRepeat, rowCapacity, rowComfort } from './rows.ts';

export const SUGGESTION_KINDS = ['icon', 'row', 'row-max', 'row-widen', 'variant'] as const;
export type SuggestionKind = (typeof SUGGESTION_KINDS)[number];

export interface ManifestSuggestion {
  /**
   * `icon`: pictures labelled photo that are pictograms. `row`: items named alike that are not a
   * row yet. `row-max`: a row with room for more items than it takes. `row-widen`: a row or grid
   * whose items could grow to fill the line when there are fewer. `variant`: a layout the slide
   * also allows (mirrored, text only, without an optional text).
   */
  kind: SuggestionKind;
  /** The pattern the suggestion is about. */
  pattern: string;
  /** That pattern's slide in the template, from 1. */
  slide: number;
  /** For `variant`: which one. */
  variant?: PatternVariant;
  /** For `variant` `without`: the text that would be left out. */
  slot?: string;
  /** What is so now, and what the change would do. */
  message: string;
  /** The changes that make it. Pass them to `changeManifest`. */
  changes: ManifestChange[];
}

export interface SuggestOptions {
  /** Only these kinds. Defaults to all of them. */
  kinds?: readonly SuggestionKind[];
  /** The factor suggested for `row-widen`. Defaults to 2: an item may become twice as wide. */
  widen?: number;
}

const pct = (fraction: number) => `${Math.round(fraction * 100)}%`;
const quoted = (names: string[]) => names.map((name) => `"${name}"`).join(', ');

/**
 * Lists what a template's slides allow that its manifest does not use yet, each with the changes
 * that make it. Works on any manifest, drafted or written by hand. Nothing is changed: apply the
 * suggestions you want with `changeManifest(template, suggestion.changes)`, then ask again, since
 * one change can open another (a variant of a pattern whose row was just made to take more).
 */
export function suggestChanges(template: Template, options: SuggestOptions = {}): ManifestSuggestion[] {
  const kinds = new Set<SuggestionKind>(options.kinds ?? SUGGESTION_KINDS);
  const factor = options.widen ?? 2;
  const info = inspectTemplate(template.pkg);
  const manifest = template.manifest;
  const context: VariantContext = { pkg: template.pkg, info, manifest };
  const found: ManifestSuggestion[] = [];

  // A variant made earlier shares its slide with the pattern it came from, which is listed first.
  const firstOnSlide = new Set<string>();
  const slidesSeen = new Set<number>();
  for (const pattern of manifest.patterns) {
    if (!slidesSeen.has(pattern.sourceSlide)) firstOnSlide.add(pattern.id);
    slidesSeen.add(pattern.sourceSlide);
  }

  for (const pattern of manifest.patterns) {
    const slide = info.slides[pattern.sourceSlide - 1];
    if (!slide) continue;
    const add = (suggestion: Omit<ManifestSuggestion, 'pattern' | 'slide'>) => found.push({ ...suggestion, pattern: pattern.id, slide: pattern.sourceSlide });
    const byId = new Map(slide.shapes.map((shape) => [shape.id, shape]));

    if (kinds.has('icon')) {
      const small = pattern.slots.filter((slot) => slot.kind === 'image' && slot.imageKind === 'photo' && isPictogramShape(byId.get(slot.shapeId), info));
      if (small.length > 0) {
        // A change to one item of a row is made to every item, so one change per row slot is enough.
        const firstOfRow = new Set((pattern.repeat?.slots ?? []).map((name) => name.replace('{n}', '1')));
        const inRow = new Set((pattern.repeat?.slots ?? []).flatMap((name) => (pattern.repeat?.shapes ?? []).map((_, index) => name.replace('{n}', String(index + 1)))));
        const targets = small.filter((slot) => !inRow.has(slot.name) || firstOfRow.has(slot.name));
        const widest = Math.max(...small.map((slot) => byId.get(slot.shapeId)?.box.w ?? 0));
        add({
          kind: 'icon',
          message: `${quoted(small.map((slot) => slot.name))} ${small.length === 1 ? 'is a photo slot, but its picture is' : 'are photo slots, but their pictures are'} no more than ${pct(widest)} of the slide wide: pictograms. As icons they are filled from the icon set, and the layout no longer waits for photographs.`,
          changes: targets.map((slot) => ({ op: 'setSlot', pattern: pattern.id, slot: slot.name, imageKind: 'icon' })),
        });
      }
    }

    const repeat = pattern.repeat;
    if (!repeat && kinds.has('row')) {
      const detected = detectRepeat(pattern.slots.filter((slot) => !(slot.kind === 'text' && slot.deckField)), slide);
      // A row that must hold exactly what it draws, with no room for more, changes nothing for a deck.
      if (detected && ((detected.max ?? 0) > (detected.min ?? 0) || rowComfort(slide, detected).most > detected.shapes.length)) {
        add({
          kind: 'row',
          message: `The slots of "${pattern.id}" name ${detected.shapes.length} ${detected.item}s alike, but the pattern has no row, so a deck must fill all ${detected.shapes.length}. As a ${detected.columns ? 'grid' : 'row'}, a slide can use ${detected.min} to ${detected.max} of them.`,
          changes: [{ op: 'setRepeat', pattern: pattern.id }],
        });
      }
    }

    if (repeat) {
      const drawn = repeat.shapes.length;
      const most = repeat.max ?? drawn;
      const fewest = repeat.min ?? drawn;
      const comfort = rowComfort(slide, repeat);
      const capacity = comfort.most;
      const physical = rowCapacity(slide, repeat);
      if (kinds.has('row-max') && capacity > most) {
        // The number holds with the row spread into the room beside it, so the span goes with it.
        const spread = comfort.span ? ` For that the row spreads from ${pct(comfort.span.left)} to ${pct(comfort.span.right)} of the slide's width.` : '';
        add({
          kind: 'row-max',
          message: `The row of ${repeat.item}s in "${pattern.id}" takes up to ${most}, and the slide has room for ${capacity} side by side at their drawn width, with the space between them no less than a third of what the template draws.${spread}${physical > capacity ? ` Up to ${physical} fit when they sit closer than that.` : ''}`,
          changes: [{ op: 'setRepeat', pattern: pattern.id, max: capacity, ...(comfort.span ? { span: comfort.span } : {}) }],
        });
      }
      const line = repeat.columns ?? drawn;
      if (kinds.has('row-widen') && repeat.widen === undefined && line >= 2 && fewest < drawn) {
        add({
          kind: 'row-widen',
          message: `With fewer ${repeat.item}s than "${pattern.id}" draws on a line (${line}), they keep their width and the line is shorter. Widening lets each grow up to ${factor} times as wide, so the line stays full and each holds more text.`,
          changes: [{ op: 'setRepeat', pattern: pattern.id, widen: factor }],
        });
      }
    }

    if (kinds.has('variant') && firstOnSlide.has(pattern.id)) {
      const requests: VariantRequest[] = [
        { variant: 'mirror' },
        { variant: 'text-only' },
        ...pattern.slots.filter((slot) => slot.kind === 'text' && slot.empty === 'remove' && !slot.deckField).map((slot): VariantRequest => ({ variant: 'without', slot: slot.name })),
      ];
      for (const request of requests) {
        const made = deriveVariant(pattern, request, context);
        if ('refused' in made || alreadyMade(manifest, pattern, made.pattern, made.suffix)) continue;
        const suggestion: Omit<ManifestSuggestion, 'pattern' | 'slide'> = {
          kind: 'variant',
          variant: request.variant,
          message: `The slide of "${pattern.id}" also allows a ${request.variant === 'mirror' ? 'mirrored' : request.variant === 'text-only' ? 'text-only' : `"${request.slot}"-less`} layout. ${made.summary}`,
          changes: [{ op: 'derivePattern', pattern: pattern.id, variant: request.variant, ...(request.slot !== undefined ? { slot: request.slot } : {}) }],
        };
        if (request.slot !== undefined) suggestion.slot = request.slot;
        add(suggestion);
      }
    }
  }

  // Listed only when the change goes through as it stands.
  return found.filter((suggestion) => changeWith(template, suggestion.changes, info).rejected.length === 0);
}

const isPictogramShape = (shape: TemplateInfo['slides'][number]['shapes'][number] | undefined, info: TemplateInfo): boolean => !!shape && isPictogram(shape, info);

/** Whether the manifest already has this variant of a pattern: by its id, or by what it holds. */
function alreadyMade(manifest: TemplateManifest, source: PatternDef, variant: PatternDef, suffix: string): boolean {
  const content = (pattern: PatternDef) => JSON.stringify([pattern.slots, pattern.repeat ?? null]);
  const wanted = content(variant);
  return manifest.patterns.some((other) => other !== source && other.sourceSlide === source.sourceSlide && (other.id === `${source.id}-${suffix}` || content(other) === wanted));
}
