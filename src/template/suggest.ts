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
import type { Box, SlideInfo, TemplateInfo } from './inspect.ts';
import { widenFactor } from './manifest.ts';
import type { PatternDef, SlotDef, TemplateManifest } from './manifest.ts';
import { isPictogram } from './measure.ts';
import { detectRepeat, droppedShapes, rowCapacity, rowComfort, standsOut, strandedSlots } from './rows.ts';
import { rowRoom } from './measure.ts';

export const SUGGESTION_KINDS = ['icon', 'row', 'row-max', 'row-widen', 'row-emphasis', 'variant'] as const;
export type SuggestionKind = (typeof SUGGESTION_KINDS)[number];

export interface ManifestSuggestion {
  /**
   * `icon`: pictures labelled photo that are pictograms. `row`: items named alike that are not a
   * row yet. `row-max`: a row with room for more items than it takes. `row-widen`: a row or grid
   * whose items could grow to fill the line when there are fewer. `row-emphasis`: a row whose
   * slide draws one item to stand out, which a deck that reads the other way should move with its
   * look (from 0.7.0). `variant`: a layout the slide also allows (mirrored, text only, without an
   * optional text).
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
  /** The changes that make it, never empty. Pass them to `changeManifest`. */
  changes: ManifestChange[];
}

export interface SuggestOptions {
  /** Only these kinds. Defaults to all of them. */
  kinds?: readonly SuggestionKind[];
  /** The factor suggested for `row-widen`. Defaults to 2: an item may become twice as wide. */
  widen?: number;
}

export type ImageSlot = Extract<SlotDef, { kind: 'image' }>;

const pct = (fraction: number) => `${Math.round(fraction * 100)}%`;
const quoted = (names: string[]) => names.map((name) => `"${name}"`).join(', ');
const middle = (box: Box) => box.x + box.w / 2;
/** The area a slot draws in: its own box when the manifest moves the shape, otherwise the shape's. */
const areaOf = (slot: SlotDef | undefined, slide: SlideInfo, shapeId: string): Box | undefined =>
  (slot && 'box' in slot && slot.box ? slot.box : undefined) ?? slide.shapes.find((shape) => shape.id === shapeId)?.box;

/**
 * Lists what a template's slides allow that its manifest does not use yet, each with the changes
 * that make it. Works on any manifest, drafted or written by hand. Nothing is changed: apply the
 * suggestions you want with `changeManifest(template, suggestion.changes)`, then ask again, since
 * one change can open another (a variant of a pattern whose row was just made to take more).
 */
export function suggestChanges(template: Template, options: SuggestOptions = {}): ManifestSuggestion[] {
  const kinds = new Set<SuggestionKind>(options.kinds ?? SUGGESTION_KINDS);
  for (const kind of kinds) if (!SUGGESTION_KINDS.includes(kind)) throw new Error(`"${String(kind)}" is not a kind of suggestion; the kinds are ${SUGGESTION_KINDS.join(', ')}`);
  const factor = options.widen ?? 2;
  if (kinds.has('row-widen') && !(typeof factor === 'number' && factor > 1 && factor <= 4)) throw new Error('widen is a factor above 1 and at most 4, such as 2 to let an item become twice as wide');
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
    const repeat = pattern.repeat;
    const removed = droppedShapes(pattern.slots);
    // The row a pattern has, or the one its slot names describe and a `row` suggestion would make.
    const row = repeat ?? detectRepeat(pattern.slots, slide);

    if (kinds.has('icon')) {
      // Pictures are judged by the row they are in or are about to be in, so that taking `icon`
      // and `row` together cannot leave a row whose items are filled unlike.
      for (const family of pictureFamilies(row ? { ...pattern, repeat: row } : pattern)) {
        const photos = family.filter((slot) => slot.imageKind === 'photo');
        // A change to one item of a row is made to every item, so it is offered only where it is
        // true of them all: each picture of the row measures as a pictogram.
        if (photos.length === 0 || !family.every((slot) => (slot.imageKind === 'photo' || slot.imageKind === 'icon') && measuresAsPictogram(slot, slide, info))) continue;
        const widest = Math.max(...photos.map((slot) => areaOf(slot, slide, slot.shapeId)?.w ?? 0));
        const one = photos.length === 1;
        add({
          kind: 'icon',
          message: `${quoted(photos.map((slot) => slot.name))} ${one ? 'is a photo slot, but its picture is' : 'are photo slots, but their pictures are'} no more than ${pct(widest)} of the slide wide: ${one ? 'a pictogram' : 'pictograms'}. As ${one ? 'an icon slot it is' : 'icon slots they are'} filled from the icon set given to the planner, and the layout no longer waits for photographs. Without an icon set, a required icon slot keeps the layout from being offered.`,
          // In a row one change reaches every item. Where the row is not made yet, each slot is changed by itself.
          changes: (repeat ? photos.slice(0, 1) : photos).map((slot) => ({ op: 'setSlot', pattern: pattern.id, slot: slot.name, imageKind: 'icon' })),
        });
      }
    }

    if (!repeat && kinds.has('row')) {
      const detected = row;
      if (detected) {
        const count = detected.shapes.length;
        const ranges = (detected.max ?? count) > (detected.min ?? count);
        // A row that must hold exactly what it draws, with no room for more, changes nothing for a deck.
        if (ranges || rowComfort(slide, detected, removed).most > count) {
          const kind = detected.columns ? 'grid' : 'row';
          add({
            kind: 'row',
            message: `The slots of "${pattern.id}" name ${count} ${detected.item}s alike, but the pattern has no ${kind}, so every slide draws all ${count}. ${ranges ? `As a ${kind}, a slide can use ${detected.min} to ${detected.max} of them.` : `As a row it takes exactly ${count}, and can then be given the room the slide has for more.`}`,
            changes: [{ op: 'setRepeat', pattern: pattern.id }],
          });
        }
      }
    }

    if (repeat) {
      const drawn = repeat.shapes.length;
      const most = repeat.max ?? drawn;
      const fewest = repeat.min ?? drawn;
      // A further item is a copy of the last. Where the last item holds a slot that is not the
      // row's and that a deck fills or removes, a copy would show the template's own sample there.
      const copies = !strandedSlots(pattern, slide).some((entry) => entry.item === drawn && pattern.slots.find((slot) => slot.name === entry.slot)?.empty !== 'keep');
      // A grid takes what it draws; one that takes fewer was set so by a person.
      if (kinds.has('row-max') && repeat.columns === undefined && copies) {
        const comfort = rowComfort(slide, repeat, removed);
        const physical = rowCapacity(slide, repeat, removed);
        if (comfort.most > most) {
          // The number holds with the row spread into the room beside it, so the span goes with it.
          const spread = comfort.span ? ` For that the row spreads from ${pct(comfort.span.left)} to ${pct(comfort.span.right)} of the slide's width.` : '';
          add({
            kind: 'row-max',
            message: `The row of ${repeat.item}s in "${pattern.id}" takes up to ${most}, and the slide has room for ${comfort.most} side by side at their drawn width, with the space between them no less than a third of what the template draws.${spread}${physical > comfort.most ? ` Up to ${physical} fit when they sit closer than that.` : ''}`,
            changes: [{ op: 'setRepeat', pattern: pattern.id, max: comfort.most, ...(comfort.span ? { span: comfort.span } : {}) }],
          });
        }
      }
      if (kinds.has('row-emphasis') && repeat.emphasis === undefined) {
        const odd = standsOut(template.pkg, slide.part, repeat.shapes);
        if (odd !== undefined) {
          add({
            kind: 'row-emphasis',
            message: `${repeat.item.charAt(0).toUpperCase()}${repeat.item.slice(1)} ${odd} of "${pattern.id}" is drawn unlike the other ${repeat.item}s. In a deck that reads the other way their texts now change places and its look stays where it is drawn, so it marks another ${repeat.item}. Marked as the one that stands out, the ${repeat.item}s change places with their looks, and it stays ${repeat.item} ${odd} for the deck's readers.`,
            changes: [{ op: 'setRepeat', pattern: pattern.id, emphasis: odd }],
          });
        }
      }
      if (kinds.has('row-widen') && repeat.widen === undefined) {
        // Offered only when it would change a slide: some number of items the row takes leaves its fullest line short.
        let widens = false;
        for (let count = fewest; count <= most && !widens; count++) widens = widenFactor({ ...repeat, widen: factor }, count) > 1;
        // Not for a row with an arrow or a line among its items: wider items would be drawn across it.
        const pinned = repeat.columns === undefined && drawn > 1 && rowRoom(slide, repeat.shapes, removed).pinned;
        if (widens && !pinned) {
          const line = repeat.columns ?? drawn;
          add({
            kind: 'row-widen',
            message: `With fewer ${repeat.item}s than "${pattern.id}" draws on a line (${line}), they keep their width and the line is shorter. Widening lets each grow up to ${factor} times as wide, so the line stays full and each holds more text.`,
            changes: [{ op: 'setRepeat', pattern: pattern.id, widen: factor }],
          });
        }
      }
    }

    if (kinds.has('variant') && firstOnSlide.has(pattern.id)) {
      const requests: VariantRequest[] = [
        { variant: 'mirror' },
        { variant: 'text-only' },
        ...pattern.slots.filter((slot) => slot.kind === 'text' && slot.empty === 'remove' && !slot.deckField).map((slot): VariantRequest => ({ variant: 'without', slot: slot.name })),
      ];
      const mirrored = deriveVariant(pattern, { variant: 'mirror' }, context);
      const mirror = 'refused' in mirrored ? undefined : mirrored.pattern;
      const others = manifest.patterns.filter((other) => other !== pattern && other.sourceSlide === pattern.sourceSlide);
      const made = new Set(others.map((other) => layoutOf(other, pattern, mirror, slide)));
      for (const request of requests) {
        const derived = deriveVariant(pattern, request, context);
        if ('refused' in derived) continue;
        // Known by the id a derived pattern gets, or by what it does to the slide.
        if (others.some((other) => other.id === `${pattern.id}-${derived.suffix}`) || made.has(layoutOf(derived.pattern, pattern, mirror, slide))) continue;
        const suggestion: Omit<ManifestSuggestion, 'pattern' | 'slide'> = {
          kind: 'variant',
          variant: request.variant,
          message: `The slide of "${pattern.id}" also allows a ${request.variant === 'mirror' ? 'mirrored layout' : request.variant === 'text-only' ? 'text-only layout' : `layout without "${request.slot}"`}: ${derived.summary}`,
          changes: [{ op: 'derivePattern', pattern: pattern.id, variant: request.variant, ...(request.slot !== undefined ? { slot: request.slot } : {}) }],
        };
        if (request.slot !== undefined) suggestion.slot = request.slot;
        add(suggestion);
      }
    }
  }

  // Listed only when there is something to do and it goes through as it stands.
  return found.filter((suggestion) => suggestion.changes.length > 0 && changeWith(template, suggestion.changes, info).rejected.length === 0);
}

/**
 * The picture slots of a pattern that something fills, in groups that change together: the same
 * slot of every item of a row is one group, because a change to one item is made to all of them;
 * a picture outside a row is a group of its own. A row slot that is not a picture in every item is
 * left out.
 */
export function pictureFamilies(pattern: PatternDef): ImageSlot[][] {
  const fillable = (slot: SlotDef | undefined): slot is ImageSlot => slot?.kind === 'image' && slot.empty !== 'keep' && slot.empty !== 'drop';
  const byName = new Map(pattern.slots.map((slot) => [slot.name, slot]));
  const inRow = new Set<string>();
  const families: ImageSlot[][] = [];
  for (const template of pattern.repeat?.slots ?? []) {
    const members = (pattern.repeat?.shapes ?? []).map((_, index) => byName.get(template.replace('{n}', String(index + 1))));
    for (const member of members) if (member) inRow.add(member.name);
    if (members.length > 0 && members.every(fillable)) families.push(members as ImageSlot[]);
  }
  for (const slot of pattern.slots) if (!inRow.has(slot.name) && fillable(slot)) families.push([slot]);
  return families;
}

/** Whether the area a picture slot draws in is that of a pictogram. A slot may move or resize its shape. */
function measuresAsPictogram(slot: ImageSlot, slide: SlideInfo, info: TemplateInfo): boolean {
  const shape = slide.shapes.find((candidate) => candidate.id === slot.shapeId);
  return !!shape && isPictogram(slot.box ? { ...shape, box: slot.box } : shape, info);
}

/**
 * What a pattern does to its slide, compared with the pattern it would be a variant of: which
 * shapes it takes away, what it fills that the source does not, and on which side it has what the
 * mirror image moves. Names, budgets, texts and exact boxes are left out, because
 * ordinary tuning changes them; a variant that was given its own id and then tuned is still the
 * same variant. A pattern with a native chart where the source has a picture is another layout,
 * and so are its variants.
 */
function layoutOf(pattern: PatternDef, source: PatternDef, mirror: PatternDef | undefined, slide: SlideInfo): string {
  const droppedBySource = droppedShapes(source.slots);
  const removed = [...droppedShapes(pattern.slots)].filter((id) => !droppedBySource.has(id)).sort();
  // What this pattern fills that the source does not fill with the same kind of thing: a picture
  // the source takes away, a native chart where the source has a picture.
  const filled = (of: PatternDef) => new Map(of.slots.filter((slot) => slot.kind !== 'text' && slot.empty !== 'drop' && slot.empty !== 'keep').map((slot) => [slot.shapeId, slot.kind]));
  const filledBySource = filled(source);
  const fills = [...filled(pattern)].filter(([id, kind]) => filledBySource.get(id) !== kind).map(([id, kind]) => `${id}:${kind}`).sort();
  const sides: string[] = [];
  for (const moved of mirror?.slots ?? []) {
    // The side is read from pictures and from what is moved as it is (a group, an ornament). A
    // text that is written is widened and tuned in ways that say nothing about sides.
    if (moved.kind === 'text' && moved.empty !== 'keep') continue;
    const from = areaOf(source.slots.find((slot) => slot.shapeId === moved.shapeId), slide, moved.shapeId);
    const to = areaOf(moved, slide, moved.shapeId);
    if (!from || !to || Math.abs(middle(to) - middle(from)) < 0.02) continue;
    const own = pattern.slots.find((slot) => slot.shapeId === moved.shapeId);
    if (own?.empty === 'drop' && !droppedBySource.has(moved.shapeId)) continue; // gone: counted above
    const at = areaOf(own, slide, moved.shapeId);
    if (at) sides.push(`${moved.shapeId}:${Math.abs(middle(at) - middle(to)) < Math.abs(middle(at) - middle(from)) ? 'mirrored' : 'as drawn'}`);
  }
  return JSON.stringify([removed, fills, [...new Set(sides)].sort()]);
}
