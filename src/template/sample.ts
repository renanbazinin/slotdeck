import type { DeckSpec, SlideSpec, SlotValue } from '../deck/spec.ts';
import type { Template } from '../render/render.ts';
import { solidPng } from '../util/png.ts';
import { measurerFor } from '../text/measure.ts';
import type { TextMeasurer } from '../text/measure.ts';
import { inspectTemplate } from './inspect.ts';
import { COPY_ID_STEP, defaultItems, effectiveSlots, itemRange, tableAtScale, textBudget } from './manifest.ts';
import type { PatternDef, TextSlotDef } from './manifest.ts';

export interface SampleOptions {
  /** How full to make each text slot, as a share of its budget. Defaults to 0.95. */
  fill?: number;
  /** For patterns with a row of items: `drawn` (default) uses the template's number, or the nearest the row takes; `max` the most allowed; `min` the fewest. */
  items?: 'drawn' | 'max' | 'min';
  /**
   * What "full" means for a text slot. `budget` (default) is the slot's character budget. `box` is
   * what the slot's box holds as measured with the template's fonts, whatever the budget says:
   * the deck then shows every box at its true capacity. `box` needs a template loaded with fonts;
   * a slot that cannot be measured is filled to its budget.
   */
  fit?: 'budget' | 'box';
}

/**
 * Builds a deck that uses every pattern once and fills every slot close to its limit, with the
 * template's own sample words. Rendering it is the quickest way to see whether a manifest's
 * budgets are right: text that overflows here will overflow in real decks.
 */
export function sampleDeck(template: Template, options: SampleOptions = {}): { deck: DeckSpec; assets: Record<string, Uint8Array> } {
  const { manifest } = template;
  const info = inspectTemplate(template.pkg);
  const fill = options.fill ?? 0.95;
  const measure = options.fit === 'box' ? measurerFor(template) : undefined;
  const assets: Record<string, Uint8Array> = {
    'sample-photo': solidPng(96, 64, [120, 150, 135]),
    'sample-icon': solidPng(32, 32, [105, 36, 232]),
    'sample-graphic': solidPng(120, 60, [200, 205, 210]),
  };
  const fields: Record<string, string> = {};

  const slides = manifest.patterns.map((pattern): SlideSpec => {
    const shapes = new Map((info.slides[pattern.sourceSlide - 1]?.shapes ?? []).map((shape) => [shape.id, shape]));
    const range = itemRange(pattern);
    // By default the number the template draws, or the nearest the row takes, so that the sample is a deck the manifest allows.
    const items = range ? (options.items === 'max' ? range.max : options.items === 'min' ? range.min : defaultItems(pattern)) : undefined;
    const slots: Record<string, SlotValue> = {};

    for (const slot of effectiveSlots(pattern, items)) {
      const policy = slot.empty ?? 'error';
      if (policy === 'drop' || policy === 'keep') continue;
      // A copied item's shapes exist only once the row is rendered; its words come from the shape it copies.
      const shape = shapes.get(slot.shapeId) ?? shapes.get(String(Number(slot.shapeId) % COPY_ID_STEP));

      if (slot.kind === 'text') {
        const words = (shape?.paragraphs ?? []).join(' ').split(/\s+/).filter(Boolean);
        if (slot.deckField) {
          fields[slot.deckField] ??= words.join(' ') || slot.deckField;
          continue;
        }
        // The sample is filled to the budgets real decks get, which the template's own text scale shrinks.
        const budget = textBudget(slot, manifest.textScale ?? 1);
        const several = (slot.maxParagraphs ?? 1) > 1;
        const paragraphs = several ? (budget.maxParagraphs ?? 2) : 1;
        const held = measure ? boxCapacity(measure, pattern, slot, words, paragraphs, manifest.textScale ?? 1, items) : undefined;
        const total =
          held !== undefined ? Math.max(1, Math.floor(held * fill)) : budget.maxChars !== undefined ? Math.max(1, Math.floor(budget.maxChars * fill)) : words.join(' ').length || 12;
        const each = Math.max(1, Math.floor(total / paragraphs));
        const text = Array.from({ length: paragraphs }, (_, index) => filler(words, each, index));
        slots[slot.name] = several ? text : (text[0] as string);
      } else if (slot.kind === 'image') {
        slots[slot.name] = { asset: slot.imageKind === 'photo' ? 'sample-photo' : slot.imageKind === 'icon' ? 'sample-icon' : 'sample-graphic' };
      } else if (slot.kind === 'table') {
        const columns = slot.maxColumns ?? 3;
        const cell = 'x'.repeat(Math.max(1, Math.floor((tableAtScale(slot, manifest.textScale ?? 1).maxCellChars ?? 10) * fill)));
        slots[slot.name] = {
          header: Array.from({ length: columns }, () => cell),
          rows: Array.from({ length: slot.maxRows ?? 3 }, () => Array.from({ length: columns }, () => cell)),
        };
      } else {
        const categories = Math.min(slot.maxCategories ?? 4, 6);
        slots[slot.name] = {
          type: 'column',
          categories: Array.from({ length: categories }, (_, index) => `C${index + 1}`),
          series: Array.from({ length: Math.min(slot.maxSeries ?? 1, 2) }, (_, series) => ({
            name: `S${series + 1}`,
            values: Array.from({ length: categories }, (_, index) => 20 + ((index * 17 + series * 11) % 60)),
          })),
        };
      }
    }
    const slide: SlideSpec = { pattern: pattern.id, slots };
    if (items !== undefined) slide.items = items;
    return slide;
  });

  const deck: DeckSpec = { version: 1, title: `Sample of ${manifest.name}`, direction: manifest.direction, slides };
  if (manifest.language) deck.language = manifest.language;
  if (Object.keys(fields).length > 0) deck.fields = fields;
  return { deck, assets };
}

/**
 * The most characters of filler text, in a number of paragraphs, that a slot's box holds as
 * measured. `undefined` when the slot cannot be measured. `items` is the number of items on the
 * slide, for a slot of a row whose items widen when there are fewer of them.
 */
export function boxCapacity(measure: TextMeasurer, pattern: PatternDef, slot: TextSlotDef, words: string[], paragraphs: number, textScale = 1, items?: number): number | undefined {
  const make = (total: number) => Array.from({ length: paragraphs }, (_, index) => filler(words, Math.max(1, Math.floor(total / paragraphs)), index));
  const options = items === undefined ? { textScale } : { textScale, items };
  const fits = (total: number) => measure.fit(pattern, slot, make(total), options)?.fits;
  if (fits(paragraphs) === undefined) return undefined;
  let low = paragraphs;
  let high = 6000;
  while (low < high) {
    const middle = Math.ceil((low + high) / 2);
    if (fits(middle)) low = middle;
    else high = middle - 1;
  }
  return low;
}

/** Text of about `length` characters, made of whole words where they fit, starting at a different word each time. */
export function filler(words: string[], length: number, start: number): string {
  if (words.length === 0) return 'x'.repeat(length);
  let text = '';
  for (let index = start % words.length; text.length < length; index = (index + 1) % words.length) {
    const next = text ? `${text} ${words[index]}` : (words[index] as string);
    if (next.length > length) {
      // No room for another whole word: pad a first word, or stop.
      if (!text) text = next.slice(0, length);
      break;
    }
    text = next;
  }
  return text;
}
