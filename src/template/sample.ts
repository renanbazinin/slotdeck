import type { DeckSpec, SlideSpec, SlotValue } from '../deck/spec.ts';
import type { Template } from '../render/render.ts';
import { solidPng } from '../util/png.ts';
import { inspectTemplate } from './inspect.ts';
import { COPY_ID_STEP, effectiveSlots, itemRange, tableAtScale, textBudget } from './manifest.ts';

export interface SampleOptions {
  /** How full to make each text slot, as a share of its budget. Defaults to 0.95. */
  fill?: number;
  /** For patterns with a row of items: `drawn` (default) uses the template's number, `max` the most allowed. */
  items?: 'drawn' | 'max' | 'min';
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
  const assets: Record<string, Uint8Array> = {
    'sample-photo': solidPng(96, 64, [120, 150, 135]),
    'sample-icon': solidPng(32, 32, [105, 36, 232]),
    'sample-graphic': solidPng(120, 60, [200, 205, 210]),
  };
  const fields: Record<string, string> = {};

  const slides = manifest.patterns.map((pattern): SlideSpec => {
    const shapes = new Map((info.slides[pattern.sourceSlide - 1]?.shapes ?? []).map((shape) => [shape.id, shape]));
    const range = itemRange(pattern);
    const items = range ? (options.items === 'max' ? range.max : options.items === 'min' ? range.min : range.drawn) : undefined;
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
        const total = budget.maxChars !== undefined ? Math.max(1, Math.floor(budget.maxChars * fill)) : words.join(' ').length || 12;
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

/** Text of about `length` characters, made of whole words where they fit, starting at a different word each time. */
function filler(words: string[], length: number, start: number): string {
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
