// Character budgets worked out from the template's fonts: how much text each text slot's box
// really holds, to set in place of the budgets estimated from the box's size.
import type { Template } from '../render/render.ts';
import { textMeasurer } from '../text/measure.ts';
import { inspectTemplate } from './inspect.ts';
import { boxCapacity } from './sample.ts';

export interface MeasuredBudget {
  pattern: string;
  slot: string;
  /** The budget the manifest states, when it states one. */
  maxChars?: number;
  /** The characters of running text the slot's box holds, as measured. */
  fits: number;
  /** A budget to set: what fits, less the margin. */
  suggested: number;
}

export interface BudgetOptions {
  /**
   * Running text to measure with, in place of each slot's own sample words: a paragraph in the
   * language the decks will be written in, when that is not the template's. Letters of one
   * language are wider than another's, so a box holds a different number of them.
   */
  sample?: string;
  /**
   * The share of what fits to suggest as a budget. Text of the same length does not always take
   * the same room, because lines break between words. Defaults to 0.9.
   */
  margin?: number;
}

/**
 * Measures, for every text slot a deck fills, how many characters its box holds at the slot's
 * typeface and size. A slot with several paragraphs is measured with as many as it allows. The
 * template must be loaded with fonts; slots set in a typeface that was not supplied are left out.
 *
 * The result is for setting budgets, e.g. with `changeManifest`:
 * `{ op: 'setSlot', pattern, slot, maxChars: suggested }`.
 */
export function measuredBudgets(template: Template, options: BudgetOptions = {}): MeasuredBudget[] {
  const measure = textMeasurer(template);
  const info = inspectTemplate(template.pkg);
  const margin = options.margin ?? 0.9;
  const given = options.sample?.split(/\s+/).filter(Boolean) ?? [];
  const budgets: MeasuredBudget[] = [];
  for (const pattern of template.manifest.patterns) {
    const shapes = new Map((info.slides[pattern.sourceSlide - 1]?.shapes ?? []).map((shape) => [shape.id, shape]));
    for (const slot of pattern.slots) {
      if (slot.kind !== 'text' || slot.empty === 'drop' || slot.empty === 'keep') continue;
      const own = (shapes.get(slot.shapeId)?.paragraphs ?? []).join(' ').split(/\s+/).filter(Boolean);
      const words = given.length > 0 ? given : own;
      if (words.length === 0) continue;
      const fits = boxCapacity(measure, pattern, slot, words, Math.max(1, slot.maxParagraphs ?? 1));
      if (fits === undefined) continue;
      const budget: MeasuredBudget = { pattern: pattern.id, slot: slot.name, fits, suggested: Math.max(1, Math.floor(fits * margin)) };
      if (slot.maxChars !== undefined) budget.maxChars = slot.maxChars;
      budgets.push(budget);
    }
  }
  return budgets;
}
