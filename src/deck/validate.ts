import { effectiveSlots, itemRange, slotSources, textBudget } from '../template/manifest.ts';
import type { PatternDef, SlotDef, TemplateManifest } from '../template/manifest.ts';
import { isChartValue, isImageRef, isTableValue, toParagraphs } from './spec.ts';
import type { ChartValue, DeckSpec, ImageRef, SlideSpec, TableValue } from './spec.ts';

export type LayoutMode = 'strict' | 'adaptive';

export interface Issue {
  severity: 'error' | 'warning';
  /** 1-based slide number in the deck; absent for deck-level issues. */
  slide?: number;
  slot?: string;
  message: string;
}

export class DeckValidationError extends Error {
  readonly issues: Issue[];

  constructor(issues: Issue[]) {
    super(
      `Deck cannot be rendered:\n${issues
        .map((issue) => `  - ${issue.slide ? `slide ${issue.slide}` : 'deck'}${issue.slot ? ` / ${issue.slot}` : ''}: ${issue.message}`)
        .join('\n')}`,
    );
    this.name = 'DeckValidationError';
    this.issues = issues;
  }
}

/** What the renderer should do with one slot on one slide. */
export type SlotAction =
  | { type: 'text'; paragraphs: string[] }
  | { type: 'image'; image: ImageRef }
  | { type: 'table'; table: TableValue }
  | { type: 'chart'; chart: ChartValue }
  | { type: 'remove' }
  | { type: 'keep' };

export interface ValidateOptions {
  layout?: LayoutMode;
  /** Ids of the assets that will be available at render time. When given, image references are checked. */
  assets?: Iterable<string>;
}

/**
 * Decides what happens to a slot, or reports why it cannot be filled.
 * Shared by validation and rendering so they can never disagree.
 * `source` is the deck slot name the value comes from (see `slotSources`).
 */
export function resolveSlot(
  slot: SlotDef,
  slide: SlideSpec,
  deck: DeckSpec,
  layout: LayoutMode,
  source: string = slot.name,
): { action: SlotAction } | { error: string } {
  let value = slide.slots[source];
  if (value === undefined && slot.kind === 'text' && slot.deckField) {
    value = deck.fields?.[slot.deckField];
  }
  const policy = slot.empty ?? 'error';

  if (value === null || value === undefined || (typeof value === 'string' && value.trim() === '')) {
    const explicit = value === null;
    if (policy === 'keep' && !explicit) return { action: { type: 'keep' } };
    if (policy === 'drop') return { action: { type: 'remove' } };
    if (policy === 'error') return { error: 'this slot is required' };
    if (layout !== 'adaptive') return { error: "leaving this slot empty needs layout: 'adaptive'" };
    return { action: { type: 'remove' } };
  }

  if (slot.kind === 'table') {
    if (!isTableValue(value)) return { error: 'expected a table such as { "header": [...], "rows": [[...]] }' };
    const header = value.header.map((cell) => String(cell ?? '').trim());
    if (header.length === 0) return { error: 'the table has no columns' };
    const rows = value.rows.map((row) => row.map((cell) => String(cell ?? '').trim()));
    const uneven = rows.findIndex((row) => row.length !== header.length);
    if (uneven >= 0) return { error: `row ${uneven + 1} has ${rows[uneven]?.length} cells; the table has ${header.length} columns` };
    return { action: { type: 'table', table: { header, rows } } };
  }

  if (slot.kind === 'chart') {
    if (!isChartValue(value)) {
      return { error: 'expected a chart such as { "type": "column", "categories": [...], "series": [{ "name": "...", "values": [...] }] }' };
    }
    if (!['column', 'bar', 'line', 'pie'].includes(value.type)) return { error: `unknown chart type "${String(value.type)}"` };
    const categories = value.categories.map((category) => String(category ?? '').trim());
    if (categories.length === 0) return { error: 'the chart has no categories' };
    const series = value.series.map((entry) => ({
      name: String(entry?.name ?? '').trim(),
      values: (Array.isArray(entry?.values) ? entry.values : []).map(Number),
    }));
    if (series.length === 0) return { error: 'the chart has no series' };
    for (const entry of series) {
      const label = entry.name ? `series "${entry.name}"` : 'a series';
      if (entry.values.length !== categories.length) {
        return { error: `${label} has ${entry.values.length} values; the chart has ${categories.length} categories` };
      }
      if (entry.values.some((number) => !Number.isFinite(number))) return { error: `${label} has a value that is not a number` };
    }
    if (value.type === 'pie' && series.length > 1) return { error: 'a pie chart takes exactly one series' };
    const chart: ChartValue = { type: value.type, categories, series };
    if (value.format === 'percent' || value.format === 'number') chart.format = value.format;
    return { action: { type: 'chart', chart } };
  }

  if (slot.kind === 'text') {
    if (isImageRef(value) || isTableValue(value) || isChartValue(value)) return { error: 'expected text' };
    const paragraphs = toParagraphs(value).filter((paragraph) => paragraph !== '');
    if (paragraphs.length === 0) return { error: 'this slot is required' };
    return { action: { type: 'text', paragraphs } };
  }

  if (!isImageRef(value)) return { error: 'expected an image reference such as { "asset": "id" }' };
  return { action: { type: 'image', image: value } };
}

/** Checks a deck against a manifest. Errors block rendering; warnings do not. */
export function validateDeck(manifest: TemplateManifest, deck: DeckSpec, options: ValidateOptions = {}): Issue[] {
  const issues: Issue[] = [];
  const layout = options.layout ?? 'strict';
  const assets = options.assets ? new Set(options.assets) : undefined;
  const patterns = new Map<string, PatternDef>(manifest.patterns.map((pattern) => [pattern.id, pattern]));
  const mirrored = (deck.direction ?? manifest.direction) !== manifest.direction;

  if (!Array.isArray(deck.slides) || deck.slides.length === 0) {
    issues.push({ severity: 'error', message: 'the deck has no slides' });
    return issues;
  }

  deck.slides.forEach((slide, index) => {
    const number = index + 1;
    const pattern = patterns.get(slide.pattern);
    if (!pattern) {
      issues.push({
        severity: 'error',
        slide: number,
        message: `unknown pattern "${slide.pattern}" (available: ${[...patterns.keys()].join(', ')})`,
      });
      return;
    }

    const range = itemRange(pattern);
    if (slide.items !== undefined && !range) {
      issues.push({ severity: 'warning', slide: number, message: `pattern "${pattern.id}" has a fixed number of items; "items" is ignored` });
    }
    if (range && slide.items !== undefined && (!Number.isInteger(slide.items) || slide.items < range.min || slide.items > range.max)) {
      issues.push({
        severity: 'error',
        slide: number,
        message: `pattern "${pattern.id}" takes ${range.min} to ${range.max} ${pattern.repeat?.item ?? 'item'}s, not ${slide.items}`,
      });
      return;
    }
    const slots = effectiveSlots(pattern, slide.items);
    const known = new Set(slots.map((slot) => slot.name));
    for (const name of Object.keys(slide.slots ?? {})) {
      if (!known.has(name)) {
        issues.push({ severity: 'warning', slide: number, slot: name, message: `pattern "${pattern.id}" has no such slot; ignored` });
      }
    }

    const sources = slotSources(pattern, mirrored, slide.items);
    for (const slot of slots) {
      // Problems are reported under the name the deck used, which differs across a mirrored sequence.
      const source = sources.get(slot.name) ?? slot.name;
      const resolved = resolveSlot(slot, { ...slide, slots: slide.slots ?? {} }, deck, layout, source);
      if ('error' in resolved) {
        issues.push({ severity: 'error', slide: number, slot: source, message: resolved.error });
        continue;
      }
      const action = resolved.action;
      if (action.type === 'text' && slot.kind === 'text') {
        // Larger deck text leaves less room in the same box.
        const budget = textBudget(slot, deck.textScale);
        const chars = action.paragraphs.reduce((sum, paragraph) => sum + paragraph.length, 0);
        if (budget.maxChars !== undefined && chars > budget.maxChars) {
          issues.push({
            severity: 'warning',
            slide: number,
            slot: source,
            message: `text is ${chars} characters; the slot fits about ${budget.maxChars}`,
          });
        }
        if (budget.maxParagraphs !== undefined && action.paragraphs.length > budget.maxParagraphs) {
          issues.push({
            severity: 'warning',
            slide: number,
            slot: source,
            message: `text has ${action.paragraphs.length} paragraphs; the slot fits ${budget.maxParagraphs}`,
          });
        }
      }
      if (action.type === 'table' && slot.kind === 'table') {
        const { header, rows } = action.table;
        const warn = (message: string) => issues.push({ severity: 'warning', slide: number, slot: source, message });
        if (slot.maxColumns !== undefined && header.length > slot.maxColumns) warn(`the table has ${header.length} columns; the slot fits ${slot.maxColumns}`);
        if (slot.maxRows !== undefined && rows.length > slot.maxRows) warn(`the table has ${rows.length} rows; the slot fits ${slot.maxRows}`);
        const longest = Math.max(0, ...[header, ...rows].flat().map((cell) => cell.length));
        if (slot.maxCellChars !== undefined && longest > slot.maxCellChars) warn(`a cell is ${longest} characters; cells fit about ${slot.maxCellChars}`);
      }
      if (action.type === 'chart' && slot.kind === 'chart') {
        const { categories, series } = action.chart;
        const warn = (message: string) => issues.push({ severity: 'warning', slide: number, slot: source, message });
        if (slot.maxCategories !== undefined && categories.length > slot.maxCategories) {
          warn(`the chart has ${categories.length} categories; the slot fits ${slot.maxCategories}`);
        }
        if (slot.maxSeries !== undefined && series.length > slot.maxSeries) warn(`the chart has ${series.length} series; the slot fits ${slot.maxSeries}`);
      }
      if (action.type === 'image' && assets && !assets.has(action.image.asset)) {
        issues.push({ severity: 'error', slide: number, slot: source, message: `asset "${action.image.asset}" was not provided` });
      }
    }
  });

  return issues;
}
