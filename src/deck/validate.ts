import { bodyTextScale, effectiveSlots, itemRange, slotSources, tableAtScale, templateScaleProblem, textBudget } from '../template/manifest.ts';
import type { Direction, PatternDef, SlotDef, TemplateManifest, TextSlotDef } from '../template/manifest.ts';
import type { FitOptions, TextFit, TextMeasurer } from '../text/measure.ts';
import { readColors, readColorsTold } from '../color/palette.ts';
import { freeSlotDefs, freeSlots, isFreeBlockKind } from '../free/blocks.ts';
import { freeFor } from '../free/resolve.ts';
import type { FreeResolved } from '../free/resolve.ts';
import type { Template } from '../render/render.ts';
import type { IssueCode } from '../messages/codes.ts';
import { said } from '../messages/en.ts';
import type { Told } from '../messages/en.ts';
import { DesignError, withLayouts } from '../template/compose.ts';
import { measurerFor } from '../text/measure.ts';
import { FREE_BLOCK_KINDS, FREE_PATTERN, isChartValue, isFreeSlide, isImageRef, isTableValue, isTextValue, plainText, toParagraphs } from './spec.ts';
import type { ChartValue, DeckSpec, ImageRef, SlideSpec, TableValue } from './spec.ts';

export type LayoutMode = 'strict' | 'adaptive';

export interface Issue {
  severity: 'error' | 'warning';
  /** 1-based slide number in the deck; absent for deck-level issues. */
  slide?: number;
  slot?: string;
  /** The English sentence. */
  message: string;
  /**
   * What the issue is, as a code that is never renamed: every warning has one, and so does every
   * note the chat tools give a person. `formatIssue` makes the sentence in another language from it.
   */
  code?: IssueCode;
  /** The numbers and labels the sentence states, by name (see `IssueValues`). Present with `code`. */
  values?: Record<string, string | number | boolean>;
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

/** The most series a chart takes: one more and PowerPoint does not open the file. */
const MOST_SERIES = 255;
/** The most blocks a free slide holds. It is the number a plan is never read beyond (`blocksOfRuns`): past it a slide is not a slide, and checking it takes minutes. */
const MOST_BLOCKS = 2000;
/**
 * What a deck's `textScale` may be: a factor on body text. A value that is no number would be
 * written into the file as it is (`sz="NaN"`), and PowerPoint does not open such a file. Within
 * these bounds every size is written, and one the template's own factor takes past what
 * PowerPoint opens is held at its edge (`openableSizes`). They are the factors at which 0.5
 * rendered the demo template to a file PowerPoint opens (0.1 and 300 did, 0.05 and 1000 did not),
 * so a deck that worked is not turned away; past them a deck is a mistake, and its text would be
 * too small to read or held at 4000 pt.
 */
const TEXT_SCALE = { least: 0.1, most: 300 };

/**
 * What is wrong with a deck's `textScale`, as a sentence; `undefined` when it is a factor the
 * deck can be made with. Left out, or null as an application stores a number that is not set,
 * it is 1. A number written as text ("1.25") is read as that number, as it always was.
 */
export function textScaleProblem(value: unknown): string | undefined {
  if (value === undefined || value === null) return undefined;
  const scale = typeof value === 'number' ? value : typeof value === 'string' && value.trim() !== '' ? Number(value) : NaN;
  if (Number.isFinite(scale) && scale >= TEXT_SCALE.least && scale <= TEXT_SCALE.most) return undefined;
  const given = typeof value === 'number' ? String(value) : JSON.stringify(value);
  return `the deck's textScale is ${given}; it is a factor on body text between ${TEXT_SCALE.least} and ${TEXT_SCALE.most}, such as 1.25`;
}

/**
 * What makes a deck no deck before any of it can be read, as errors: a deck that is no object, one
 * without slides, a slide that is no object, or its own layouts given as something other than a
 * list of layouts. Empty for a deck of the right shape. Validation and the renderer ask this
 * first, so that such a deck is told what is wrong and does not trip the code that reads it.
 */
export function deckShapeIssues(deck: unknown): Issue[] {
  if (typeof deck !== 'object' || deck === null || Array.isArray(deck)) return [{ severity: 'error', message: 'the deck is no object with a list of slides, such as { "version": 1, "slides": [...] }' }];
  const { slides, layouts } = deck as { slides?: unknown; layouts?: unknown };
  if (!Array.isArray(slides) || slides.length === 0) return [{ severity: 'error', message: 'the deck has no slides' }];
  const issues: Issue[] = [];
  slides.forEach((slide: unknown, index) => {
    if (typeof slide !== 'object' || slide === null || Array.isArray(slide)) {
      issues.push({ severity: 'error', slide: index + 1, message: 'the slide is no object with a pattern and slots, such as { "pattern": "cover", "slots": { ... } }' });
    }
  });
  const isObject = (value: unknown) => typeof value === 'object' && value !== null && !Array.isArray(value);
  if (layouts !== undefined && layouts !== null && !(Array.isArray(layouts) && layouts.every(isObject))) {
    const given = Array.isArray(layouts) ? 'a list that holds something other than layouts' : typeof layouts === 'object' ? 'an object' : `a ${typeof layouts}`;
    issues.push({ severity: 'error', message: `"layouts" is ${given}; it is a list of the deck's own layouts, such as [{ "id": "three-goals", "elements": [...] }]` });
  }
  return issues;
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
  /**
   * Measures text with the template's fonts (see `textMeasurer`). With it, text that does not fit
   * is reported by the lines it runs over, in place of the character count. A slot it cannot
   * measure keeps the character check.
   */
  measure?: TextMeasurer;
}

/**
 * How the text of a slide is measured: at the deck's text size, shrinkable as far as the manifest
 * allows, in a row of as many items as the slide has. Validation, the renderer and the design
 * checks all measure with these, so they cannot disagree.
 */
export function fitOptions(manifest: TemplateManifest, deck: Pick<DeckSpec, 'textScale' | 'direction'>, slide: Pick<SlideSpec, 'items'>, least?: number): FitOptions {
  const options: FitOptions = { textScale: bodyTextScale(manifest, deck.textScale) };
  // Paragraphs that mix scripts are written, and so measured, by the direction the deck reads.
  const direction = deck.direction ?? manifest.direction;
  if (direction) options.direction = direction;
  // `least` is a size in points under which this text is not set, whatever the manifest allows (see `freeTextFloor`).
  if (manifest.shrinkText) options.shrink = least === undefined ? manifest.shrinkText : { ...manifest.shrinkText, minSize: Math.max(manifest.shrinkText.minSize ?? 0, least) };
  if (slide.items !== undefined) options.items = slide.items;
  return options;
}

/**
 * The least size, in points, that a text of a free slide is set at in a deck that asks for larger
 * text: the size it has in the same deck without that wish, as it is drawn there, set smaller to
 * fit where the manifest's `shrinkText` allows. A deck that asks for larger text never gets
 * smaller text for it, and that holds for the renderer's shrinking too: a text that does not fit
 * at the larger size is set smaller down to this size and no further, and is reported when it
 * still does not fit. `undefined` when there is nothing to hold: another slide, a deck without
 * the wish, a template that does not shrink text or was loaded without fonts.
 */
export function freeTextFloor(free: FreeResolved | undefined, slot: TextSlotDef, paragraphs: string[], manifest: TemplateManifest, slide: Pick<SlideSpec, 'items'>, measure: TextMeasurer | undefined, direction?: Direction): number | undefined {
  if (!free?.plain || !measure || !manifest.shrinkText) return undefined;
  // The title and the deck's own texts are the template's shapes: they follow the deck's text size as on any slide.
  const was = free.plain.pattern.slots.find((entry) => entry.name === slot.name);
  if (!was || was.kind !== 'text' || !('styleOf' in was) || was.fontSize === undefined) return undefined;
  const fit = measure.fit(free.plain.pattern, was, paragraphs, fitOptions(manifest, direction ? { textScale: 1, direction } : { textScale: 1 }, slide));
  return fit ? Math.min(fit.size, was.fontSize) : was.fontSize;
}

/**
 * The warning for a text that does not fit its slot, as measured. `undefined` when it fits,
 * which includes fitting once it is set smaller within the manifest's `shrinkText`.
 */
export function overflowMessage(fit: TextFit): string | undefined {
  return overflowTold(fit)?.message;
}

/** The same warning with its code and values. */
export function overflowTold(fit: TextFit): Told | undefined {
  if (fit.fits) return undefined;
  const approximate = fit.approximate === true;
  const shrunk = fit.scale < 1;
  if (fit.overflow === 0) return said('text-too-wide', { approximate, percent: Math.round((fit.widthRatio - 1) * 100), shrunk, size: fit.size, fitsChars: fit.fitsChars });
  return said('text-overflow', { approximate, lines: fit.lines, shrunk, size: fit.size, capacity: fit.capacity, over: fit.overflow, fitsChars: fit.fitsChars });
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
    const header = value.header.map((cell) => plainText(cell).trim());
    if (header.length === 0) return { error: 'the table has no columns' };
    const rows = value.rows.map((row) => row.map((cell) => plainText(cell).trim()));
    const uneven = rows.findIndex((row) => row.length !== header.length);
    if (uneven >= 0) return { error: `row ${uneven + 1} has ${rows[uneven]?.length} cells; the table has ${header.length} columns` };
    return { action: { type: 'table', table: { header, rows } } };
  }

  if (slot.kind === 'chart') {
    if (!isChartValue(value)) {
      return { error: 'expected a chart such as { "type": "column", "categories": [...], "series": [{ "name": "...", "values": [...] }] }' };
    }
    if (!['column', 'bar', 'line', 'pie'].includes(value.type)) return { error: `unknown chart type "${String(value.type)}"` };
    const categories = value.categories.map((category) => plainText(category).trim());
    if (categories.length === 0) return { error: 'the chart has no categories' };
    const series = value.series.map((entry) => ({
      name: plainText(entry?.name).trim(),
      values: (Array.isArray(entry?.values) ? entry.values : []).map(Number),
    }));
    if (series.length === 0) return { error: 'the chart has no series' };
    // PowerPoint does not open a file that holds a chart of more series than this.
    if (series.length > MOST_SERIES) return { error: `the chart has ${series.length} series; a chart takes at most ${MOST_SERIES}` };
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
    // A number, an object or a list that holds one is no text, wherever it came from.
    if (!isTextValue(value)) return { error: 'expected text' };
    const paragraphs = toParagraphs(value).filter((paragraph) => paragraph !== '');
    if (paragraphs.length === 0) return { error: 'this slot is required' };
    return { action: { type: 'text', paragraphs } };
  }

  if (!isImageRef(value)) return { error: 'expected an image reference such as { "asset": "id" }' };
  return { action: { type: 'image', image: value.alt === undefined ? value : { ...value, alt: plainText(value.alt) } } };
}

/**
 * Checks a deck. Errors block rendering; warnings do not.
 *
 * Give it the template, and it says exactly what `renderDeck` would say: the deck's own layouts
 * are built in, free slides are laid out and their text is checked against the room it gets, and
 * a template loaded with fonts measures. Given only a manifest it checks what a manifest can
 * tell: for a free slide, that every block has content of the right kind, and nothing about fit.
 */
export function validateDeck(target: TemplateManifest | Template, deck: DeckSpec, options: ValidateOptions = {}): Issue[] {
  if (!('pkg' in target)) return checkDeck(target, deck, options, undefined);
  // A deck of the wrong shape is told so before its layouts are built.
  const shape = deckShapeIssues(deck);
  if (shape.length > 0) return shape;
  let built: Template;
  try {
    built = withLayouts(target, deck.layouts);
  } catch (error) {
    // A deck whose own layout no longer fits the template cannot be rendered; that is an answer, not a crash.
    if (error instanceof DesignError) return [{ severity: 'error', message: error.message }];
    throw error;
  }
  const measure = options.measure ?? measurerFor(built);
  return checkDeck(built.manifest, deck, measure ? { ...options, measure } : options, (slide) => freeFor(target, deck, slide));
}

/** The check itself. `free` lays a free slide out; without it only what needs no template is checked. */
export function checkDeck(manifest: TemplateManifest, deck: DeckSpec, options: ValidateOptions, free: ((slide: SlideSpec) => FreeResolved | string) | undefined): Issue[] {
  const issues: Issue[] = [];
  // A deck that is no object, has no slides, or a slide that is no object: nothing more of it can be read.
  const shape = deckShapeIssues(deck);
  if (shape.length > 0) return shape;
  const layout = options.layout ?? 'strict';
  const assets = options.assets ? new Set(options.assets) : undefined;
  const patterns = new Map<string, PatternDef>(manifest.patterns.map((pattern) => [pattern.id, pattern]));
  const mirrored = (deck.direction ?? manifest.direction) !== manifest.direction;
  // A text size that is no size stops the deck: every size on every slide follows from it.
  const scaleProblem = templateScaleProblem(manifest.textScale) ?? textScaleProblem(deck.textScale);
  if (scaleProblem) {
    issues.push({ severity: 'error', message: scaleProblem });
    return issues;
  }

  // A colour that is not a colour is left out and said so. It never stops a deck, and it is told
  // without a slide number, so that a planner does not send the slide back to be written again.
  for (const problem of readColorsTold(deck.colors).told) issues.push({ severity: 'warning', ...problem });
  deck.slides.forEach((slide, index) => {
    for (const problem of readColorsTold(slide.colors, index + 1).told) issues.push({ severity: 'warning', ...problem });
  });
  if (manifest.colors === false && (readColors(deck.colors).colors || deck.slides.some((slide) => readColors(slide.colors).colors))) {
    issues.push({ severity: 'warning', ...said('colors-kept-by-template', {}) });
  }

  deck.slides.forEach((slide, index) => {
    const number = index + 1;
    let pattern = patterns.get(slide.pattern);
    // What a slide may leave out is its layout mode's to say; on a free slide a part without content is simply not drawn.
    let mode = layout;
    let parts: Set<string> | undefined;
    let laid: FreeResolved | undefined;
    if (isFreeSlide(slide)) {
      const blocks = slide.blocks as unknown[];
      if (blocks.length > MOST_BLOCKS) {
        issues.push({ severity: 'error', slide: number, message: `a free slide holds at most ${MOST_BLOCKS} blocks, not ${blocks.length}` });
        return;
      }
      const unknown = blocks.filter((kind) => !isFreeBlockKind(kind));
      if (blocks.length === 0 || unknown.length > 0) {
        issues.push({
          severity: 'error',
          slide: number,
          message: blocks.length === 0 ? 'a free slide needs at least one block, such as "blocks": ["text"]' : `"blocks" names ${unknown.map((kind) => JSON.stringify(kind)).join(', ')}; a free slide is made of ${FREE_BLOCK_KINDS.join(', ')}`,
        });
        return;
      }
      const kinds = blocks.filter(isFreeBlockKind);
      mode = 'adaptive';
      parts = new Set(['title', ...freeSlots(kinds).map((entry) => entry.name)]);
      const columns = slide.arrange?.columns;
      if (columns !== undefined && !(Number.isInteger(columns) && columns >= 1)) {
        issues.push({ severity: 'warning', ...said('columns-not-whole', { slide: number, given: String(JSON.stringify(columns)) }) });
      }
      if (free) {
        const resolved = free(slide);
        if (typeof resolved === 'string') {
          issues.push({ severity: 'error', slide: number, message: `a free slide cannot be made from this template: ${resolved}` });
          return;
        }
        laid = resolved;
        pattern = resolved.plan.pattern;
        // The slide and its number of blocks, which the notes on a free slide begin with.
        const holds = { slide: number, blocks: kinds.length };
        // A number of blocks side by side that would have left some block under its smallest size was not followed.
        const relaxed = resolved.plan.layout.relaxed;
        if (relaxed) {
          issues.push({ severity: 'warning', ...said('columns-relaxed', { slide: number, asked: relaxed.asked, columns: relaxed.columns }) });
        }
        // Told without a slide number: no rewriting of text changes how many blocks a slide was asked to hold.
        if (resolved.plan.smallest < resolved.plan.comfortable) {
          // Sizes are worked out in fractions of a point; the file and the reader get half points.
          issues.push({ severity: 'warning', ...said('free-text-small', { ...holds, size: Math.round(resolved.plan.smallest * 2) / 2 }) });
        }
        // A deck's larger text is never bought with smaller text: a slide on which some text would
        // come out smaller than it is without the wish takes half of it, or none, and that is said,
        // not done in silence. Text that is no smaller and no longer fits its box at the larger size
        // is another matter: the slide keeps the larger text, and that text is reported on its slot.
        if (resolved.enlarged) {
          issues.push({ severity: 'warning', ...said('free-larger-text', { ...holds, half: resolved.enlarged === 'half' }) });
        }
        const bare = resolved.plan.pattern.slots.filter((slot) => slot.kind === 'chart' && slot.detail === 'spark').length;
        if (bare > 0) {
          issues.push({ severity: 'warning', ...said('free-charts-bare', { ...holds, charts: bare }) });
        }
        // Nothing is laid out below the room a slide has. Should a block get there all the same, the deck is told.
        const low = resolved.plan.layout.blocks.filter((block) => block.box.y + block.box.h > resolved.kit.dense.bottom + 0.002).length;
        if (low > 0) {
          issues.push({ severity: 'warning', ...said('free-blocks-low', { ...holds, count: low }) });
        }
        if (slide.base !== undefined && resolved.kit.base !== slide.base) {
          // The layout is gone, or it is there and cannot be a frame: then the reason is said.
          const why = resolved.refused;
          const frame = { slide: number, base: String(slide.base), stands: resolved.kit.base };
          issues.push({ severity: 'warning', ...(why ? said('free-base-refused', { ...frame, why }) : said('free-base-gone', frame)) });
        }
      } else {
        pattern = { id: FREE_PATTERN, sourceSlide: 0, description: '', slots: freeSlotDefs(kinds) };
      }
    }
    if (!pattern) {
      issues.push({
        severity: 'error',
        slide: number,
        message: deck.layouts?.some((design) => design.id === slide.pattern)
          ? `pattern "${slide.pattern}" is one of the deck's own layouts: check the deck against withLayouts(template, deck.layouts)`
          : `unknown pattern "${slide.pattern}" (available: ${[...patterns.keys()].join(', ')})`,
      });
      return;
    }

    const range = itemRange(pattern);
    if (slide.items !== undefined && !range) {
      issues.push({ severity: 'warning', slide: number, ...said('items-ignored', { pattern: pattern.id }) });
    }
    if (range && slide.items !== undefined && (!Number.isInteger(slide.items) || slide.items < range.min || slide.items > range.max)) {
      issues.push({
        severity: 'error',
        slide: number,
        message: `pattern "${pattern.id}" takes ${range.min === range.max ? `exactly ${range.min}` : `${range.min} to ${range.max}`} ${pattern.repeat?.item ?? 'item'}s, not ${slide.items}`,
      });
      return;
    }
    const slots = effectiveSlots(pattern, slide.items);
    const known = new Set(slots.map((slot) => slot.name));
    for (const name of Object.keys(slide.slots ?? {})) {
      if (!known.has(name) && !parts?.has(name)) {
        issues.push({ severity: 'warning', slide: number, slot: name, ...said('slot-unknown', { pattern: pattern.id }) });
      }
    }

    const sources = slotSources(pattern, mirrored, slide.items);
    for (const slot of slots) {
      // Problems are reported under the name the deck used, which differs across a mirrored sequence.
      const source = sources.get(slot.name) ?? slot.name;
      const resolved = resolveSlot(slot, { ...slide, slots: slide.slots ?? {} }, deck, mode, source);
      if ('error' in resolved) {
        issues.push({ severity: 'error', slide: number, slot: source, message: resolved.error });
        continue;
      }
      const action = resolved.action;
      if (action.type === 'text' && slot.kind === 'text') {
        // Larger text, asked for by the template or the deck, leaves less room in the same box.
        const scale = bodyTextScale(manifest, deck.textScale);
        const budget = textBudget(slot, scale);
        const chars = action.paragraphs.reduce((sum, paragraph) => sum + paragraph.length, 0);
        // With fonts the text is measured as it will be set, and the character budget steps aside.
        const fit = options.measure?.fit(pattern, slot, action.paragraphs, fitOptions(manifest, deck, slide, freeTextFloor(laid, slot, action.paragraphs, manifest, slide, options.measure, deck.direction)));
        if (fit) {
          const over = overflowTold(fit);
          if (over) issues.push({ severity: 'warning', slide: number, slot: source, ...over });
        } else if (budget.maxChars !== undefined && chars > budget.maxChars) {
          issues.push({ severity: 'warning', slide: number, slot: source, ...said('slot-over-chars', { chars, most: budget.maxChars }) });
        }
        if (budget.maxParagraphs !== undefined && action.paragraphs.length > budget.maxParagraphs) {
          issues.push({ severity: 'warning', slide: number, slot: source, ...said('slot-over-paragraphs', { paragraphs: action.paragraphs.length, most: budget.maxParagraphs }) });
        }
      }
      const warn = (told: Told) => issues.push({ severity: 'warning', slide: number, slot: source, ...told });
      if (action.type === 'table' && slot.kind === 'table') {
        const { header, rows } = action.table;
        if (slot.maxColumns !== undefined && header.length > slot.maxColumns) warn(said('table-over-columns', { columns: header.length, most: slot.maxColumns }));
        if (slot.maxRows !== undefined && rows.length > slot.maxRows) warn(said('table-over-rows', { rows: rows.length, most: slot.maxRows }));
        // Counted in a loop: a table of a hundred thousand cells is reported like any other, and is not more arguments than a call takes.
        let longest = 0;
        for (const row of [header, ...rows]) for (const cell of row) if (cell.length > longest) longest = cell.length;
        // Enlarged table text leaves less room in a cell.
        const cellChars = tableAtScale(slot, bodyTextScale(manifest, deck.textScale)).maxCellChars;
        if (cellChars !== undefined && longest > cellChars) warn(said('table-cell-long', { chars: longest, most: cellChars }));
      }
      if (action.type === 'chart' && slot.kind === 'chart') {
        const { categories, series } = action.chart;
        if (slot.maxCategories !== undefined && categories.length > slot.maxCategories) {
          warn(said('chart-over-categories', { categories: categories.length, most: slot.maxCategories }));
        }
        if (slot.maxSeries !== undefined && series.length > slot.maxSeries) warn(said('chart-over-series', { series: series.length, most: slot.maxSeries }));
        // A format the renderer does not know is shown as plain numbers, which is worth saying.
        const format = (slide.slots[source] as { format?: unknown } | null | undefined)?.format;
        if (format !== undefined && format !== 'number' && format !== 'percent') {
          warn(said('chart-format-unknown', { format: String(JSON.stringify(format)) }));
        }
      }
      if (action.type === 'image' && assets && !assets.has(action.image.asset)) {
        issues.push({ severity: 'error', slide: number, slot: source, message: `asset "${action.image.asset}" was not provided` });
      }
    }
  });

  return issues;
}
