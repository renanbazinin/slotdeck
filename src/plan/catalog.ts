import { FREE_PATTERN } from '../deck/spec.ts';
import type { FreeBlockKind } from '../deck/spec.ts';
import type { LayoutMode } from '../deck/validate.ts';
import type { JsonSchema } from '../llm/provider.ts';
import { defaultItems, effectiveSlots, itemRange, rowLayout, tableAtScale, textBudget } from '../template/manifest.ts';
import type { ChartSlotDef, ImageSlotDef, PatternDef, SlotDef, TableSlotDef, TemplateManifest, TextSlotDef } from '../template/manifest.ts';
import { askedSlidesProperty, documentLanguageProperty, leftOutProperty, passagesProperty } from './document.ts';
import { MOST_WORDS, WORDS_NOTE } from './length.ts';

/** An icon the app offers; the planning model picks from these by id. */
export interface IconDef {
  id: string;
  /** What the icon shows, e.g. "clock" or "handshake". */
  description: string;
  data: Uint8Array;
}

/** A ready-made picture the app supplies, e.g. a rendered Mermaid diagram or a product photo. */
export interface AssetDef {
  id: string;
  kind: 'photo' | 'graphic';
  /** What the picture shows; the planning model decides where it fits from this. */
  description: string;
  data: Uint8Array;
}

export type ImageDensity = 'none' | 'low' | 'medium' | 'high';

export interface CatalogContext {
  layout: LayoutMode;
  /** Deck-level values the caller already has. */
  fields: Record<string, string>;
  icons: IconDef[];
  assets: AssetDef[];
  imageDensity: ImageDensity;
  /** Whether patterns built around numbers may be offered. */
  figures: boolean;
  /** Size of the deck's body text as a factor on the template's; larger text gets smaller budgets. */
  textScale: number;
  /** The template's own factor on body text, applied on top. Defaults to 1. */
  templateScale?: number;
  /** Restrict planning to these pattern ids. */
  allow?: string[];
  /**
   * Ids of the layouts made for this deck (`DeckSpec.layouts`). They were made for its content,
   * so they are offered whatever the brief holds, and the list says they were made for the deck.
   */
  own?: string[];
}

/** How one manifest slot is presented to the planning model. */
export type PlannedSlot =
  | { slot: TextSlotDef; mode: 'text' | 'paragraphs'; required: boolean }
  | { slot: ImageSlotDef; mode: 'photo' | 'icon' | 'graphic'; required: boolean }
  | { slot: TableSlotDef; mode: 'table'; required: boolean; mayBeEmpty?: boolean }
  | { slot: ChartSlotDef; mode: 'chart'; required: boolean; mayBeEmpty?: boolean };

export interface UsablePattern {
  def: PatternDef;
  /** The slots the model fills. Slots fed by deck fields, or left to the template, are not listed. */
  slots: PlannedSlot[];
  /** Whether using the pattern always puts a photo on the slide. */
  needsPhoto: boolean;
  /** For a pattern with a row of items: the number of items these slots are for. */
  items?: number;
  /** Whether it is one of the layouts made for this deck. */
  own?: boolean;
}

export interface Catalog {
  patterns: UsablePattern[];
  /** Patterns that cannot be used with this context, and why. */
  excluded: Array<{ id: string; reason: string }>;
}

/**
 * Works out which slots of a pattern the model writes, or why the pattern cannot be used.
 * Slots named in `skip` are already taken care of (carried over from an existing slide)
 * and are neither offered nor required.
 */
export function planPattern(
  def: PatternDef,
  context: CatalogContext,
  skip: ReadonlySet<string> = new Set(),
  items?: number,
): { pattern: UsablePattern } | { reason: string } {
  const own = context.own?.includes(def.id) ?? false;
  if (def.figures && !context.figures && !own) return { reason: 'is built around numbers, and the brief gives none' };
  const hasGraphics = context.assets.some((asset) => asset.kind === 'graphic');
  const slots: PlannedSlot[] = [];
  let reason: string | undefined;
  let needsPhoto = false;

  for (const slot of effectiveSlots(def, items)) {
    if (skip.has(slot.name)) continue;
    const policy = slot.empty ?? 'error';
    // In strict layout nothing may be removed, so a removable slot has to be filled like any other.
    const mustFill = policy === 'error' || (policy === 'remove' && context.layout === 'strict');
    const mayFill = policy !== 'keep' && policy !== 'drop';

    if (slot.kind === 'text') {
      if (slot.deckField) {
        if (!context.fields[slot.deckField] && mustFill) reason = `needs the deck field "${slot.deckField}"`;
        continue;
      }
      if (!mayFill) continue;
      slots.push({ slot, mode: (slot.maxParagraphs ?? 1) > 1 ? 'paragraphs' : 'text', required: mustFill });
      continue;
    }

    if (slot.kind === 'table') {
      if (mayFill) slots.push({ slot, mode: 'table', required: mustFill });
      continue;
    }

    if (slot.kind === 'chart') {
      if (mayFill) slots.push({ slot, mode: 'chart', required: mustFill });
      continue;
    }

    let available: boolean;
    if (slot.imageKind === 'photo') available = context.imageDensity !== 'none';
    else if (slot.imageKind === 'icon') available = context.icons.length > 0;
    else available = hasGraphics;

    if (!available || !mayFill) {
      if (mustFill) {
        reason =
          slot.imageKind === 'photo'
            ? 'needs a photo, and images are turned off'
            : slot.imageKind === 'icon'
              ? 'needs icons, and no icon set was supplied'
              : 'needs a graphic asset, and none was supplied';
      }
      continue;
    }
    if (slot.imageKind === 'photo' && mustFill) needsPhoto = true;
    slots.push({ slot, mode: slot.imageKind, required: mustFill });
  }

  if (reason) return { reason };
  const pattern: UsablePattern = { def, slots, needsPhoto };
  if (def.repeat && items !== undefined) pattern.items = items;
  if (own) pattern.own = true;
  return { pattern };
}

/** The factor that budgets are worked out with: the deck's text size on top of the template's. */
export const budgetScale = (context: CatalogContext): number => context.textScale * (context.templateScale ?? 1);

/**
 * Whether a slide made from the pattern is running text and nothing else: a title with paragraphs
 * or a list. Cards, figures, tables, charts, icons and pictures all make a pattern something more.
 * A picture that the pattern always drops does not count.
 */
export function isPlainText(def: PatternDef): boolean {
  return !def.repeat && !def.figures && def.slots.every((slot) => slot.kind === 'text' || slot.empty === 'drop');
}

/** Works out which patterns can be filled with what the caller supplied, and which of their slots the model writes. */
export function buildCatalog(manifest: TemplateManifest, context: CatalogContext): Catalog {
  const patterns: UsablePattern[] = [];
  const excluded: Catalog['excluded'] = [];
  for (const def of manifest.patterns) {
    if (context.allow && !context.allow.includes(def.id)) continue;
    const planned = planPattern(def, context);
    if ('reason' in planned) excluded.push({ id: def.id, reason: planned.reason });
    else patterns.push(planned.pattern);
  }
  return { patterns, excluded };
}

/** A limit in words: "at most 4 rows", and "at most 1 row" for one. */
const atMost = (count: number, one: string, many: string): string => `at most ${count} ${count === 1 ? one : many}`;

function limits(given: SlotDef, scale = 1): string {
  const slot = given.kind === 'table' ? tableAtScale(given, scale) : given;
  if (slot.kind === 'table') {
    const parts: string[] = [];
    if (slot.maxColumns !== undefined) parts.push(atMost(slot.maxColumns, 'column', 'columns'));
    if (slot.maxRows !== undefined) parts.push(atMost(slot.maxRows, 'row', 'rows'));
    if (slot.maxCellChars !== undefined) parts.push(`${atMost(slot.maxCellChars, 'character', 'characters')} in a cell`);
    return parts.join(', ');
  }
  if (slot.kind === 'chart') {
    const parts: string[] = [];
    if (slot.maxCategories !== undefined) parts.push(atMost(slot.maxCategories, 'category', 'categories'));
    if (slot.maxSeries !== undefined) parts.push(atMost(slot.maxSeries, 'series', 'series'));
    return parts.join(', ');
  }
  if (slot.kind !== 'text') return '';
  const budget = textBudget(slot, scale);
  const several = (slot.maxParagraphs ?? 1) > 1;
  const parts: string[] = [];
  if (budget.maxChars !== undefined) parts.push(`${atMost(budget.maxChars, 'character', 'characters')}${several ? ' in total' : ''}`);
  if (several && budget.maxParagraphs !== undefined) parts.push(atMost(budget.maxParagraphs, 'item', 'items'));
  return parts.join(', ');
}

/** The pattern list as the outline prompt shows it. */
export function describeCatalog(catalog: Catalog, textScale = 1): string {
  return catalog.patterns
    .map((pattern) => {
      const repeat = pattern.def.repeat;
      const range = itemRange(pattern.def);
      // A row of items is described once, by its first item, with the number it may hold.
      const itemSlot = new Map<string, { n: number; label: string }>();
      for (let n = 1; repeat && n <= (range?.drawn ?? 0); n++) {
        for (const name of repeat.slots) itemSlot.set(name.replace('{n}', String(n)), { n, label: name.replace('{n}', 'N') });
      }
      const slots = pattern.slots
        .filter((planned) => (itemSlot.get(planned.slot.name)?.n ?? 1) === 1)
        .map((planned) => {
          const { slot } = planned;
          const kind =
            planned.mode === 'text' || planned.mode === 'paragraphs'
              ? limits(slot, textScale) || 'text'
              : planned.mode === 'table' || planned.mode === 'chart'
                ? [planned.mode, limits(slot, textScale)].filter(Boolean).join(': ')
                : planned.mode;
          return `${itemSlot.get(slot.name)?.label ?? slot.name}${planned.required ? '' : '?'} (${kind})`;
        })
        .join('; ');
      // A row that widens its items gives each of fewer items more room; the limits shown are those of the drawn number.
      const roomier = repeat?.widen && repeat.widen > 1 && range && range.min < (repeat.columns ?? range.drawn) ? ` With fewer on a line than ${repeat.columns ?? range.drawn}, each ${repeat.item} is wider and holds more text.` : '';
      // How the items sit is said too: a wish for columns side by side is met by a row, not by a
      // grid that breaks them over lines. A grid that never takes more than one line holds is a row.
      const placed =
        !repeat || repeat.columns === undefined || !range || range.max <= repeat.columns
          ? 'side by side in one row'
          : range.max === range.min
            ? `in lines of ${Math.max(...rowLayout(repeat, range.min))}`
            : `in lines of up to ${repeat.columns}`;
      // A row that takes one number of items says so too: its description gives no number.
      const count = !repeat || !range
        ? ''
        : range.max > range.min
          ? `\n  items: ${range.min} to ${range.max} ${repeat.item}s ${placed}, numbered N = 1, 2, ...${roomier}`
          : range.min > 1
            ? `\n  items: always ${range.min} ${repeat.item}s ${placed}, numbered N = 1, ..., ${range.min}`
            : `\n  items: always 1 ${repeat.item}, numbered N = 1`;
      // Said in so many words, so that a wish for text-only slides can be met without guessing from the slots.
      const plain = isPlainText(pattern.def) ? ' Plain text.' : '';
      const own = pattern.own ? ' Made for this deck.' : '';
      return `- ${pattern.def.id}: ${pattern.def.description}${plain}${own}\n  slots: ${slots}${count}`;
    })
    .join('\n');
}

/** What a slot's schema says of it: its description, then its limits at the deck's text size. */
export function slotNote(slot: SlotDef, context: CatalogContext): string {
  return [slot.description, limits(slot, budgetScale(context))].filter(Boolean).join('. ');
}

/** Said with every chart a model is asked for. */
export const CHART_RULE = 'A chart of numbers the brief states. Use only those numbers: never estimate, interpolate or invent a value to complete a series.';
export const CHART_TYPE_NOTE = 'column: compare items. bar: compare items with long names. line: change over time. pie: shares of a whole.';
export const CHART_FORMAT_NOTE = 'percent when the values are percentages (write 45 for 45%)';
export const PHOTO_PROMPT_NOTE = 'What the photo shows, in English: subject, setting, light and mood, in one or two sentences. No text, logos or brand names in the picture.';
export const PHOTO_ALT_NOTE = 'Short alternative text for screen readers, in the deck language';
export const PHOTO_ASSET_NOTE = 'Set only when one of the supplied photos fits; it is then used instead of generating one.';
/**
 * Said of the lists of a chart and of a table that may be left empty, in a slide written with the
 * deck. Such a part is a property every answer has to hold, so "leave it empty" needs a way to
 * do it. Without one, "one or more series" reads as an order to find some.
 */
const EMPTY_CHART_NOTE = '[] for a chart that is left empty';
const EMPTY_TABLE_NOTE = '[] for a table that is left empty';

function slotSchema(planned: PlannedSlot, context: CatalogContext): JsonSchema {
  const { slot } = planned;
  const note = slotNote(slot, context);

  switch (planned.mode) {
    case 'text':
      return { type: 'string', description: note || 'Text' };
    case 'paragraphs': {
      const schema: JsonSchema = { type: 'array', items: { type: 'string' }, minItems: 1, description: note || 'One string per paragraph' };
      const most = textBudget(planned.slot, budgetScale(context)).maxParagraphs;
      if (most !== undefined) schema.maxItems = most;
      return schema;
    }
    case 'table': {
      // A table or a chart that may be left empty (a part of a free slide the brief has nothing for) states no least number.
      // Its description says when to leave it empty, and the two lists say how: the property itself is asked of every answer.
      const none = planned.mayBeEmpty ? EMPTY_TABLE_NOTE : '';
      const cells: JsonSchema = { type: 'array', items: { type: 'string' }, ...(planned.mayBeEmpty ? {} : { minItems: 2 }) };
      const rows: JsonSchema = { type: 'array', ...(planned.mayBeEmpty ? {} : { minItems: 1 }), description: `Body rows. Every row has one cell per column, in column order.${none ? ` ${none}.` : ''}` };
      if (planned.slot.maxColumns !== undefined) cells.maxItems = planned.slot.maxColumns;
      if (planned.slot.maxRows !== undefined) rows.maxItems = planned.slot.maxRows;
      rows.items = { ...cells };
      return {
        type: 'object',
        description: note || 'A table',
        properties: { header: { ...cells, description: `Column headings, a word or two each${none ? `. ${none}` : ''}` }, rows },
        required: ['header', 'rows'],
      };
    }
    case 'chart': {
      const none = planned.mayBeEmpty ? EMPTY_CHART_NOTE : '';
      const categories: JsonSchema = { type: 'array', items: { type: 'string' }, ...(planned.mayBeEmpty ? {} : { minItems: 2 }), description: `Labels along the axis, or the names of the pie slices${none ? `. ${none}` : ''}` };
      const series: JsonSchema = {
        type: 'array',
        ...(planned.mayBeEmpty ? {} : { minItems: 1 }),
        description: none ? `One or more series. A pie takes exactly one. ${none}.` : 'One or more series. A pie takes exactly one.',
        items: {
          type: 'object',
          properties: {
            name: { type: 'string', description: 'What the series measures, a word or two' },
            values: { type: 'array', items: { type: 'number' }, description: 'One number per category, in category order' },
          },
          required: ['name', 'values'],
        },
      };
      if (planned.slot.maxCategories !== undefined) categories.maxItems = planned.slot.maxCategories;
      if (planned.slot.maxSeries !== undefined) series.maxItems = planned.slot.maxSeries;
      return {
        type: 'object',
        description: (note ? `${note}. ` : '') + CHART_RULE,
        properties: {
          type: { type: 'string', enum: ['column', 'bar', 'line', 'pie'], description: CHART_TYPE_NOTE },
          categories,
          series,
          format: { type: 'string', enum: ['number', 'percent'], description: CHART_FORMAT_NOTE },
        },
        required: ['type', 'categories', 'series', 'format'],
      };
    }
    case 'icon':
      return { type: 'string', enum: context.icons.map((icon) => icon.id), description: note || 'Id of the icon that best matches the item' };
    case 'graphic':
      return {
        type: 'string',
        enum: context.assets.filter((asset) => asset.kind === 'graphic').map((asset) => asset.id),
        description: note || 'Id of the supplied graphic to show here',
      };
    case 'photo': {
      const properties: Record<string, JsonSchema> = {
        prompt: { type: 'string', description: PHOTO_PROMPT_NOTE },
        alt: { type: 'string', description: PHOTO_ALT_NOTE },
      };
      const photos = context.assets.filter((asset) => asset.kind === 'photo');
      if (photos.length > 0) {
        properties.asset = { type: 'string', enum: photos.map((asset) => asset.id), description: PHOTO_ASSET_NOTE };
      }
      return { type: 'object', properties, required: ['prompt', 'alt'], description: note || 'Photo for this slide' };
    }
  }
}

/** One slide the model is asked to write. */
export interface WriteItem {
  /** 1-based position of the slide in the finished deck. */
  number: number;
  pattern: UsablePattern;
  /** What the slide is for, or what to change on it. */
  purpose: string;
}

/**
 * Schema for the content of a fixed set of slides: one property per slide (`s1`, `s2`, ...),
 * each holding exactly that pattern's slots. Having no unions keeps it small and easy to enforce.
 */
export function contentSchema(items: WriteItem[], context: CatalogContext): JsonSchema {
  const properties: Record<string, JsonSchema> = {};
  for (const { number, pattern, purpose } of items) {
    const slotProperties: Record<string, JsonSchema> = {};
    for (const planned of pattern.slots) slotProperties[planned.slot.name] = slotSchema(planned, context);
    properties[`s${number}`] = {
      type: 'object',
      description: `Slide ${number} (${pattern.def.id}): ${purpose}`,
      properties: slotProperties,
      required: pattern.slots.filter((planned) => planned.required).map((planned) => planned.slot.name),
    };
  }
  return { type: 'object', properties, required: Object.keys(properties) };
}

/** Where the first part of a free slide stands, as a plan or a change to a deck states it. `auto` leaves it to the software. */
export const SHOWS_MAIN = ['auto', 'left', 'right', 'top', 'none'] as const;
/** Said of the count of a run of like parts. */
export const SHOWS_COUNT_NOTE = 'How many parts of this kind stand here, one after another: 8 for eight charts. One entry for them all, not one entry for each.';
export const SHOWS_MAIN_NOTE = 'Only for free: where the first part stands when it is the main one and larger. auto lets the software decide; none sets all parts alike.';

/**
 * Schema of one slide of an outline: its pattern, the point it makes and, where a row of items
 * varies, how many it has. A new plan and a change to a plan ask for the same entry, so that what
 * one can say about a slide the other can too. With `free`, a slide may also be a free slide,
 * and the entry says what it is made of; without it the entry is what it was before free slides.
 * With `document`, a deck made from a document of so many passages, each slide says which of them
 * it carries.
 */
export function outlineEntrySchema(catalog: Catalog, free?: FreeShows, document?: { passages: number }): JsonSchema {
  const ids = catalog.patterns.map((pattern) => pattern.def.id);
  const purpose: JsonSchema = { type: 'string', description: 'The one point this slide makes: a single sentence of at most 25 words' };
  // An amount of text asked for a slide. Optional: most slides have none, and are planned as before.
  const words: JsonSchema = { type: 'integer', minimum: 1, maximum: MOST_WORDS, description: WORDS_NOTE };
  // The passages of a document a slide carries, asked of every slide: [] says it carries none.
  const passages: Record<string, JsonSchema> = document ? { passages: passagesProperty(document.passages) } : {};
  const carries = document ? ['passages'] : [];
  if (!free) return { type: 'object', properties: { pattern: { type: 'string', enum: ids }, purpose, ...itemsProperty(catalog), words, ...passages }, required: ['pattern', 'purpose', ...carries] };

  // A free slide names its parts where a pattern has its slots. The three properties are asked of
  // every slide, so that a model cannot choose `free` and leave its parts out.
  const items = itemsProperty(catalog);
  if (items.items) items.items = { ...items.items, description: `${items.items.description} Not for free: its parts go in "shows".` };
  return {
    type: 'object',
    properties: {
      pattern: { type: 'string', enum: [...ids, FREE_PATTERN] },
      purpose,
      ...items,
      shows: {
        type: 'array',
        description: 'Only for the layout free: its parts in reading order, like parts together. [] for every other layout.',
        items: {
          type: 'object',
          properties: { count: { type: 'integer', minimum: 1, description: SHOWS_COUNT_NOTE }, kind: { type: 'string', enum: [...free.kinds] } },
          required: ['count', 'kind'],
        },
      },
      shows_main: { type: 'string', enum: [...SHOWS_MAIN], description: SHOWS_MAIN_NOTE },
      side_by_side: { type: 'integer', minimum: 0, description: 'Only for free, when the brief asks for a number of columns; 0 lets the software decide.' },
      words,
      ...passages,
    },
    required: ['pattern', 'purpose', 'shows', 'shows_main', 'side_by_side', ...carries],
  };
}

/** What an outline entry may name as the parts of a free slide. Never empty. */
export interface FreeShows {
  kinds: readonly FreeBlockKind[];
}

/**
 * Schema of an outline. With `document`, a deck made from a document of so many passages: the
 * answer also names the passages it leaves out, where the brief gave no number of slides
 * (`askedSlides`) the number of slides the brief asks for in its words, and the language the
 * document is written in.
 */
export function outlineSchema(catalog: Catalog, count: { min: number; max: number }, free?: FreeShows, document?: { passages: number; askedSlides: boolean }): JsonSchema {
  const more: Record<string, JsonSchema> = document
    ? { left_out: leftOutProperty(document.passages, 'brief'), ...(document.askedSlides ? { asked_slides: askedSlidesProperty('brief') } : {}), document_language: documentLanguageProperty }
    : {};
  return {
    type: 'object',
    properties: {
      title: { type: 'string', description: 'Title of the deck' },
      slides: {
        type: 'array',
        minItems: count.min,
        maxItems: count.max,
        items: outlineEntrySchema(catalog, free, document ? { passages: document.passages } : undefined),
      },
      ...more,
    },
    required: ['title', 'slides', ...Object.keys(more)],
  };
}

/** The `items` property of an outline entry, when some pattern on offer has a row of items to vary. */
export function itemsProperty(catalog: Catalog): Record<string, JsonSchema> {
  const ranges = catalog.patterns.map((pattern) => itemRange(pattern.def)).filter((range) => range !== undefined);
  if (!ranges.some((range) => (range as { min: number; max: number }).max > (range as { min: number }).min)) return {};
  // A row that takes one number is counted in too: the list names that number, so it must be allowed.
  return {
    items: {
      type: 'integer',
      minimum: Math.min(...ranges.map((range) => (range as { min: number }).min)),
      maximum: Math.max(...ranges.map((range) => (range as { max: number }).max)),
      description: 'For a pattern with a row of items: how many items the content calls for. Leave it out for other patterns.',
    },
  };
}

/** Reads the item count of an outline entry: the pattern's own number when it has no choice or none was given. */
export function readItems(pattern: PatternDef, value: unknown): { items?: number; problem?: string } {
  const range = itemRange(pattern);
  if (!range) return {};
  const items = Number(value);
  // With no number given, the number drawn, or the nearest the row takes.
  if (value === undefined || value === null || !Number.isInteger(items)) return { items: defaultItems(pattern) as number };
  if (items < range.min || items > range.max) {
    return { problem: `pattern "${pattern.id}" takes ${range.min === range.max ? `exactly ${range.min}` : `${range.min} to ${range.max}`} ${pattern.repeat?.item ?? 'item'}s, not ${items}` };
  }
  return { items };
}
