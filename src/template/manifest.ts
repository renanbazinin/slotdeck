export type Direction = 'rtl' | 'ltr';
export type AlignPolicy = 'flip' | 'keep';

/**
 * Describes how a pptx can be reused as a template: which of its slides are
 * reusable patterns, and which shapes on them are slots to fill.
 * Shapes that are not listed as slots are copied untouched (logo, rules, decorations).
 */
export interface TemplateManifest {
  version: 1;
  name: string;
  /** BCP 47 tag of the template's sample text, e.g. `he`. */
  language?: string;
  /** Reading direction the template was designed for. */
  direction: Direction;
  /**
   * What happens to left/right paragraph alignment when a deck reads the other way from the
   * template. `flip` (default) mirrors it. `keep` leaves it as designed, which suits templates
   * whose text boxes share one edge but not the other. Slots can override this.
   */
  alignOnDirectionChange?: AlignPolicy;
  patterns: PatternDef[];
}

export interface PatternDef {
  id: string;
  /** 1-based slide number in the template pptx that this pattern is cloned from. */
  sourceSlide: number;
  /** What the pattern is for. Written for the model that picks patterns. */
  description: string;
  slots: SlotDef[];
  /**
   * The pattern is built around numbers. The planner offers it only when the brief contains
   * figures, so the model is not led into inventing statistics to fill it.
   */
  figures?: boolean;
  /**
   * Rows of items (cards, steps, figures). Each row lists its items, first item first in the
   * template's reading direction, and each item lists its slot names in the same order.
   * When a deck reads the other way, values are placed mirror-wise within each row, so the
   * deck's first item still comes first for its readers. A grid is one row per line of cards.
   * Items of a row should have matching limits.
   */
  sequences?: string[][][];
  /** A row of like items (cards, steps, figures) whose number a deck may choose. */
  repeat?: RepeatDef;
}

/**
 * A row of items that can hold fewer or more than the template slide shows. Items are removed
 * or copied and the row is spaced out again; nothing is redrawn.
 */
export interface RepeatDef {
  /** What one item is called in prompts, e.g. `card`. */
  item: string;
  /** Slot name templates of one item; `{n}` stands for the item number, from 1. */
  slots: string[];
  /**
   * For each item on the template slide, first in reading order first: the shapes that move
   * together as that item (its group, and any text box that sits outside the group).
   */
  shapes: string[][];
  /** Fewest and most items the row may hold. Both default to the number on the template slide. */
  min?: number;
  max?: number;
  /** Left and right limits of the row, as fractions of the slide width. Defaults to the row as drawn. */
  span?: { left: number; right: number };
  /**
   * For a grid: how many items one row holds on the template slide. Items then fill rows in
   * reading order, spread evenly over as few rows as they need. Leave it out for a single row.
   */
  columns?: number;
}

/**
 * How many items each row of a repeat holds for a given count. A single row takes them all.
 * A grid uses as few rows as the count needs and spreads the items evenly: four items in a
 * three-column grid make two rows of two, not three and one.
 */
export function rowLayout(repeat: RepeatDef, count: number): number[] {
  if (!repeat.columns || repeat.columns >= count) return count > 0 ? [count] : [];
  const rows = Math.ceil(count / repeat.columns);
  const layout: number[] = [];
  for (let row = 0, left = count; row < rows; row++) {
    const here = Math.ceil(left / (rows - row));
    layout.push(here);
    left -= here;
  }
  return layout;
}

/**
 * For each slot of a pattern, the deck slot name its value is taken from: the slot's own name,
 * except across a row of `sequences` when the deck reads against the template.
 */
export function slotSources(pattern: PatternDef, mirrored: boolean, items?: number): Map<string, string> {
  const sources = new Map<string, string>(effectiveSlots(pattern, items).map((slot) => [slot.name, slot.name]));
  if (!mirrored) return sources;
  const rows = [...(pattern.sequences ?? [])];
  if (pattern.repeat) {
    // Each row of items is mirrored on its own, so a grid keeps its rows.
    const repeat = pattern.repeat;
    let next = 1;
    for (const size of rowLayout(repeat, itemCount(pattern, items))) {
      rows.push(
        Array.from({ length: size }, () => {
          const n = next++;
          return repeat.slots.map((name) => name.replace('{n}', String(n)));
        }),
      );
    }
  }
  for (const items of rows) {
    items.forEach((item, index) => {
      const opposite = items[items.length - 1 - index] ?? [];
      item.forEach((name, position) => {
        const source = opposite[position];
        if (source !== undefined) sources.set(name, source);
      });
    });
  }
  return sources;
}

/**
 * What to do when a deck gives no value for a slot:
 * - `error`: the slot is required (default)
 * - `remove`: optional content; the shape is deleted, which needs `layout: 'adaptive'`
 * - `keep`: leave the template's own content in place (decoration that suits any content)
 * - `drop`: delete the shape in any layout (decoration drawn for the template's sample
 *   content, such as a sticker placed on a particular photo)
 *
 * Slots with `keep` or `drop` are never offered to the planning model.
 */
export type EmptyPolicy = 'error' | 'remove' | 'keep' | 'drop';

interface SlotBase {
  /** Name the deck uses to address the slot; unique within the pattern. */
  name: string;
  /** `p:cNvPr/@id` of the shape on the source slide. */
  shapeId: string;
  description?: string;
  empty?: EmptyPolicy;
}

export interface TextSlotDef extends SlotBase {
  kind: 'text';
  /** Soft limit on total characters; exceeding it produces a warning. */
  maxChars?: number;
  /** Soft limit on the number of paragraphs. */
  maxParagraphs?: number;
  /** Take the value from `deck.fields[deckField]` when the slide does not set it. */
  deckField?: string;
  /** Overrides the manifest's `alignOnDirectionChange` for this slot. */
  alignOnDirectionChange?: AlignPolicy;
  /**
   * Which run of the template paragraph supplies the formatting, for paragraphs that mix
   * styles (e.g. a coloured label followed by plain text). Defaults to `first`.
   */
  styleRun?: 'first' | 'last';
  /** Font size of the slot's text on the template slide, in points. Decides whether a deck's text size applies. */
  size?: number;
  /** Sets the font size outright, in points: for a pattern derived from a slide that uses the shape differently. */
  fontSize?: number;
  /** Moves and resizes the text box: for a derived pattern. Only for a shape that is not inside a group. */
  box?: SlideBox;
  /** Space after each paragraph, in points. */
  paragraphGap?: number;
  /** Whether a deck's text size applies. Defaults to body text: 12 pt and under, and not a deck field. */
  scales?: boolean;
}

export type ImageKind = 'photo' | 'icon' | 'graphic';

/** A rectangle as fractions of the slide (0..1). */
export interface SlideBox {
  x: number;
  y: number;
  w: number;
  h: number;
}

export interface ImageSlotDef extends SlotBase {
  kind: 'image';
  imageKind: ImageKind;
  /** `cover` crops the image to fill the frame; `contain` shrinks the frame to the image. Defaults by kind. */
  fit?: 'cover' | 'contain';
  /**
   * The area the picture may use, when that is larger or other than the sample picture's own
   * frame (a chart drawn small in the middle of a slide that has room for a wide diagram).
   * Only for pictures that are not inside a group.
   */
  box?: SlideBox;
  /** Set on slots of copied items: the template shape the copy was made from. */
  copyOf?: string;
}

/** Formatting for table text, for templates whose sample table leaves it undefined. */
export interface TableTextStyle {
  /** Font size in points. */
  size?: number;
  bold?: boolean;
  /** Hex colour without `#`, e.g. `FFFFFF`. */
  color?: string;
  /** Typeface name. */
  font?: string;
}

export interface TableSlotDef extends SlotBase {
  kind: 'table';
  /** Soft limit on columns. Defaults to the sample table's. */
  maxColumns?: number;
  /** Soft limit on body rows (the header row not counted). Defaults to the sample table's. */
  maxRows?: number;
  /** Soft limit on the characters in one cell. */
  maxCellChars?: number;
  headerText?: TableTextStyle;
  bodyText?: TableTextStyle;
}

/**
 * A native, editable chart. The slot points at the sample picture of a chart; the picture is
 * replaced by a real chart with its data in an embedded workbook.
 */
export interface ChartSlotDef extends SlotBase {
  kind: 'chart';
  /** The area the chart may use, when other than the sample picture's frame. */
  box?: SlideBox;
  /** Soft limit on the number of categories (points along the axis, or pie slices). */
  maxCategories?: number;
  /** Soft limit on the number of series. */
  maxSeries?: number;
  /** Series colours in order, as hex without `#`. They repeat when there are more series. */
  colors?: string[];
  /** Formatting of axis labels, data labels and the legend. */
  text?: TableTextStyle;
  /** Colour of the gridlines and axis lines, as hex without `#`. */
  lineColor?: string;
}

export type SlotDef = TextSlotDef | ImageSlotDef | TableSlotDef | ChartSlotDef;

export function defaultFit(slot: ImageSlotDef): 'cover' | 'contain' {
  return slot.fit ?? (slot.imageKind === 'photo' ? 'cover' : 'contain');
}

/** Copied shapes get the id of their original plus a multiple of this, so the id is known without the file. */
export const COPY_ID_STEP = 10000;

/** How many items a slide's row holds: what the deck asks for, or the number on the template slide. */
export function itemCount(pattern: PatternDef, items?: number): number {
  if (!pattern.repeat) return 0;
  return items ?? pattern.repeat.shapes.length;
}

/** The range of item counts a pattern accepts, or `undefined` when it has no row to vary. */
export function itemRange(pattern: PatternDef): { min: number; max: number; drawn: number } | undefined {
  if (!pattern.repeat) return undefined;
  const drawn = pattern.repeat.shapes.length;
  return { min: pattern.repeat.min ?? drawn, max: pattern.repeat.max ?? drawn, drawn };
}

/**
 * The slots of a pattern for a given number of items. Slots of items beyond the count are left
 * out; items beyond those drawn on the template get copies of the last item's slots, pointing
 * at the shapes the renderer will copy.
 */
export function effectiveSlots(pattern: PatternDef, items?: number): SlotDef[] {
  const repeat = pattern.repeat;
  if (!repeat) return pattern.slots;
  const drawn = repeat.shapes.length;
  const count = itemCount(pattern, items);
  const itemOf = new Map<string, number>();
  for (let n = 1; n <= drawn; n++) for (const name of repeat.slots) itemOf.set(name.replace('{n}', String(n)), n);

  const slots = pattern.slots.filter((slot) => (itemOf.get(slot.name) ?? 0) <= count);
  for (let n = drawn + 1; n <= count; n++) {
    for (const name of repeat.slots) {
      const original = pattern.slots.find((slot) => slot.name === name.replace('{n}', String(drawn)));
      if (!original) continue;
      const copy = { ...original, name: name.replace('{n}', String(n)), shapeId: String(Number(original.shapeId) + (n - drawn) * COPY_ID_STEP) };
      slots.push(copy.kind === 'image' ? { ...copy, copyOf: original.shapeId } : copy);
    }
  }
  return slots;
}

/** Deck text sizes by name, as factors on the template's body text. */
export const TEXT_SIZES = { normal: 1, large: 1.25, 'x-large': 1.5 } as const;
export type TextSize = keyof typeof TEXT_SIZES;

export function scalesWithText(slot: TextSlotDef): boolean {
  if (slot.scales !== undefined) return slot.scales;
  const size = slot.fontSize ?? slot.size;
  return !slot.deckField && size !== undefined && size <= 12;
}

/**
 * A text slot's limits at a given text scale. Larger text fits fewer characters per line and
 * fewer lines in the same box, so the character budget falls with the square of the scale.
 */
export function textBudget(slot: TextSlotDef, scale = 1): { maxChars?: number; maxParagraphs?: number } {
  const factor = scalesWithText(slot) ? scale : 1;
  const budget: { maxChars?: number; maxParagraphs?: number } = {};
  if (slot.maxChars !== undefined) budget.maxChars = Math.max(1, Math.floor(slot.maxChars / (factor * factor)));
  if (slot.maxParagraphs !== undefined) budget.maxParagraphs = Math.max(1, Math.floor(slot.maxParagraphs / factor));
  return budget;
}
