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
  /**
   * The direction the charts of a deck that reads right to left read in. `rtl`: the first category
   * of columns and lines stands on the right, with their value axis, horizontal bars grow from the
   * right, and the chart's data sheet is right to left. Left out, or `ltr`: charts read left to
   * right, as many charts of time do in Hebrew too. A deck that reads left to right has charts
   * that read left to right; a deck may state its own (`DeckSpec.chartDirection`). From 0.7.0.
   */
  chartDirection?: Direction;
  /**
   * Enlarges the body text of every deck made from this template, as a factor on the designed
   * size: 1.3 turns 9 pt into about 12 pt. For a template whose body text is too small to read
   * comfortably. Budgets shrink with it, and a deck's own text size is applied on top.
   */
  textScale?: number;
  /**
   * How wide PowerPoint sets a character of the template's running text, as a share of its size,
   * in the template's own face: the width that 19 full lines of it in 20 stay under, such as 0.44
   * for a narrow face. Stated, the free slides of a deck count their running text (its texts and
   * lists) with it, in place of an estimate that leaves room for any face (0.55, and 0.48 for a
   * long text set small): such text is set larger, up to the template's size.
   * Between 0.3 and 0.8; left out, the estimate. A value under the face's own lets text run past
   * its box, so it is measured, not guessed (see docs/text-fit.md).
   */
  characterWidth?: number;
  /**
   * The least size, in points, that a long text a person asked to have whole (set dense on a free
   * slide) is set at: it is set no smaller, and is cut, or given more slides, where it needs
   * smaller. A value over the template's own size of running text is taken as that size, and one
   * between half points as the next half point up. Left out,
   * 9 pt on a slide 13.33 inches wide and less on a narrower one, by its width (7 pt on one 10
   * inches wide). A lower value also lowers the size under which a slide is told its text is
   * small. Between 5 and 24.
   */
  leastTextSize?: number;
  /**
   * Lets text that does not fit its box be set smaller: down to `to` of its size (0.85 is 85%)
   * and never under `minSize` points. It needs fonts to measure with, so it only acts for a
   * template loaded with fonts; without them nothing is shrunk.
   */
  shrinkText?: { to: number; minSize?: number };
  /**
   * The template's main colour and its second colour, as six hex digits without `#`: what a
   * deck's `colors.accent` and `colors.accent2` take the place of. Left out, the software finds
   * them by how much each colour is used (see `templatePalette`). State them for a template where
   * that picks the wrong one. `false` keeps every deck in the template's own colours: a deck's
   * `colors` are then left out, with a warning.
   */
  colors?: false | { accent?: string; accent2?: string };
  /**
   * Free slides: slides the software lays out itself on the frame of one of the template's
   * layouts. `false` stops them being offered when a deck is planned (a stored deck that has
   * one still renders). An object names the layout whose frame they stand on, and which of the
   * template's shapes each part is a copy of, where the software's own choice does not suit.
   */
  free?: false | { base?: string; styles?: Record<string, { slide: number; shape: string }> };
  patterns: PatternDef[];
}

const halfPoints = (size: number, scale: number): number => Math.round(size * scale * 2) / 2;

/**
 * A table slot as it is at a body text scale: the text sizes the manifest states are enlarged,
 * and a cell holds less. A table whose sizes come from the sample table alone is left as it is,
 * since there is nothing stated to enlarge. So is the table of a slide the software laid out,
 * which has a box of its own: its sizes and what it holds were worked out at the deck's scale.
 */
export function tableAtScale(slot: TableSlotDef, scale: number): TableSlotDef {
  if (scale === 1 || slot.bodyText?.size === undefined || (slot as { box?: SlideBox }).box !== undefined) return slot;
  const scaled: TableSlotDef = { ...slot, bodyText: { ...slot.bodyText, size: halfPoints(slot.bodyText.size, scale) } };
  if (slot.headerText?.size !== undefined) scaled.headerText = { ...slot.headerText, size: halfPoints(slot.headerText.size, scale) };
  if (slot.maxCellChars !== undefined) scaled.maxCellChars = Math.max(1, Math.floor(slot.maxCellChars / (scale * scale)));
  return scaled;
}

/** A chart slot as it is at a body text scale: the label size the manifest states is enlarged. */
export function chartAtScale(slot: ChartSlotDef, scale: number): ChartSlotDef {
  if (scale === 1 || slot.text?.size === undefined) return slot;
  return { ...slot, text: { ...slot.text, size: halfPoints(slot.text.size, scale) } };
}

/**
 * What is wrong with a manifest's own `textScale`, as a sentence; `undefined` when it is a factor
 * a deck can be made with. Every scaled size of every deck is multiplied by it, and one that is
 * no number is written into the file as it is (`sz="NaN"`), which PowerPoint does not open. Left
 * out or null it is 1; a number written as text is read as that number, as it always was.
 */
export function templateScaleProblem(value: unknown): string | undefined {
  if (value === undefined || value === null) return undefined;
  const scale = typeof value === 'number' ? value : typeof value === 'string' && value.trim() !== '' ? Number(value) : NaN;
  if (Number.isFinite(scale) && scale > 0) return undefined;
  const given = typeof value === 'number' ? String(value) : JSON.stringify(value);
  return `the template's textScale is ${given}; it is a factor on body text such as 1.25, which changeManifest sets with "textScale"`;
}

/** The range of a manifest's `characterWidth`, as a share of the size. */
export const CHARACTER_WIDTH = { least: 0.3, most: 0.8 } as const;
/** The range of a manifest's `leastTextSize`, in points. */
export const LEAST_TEXT_SIZE = { least: 5, most: 24 } as const;

/**
 * What is wrong with a manifest's own `characterWidth` or `leastTextSize`, as a sentence;
 * `undefined` when it is left out (or null) or a number in its range. A setting that is wrong is
 * not used: the software counts as it does without it.
 */
export function templateTextProblem(name: 'characterWidth' | 'leastTextSize', value: unknown): string | undefined {
  if (value === undefined || value === null) return undefined;
  const range = name === 'characterWidth' ? CHARACTER_WIDTH : LEAST_TEXT_SIZE;
  if (typeof value === 'number' && Number.isFinite(value) && value >= range.least && value <= range.most) return undefined;
  const given = typeof value === 'number' ? String(value) : JSON.stringify(value);
  return name === 'characterWidth'
    ? `the template's characterWidth is ${given}; it is how wide a character of its text is set, as a share of its size, between ${range.least} and ${range.most}, such as 0.43; it is not used`
    : `the template's leastTextSize is ${given}; it is a size in points between ${range.least} and ${range.most}, such as 10; it is not used`;
}

/**
 * A manifest's `characterWidth` or `leastTextSize` where it is one the software uses (see
 * `templateTextProblem`). A `leastTextSize` between half points is taken up to the next: text is
 * set in half points, and 12.2 set at 12 would be under it.
 */
export function templateTextSetting(manifest: Pick<TemplateManifest, 'characterWidth' | 'leastTextSize'>, name: 'characterWidth' | 'leastTextSize'): number | undefined {
  const value: unknown = manifest[name];
  if (value === undefined || value === null || templateTextProblem(name, value) !== undefined) return undefined;
  return name === 'leastTextSize' ? Math.ceil((value as number) * 2 - 1e-9) / 2 : (value as number);
}

/**
 * Whether the charts of a deck read right to left: as the deck states (`chartDirection`), or else
 * as the template states for a deck that reads right to left. Anything but `rtl` reads left to right.
 */
export function chartsReadRightToLeft(manifest: Pick<TemplateManifest, 'direction' | 'chartDirection'>, deck: { direction?: Direction | undefined; chartDirection?: Direction | undefined }): boolean {
  const stated = deck.chartDirection ?? ((deck.direction ?? manifest.direction) === 'rtl' ? manifest.chartDirection : undefined);
  return stated === 'rtl';
}

/** The factor on body text for a deck: the template's own setting times the deck's. A deck's factor that is null, as an application stores a number that is not set, is 1. */
export function bodyTextScale(manifest: Pick<TemplateManifest, 'textScale'>, deckScale: number | null = 1): number {
  return (manifest.textScale ?? 1) * (deckScale ?? 1);
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
  /**
   * Lets items grow wider when a line holds fewer of them than the template draws, so that they
   * fill the row's span and do not float in the middle of it. The value is the most an item may
   * grow, as a factor on its drawn width: 1.5 lets a card become half as wide again. Backgrounds
   * and text boxes widen; pictures keep their size and their side. Left out, items keep their size.
   */
  widen?: number;
}

/**
 * The factor by which the items of a row are at least wider than drawn for a given count: 1 when
 * the row does not widen. Items grow when the fullest line holds fewer than the template draws on
 * a line; with the same gaps between them, they are then at least `drawn / held` times as wide.
 */
export function widenFactor(repeat: RepeatDef, count: number): number {
  if (!repeat.widen || !(repeat.widen > 1)) return 1;
  const drawn = repeat.columns ?? repeat.shapes.length;
  const held = Math.max(0, ...rowLayout(repeat, count));
  return held > 0 && held < drawn ? Math.min(repeat.widen, drawn / held) : 1;
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
  /**
   * How much a chart shows besides its data. `full` (the default): numbers on the bars,
   * gridlines, axes and a legend. `compact`: no numbers and no gridlines, for a chart a few
   * inches wide. `spark`: no axes and no legend either, for a chart among many on a slide.
   */
  detail?: 'full' | 'compact' | 'spark';
  /**
   * How the chart is drawn in a place whose size is known. The software sets it for the charts
   * of a slide it lays out itself; a manifest does not state it, and a chart slot without it is
   * drawn as it always was.
   */
  drawn?: ChartDrawn;
}

/** What is worked out for a chart whose place on the slide is known. Lengths in points. */
export interface ChartDrawn {
  /** The frame the chart is drawn in. */
  width: number;
  height: number;
  /** The colour the chart stands on: colours past the end of the slot's are tints towards it. */
  ground?: string;
  /**
   * Where the value axis ends, and the step between its lines when the axis is shown. It starts
   * at `min`, or at nothing when that is left out. Like charts of one slide are given the same
   * axis, so that a bar of one can be read against a bar of the next. Left out, the chart keeps
   * the scale its own values give it.
   */
  axis?: { max: number; step?: number; min?: number };
  /** The numbers are written at the bars. Left out: as the slot's `detail` says. */
  numbers?: boolean;
  /** `false`: no value axis is drawn, because the numbers at the bars say it all. */
  valueAxis?: boolean;
  /** `false`: the chart draws no legend, because its slide shows one for all its charts. */
  legend?: boolean;
  /** How its numbers are written, where that is decided for several charts together. Left out, by the chart's own values. */
  look?: { divisor: 1 | 1000 | 1000000; suffix: '' | 'K' | 'M' | '%'; decimals: number };
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

/**
 * The number of items to plan a row with when none is given: the number the template draws, or
 * the nearest the row takes. `undefined` for a pattern without a row.
 */
export function defaultItems(pattern: PatternDef): number | undefined {
  const range = itemRange(pattern);
  return range ? Math.min(Math.max(range.drawn, range.min), range.max) : undefined;
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
  // Wider items hold more text: the budgets of the row's texts grow with the items.
  const factor = widenFactor(repeat, count);
  if (factor === 1) return slots;
  return slots.map((slot) => (slot.kind === 'text' && slot.maxChars !== undefined && itemOf.has(slot.name) ? { ...slot, maxChars: Math.floor(slot.maxChars * factor) } : slot));
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
