import type { PatternDesign } from '../template/compose.ts';
import type { Direction } from '../template/manifest.ts';
import { writable } from '../util/chars.ts';

/**
 * A deck as plain JSON: the source of truth that is rendered, stored and edited.
 * It holds no binary data; images are referenced by asset id.
 */
export interface DeckSpec {
  version: 1;
  title?: string;
  /** The request the deck was first written from. Later edits use it as background. */
  brief?: string;
  /** BCP 47 tag of the deck's text. */
  language?: string;
  /** Reading direction of the deck. Defaults to the template's. */
  direction?: Direction;
  /** Values shared by every slide, e.g. the header text. */
  fields?: Record<string, string>;
  /** Size of body text as a factor on the template's: 1 as designed, 1.25 large, 1.5 extra large. */
  textScale?: number;
  /**
   * Layouts made for this deck because the template had none that suited some of its slides (see
   * `composeLayouts`). They are built onto the template whenever the deck is rendered; the
   * template file is not changed. Slides use them by their ids like any other pattern.
   */
  layouts?: PatternDesign[];
  /**
   * Colours in place of the template's, for every slide: "make it blue". A slide may state its
   * own. Left out, a deck has exactly the template's colours.
   */
  colors?: DeckColors;
  /**
   * The only layouts the deck was asked to use, by id (`free` among them when free slides were
   * asked for). It is a wish the planner keeps to when the deck is changed later; rendering and
   * validation do not look at it.
   */
  patterns?: string[];
  slides: SlideSpec[];
  /**
   * The language the document the deck was made from is written in, as a BCP 47 tag, where the
   * planner said (`DeckPlan.documentLanguage`). A later change sets the document into its slides
   * as it is only where that is the deck's language.
   */
  documentLanguage?: string;
  /**
   * The document the deck was made from (`Brief.document`), word for word. A later change draws
   * on it: a slide that is written again says its own passages of it. Rendering and validation do
   * not read it.
   */
  document?: string;
}

/** The id a slide carries in `pattern` when the software lays it out itself (see {@link SlideSpec.blocks}). */
export const FREE_PATTERN = 'free';

/**
 * What a free slide can be made of. `text` is running text, `list` a list of points, `card` a
 * heading and a text on a panel, `figure` a big number with what it counts, `step` a numbered
 * card, `quote` a quotation, `note` a highlighted line across the slide, `chart` a native chart,
 * `table` a table, `photo` a picture that fills its frame, `graphic` a picture shown whole.
 */
export const FREE_BLOCK_KINDS = ['text', 'list', 'card', 'figure', 'step', 'quote', 'note', 'chart', 'table', 'photo', 'graphic'] as const;
export type FreeBlockKind = (typeof FREE_BLOCK_KINDS)[number];

/** How the blocks of a free slide are set out, where the deck says so. Left out, the software decides. */
export interface FreeArrange {
  /** How many blocks stand side by side on one line. */
  columns?: number;
  /**
   * Where the first block stands when it is the main one and larger than the rest: at the
   * `left` or the `right` of them, or on `top`. `none` sets all blocks alike. Left out, a single
   * chart, table or picture among a few other blocks leads by itself.
   */
  lead?: 'left' | 'right' | 'top' | 'none';
  /**
   * The slide carries a text a person asked to have whole: an amount of text they asked for, or
   * their own text. A free slide of one text, with or without a heading. Where that text is too
   * long for four fifths of the template's text size it is set smaller, down to the least size that
   * still reads (9 pt on a slide 13.33 inches wide, less on a narrower one), in two or three
   * columns, in the most room the frame has. A text that fits is set as large as it fits, as without
   * the flag. One rule says whether such a text fits and what it holds, when it is planned, written
   * and checked: its paragraphs counted by length, each in whole lines, its columns filled line by
   * line as PowerPoint fills them, each keeping a line free. Set by the planner; left out, the slide
   * is set out as before.
   */
  dense?: boolean;
}

/**
 * Colours a deck asks for in place of the template's. Each is six hex digits without `#`. The
 * template's other colours, its greys, its logo and its photos stay as they are.
 */
export interface DeckColors {
  /** Takes the place of the template's main colour. Its tints and shades follow. */
  accent?: string;
  /** The same for the template's second colour. */
  accent2?: string;
  /** Exact replacements of single colours of the template, from one hex to another. */
  replace?: Record<string, string>;
}

export interface SlideSpec {
  /** Id of a pattern in the template manifest. */
  pattern: string;
  /** For a pattern with a row of items: how many items this slide has. Defaults to the template's number. */
  items?: number;
  slots: Record<string, SlotValue>;
  /**
   * Makes the slide a free slide: the software lays it out itself, on the frame of one of the
   * template's layouts, from these blocks in reading order. Any number of blocks, of any kinds.
   * `pattern` is then `free`. The content of block n is in `slots` under the names `b1Text`,
   * `b2Chart`, ... (see `freeSlide`, which writes them), and the slide's title under `title`.
   */
  blocks?: FreeBlockKind[];
  /** For a free slide: how its blocks are set out. */
  arrange?: FreeArrange;
  /**
   * For a free slide: the layout whose frame it stands on (background, header, logo, title).
   * Recorded when the slide is planned, so that it stands on the same frame later. Left out,
   * the software picks the layout whose title sits where most titles do.
   */
  base?: string;
  /** Colours for this slide alone, in place of the deck's. */
  colors?: DeckColors;
  /**
   * In a deck made from a document: the passages of it this slide carries, by number from 1, as
   * `documentPassages` numbers them. Rendering does not read it.
   */
  passages?: number[];
}

/** Whether a slide is one the software lays out itself: it lists its `blocks`. */
export function isFreeSlide(slide: Pick<SlideSpec, 'blocks'>): boolean {
  return Array.isArray(slide.blocks);
}

/** Reference to an image passed to the renderer alongside the deck. */
export interface ImageRef {
  asset: string;
  /** Alternative text for screen readers. */
  alt?: string;
  /**
   * Description the image can be generated from when the asset is not supplied.
   * Kept in the deck so a later re-render can recreate the picture.
   */
  prompt?: string;
}

/** Content of a table slot. */
export interface TableValue {
  /** Column headings. Their number sets the number of columns. */
  header: string[];
  /** Body rows, each with one cell per column. */
  rows: string[][];
}

export type ChartType = 'column' | 'bar' | 'line' | 'pie';

/** Content of a chart slot. */
export interface ChartValue {
  /** `column` has vertical bars, `bar` horizontal ones. */
  type: ChartType;
  /** Labels along the axis, or the names of the pie slices. */
  categories: string[];
  /** One or more series, each with one value per category. A pie takes exactly one. */
  series: Array<{ name: string; values: number[] }>;
  /** `percent` shows values with a percent sign (45 is shown as 45%). */
  format?: 'number' | 'percent';
}

/**
 * - text slot: a string (newlines split paragraphs) or one string per paragraph
 * - image slot: an {@link ImageRef}
 * - table slot: a {@link TableValue}
 * - chart slot: a {@link ChartValue}
 * - `null`: explicitly leave the slot out (the shape is removed)
 */
export type SlotValue = string | string[] | ImageRef | TableValue | ChartValue | null;

export function isImageRef(value: unknown): value is ImageRef {
  return typeof value === 'object' && value !== null && !Array.isArray(value) && typeof (value as ImageRef).asset === 'string';
}

export function isTableValue(value: unknown): value is TableValue {
  if (typeof value !== 'object' || value === null || Array.isArray(value)) return false;
  const table = value as TableValue;
  return Array.isArray(table.header) && Array.isArray(table.rows) && table.rows.every((row) => Array.isArray(row));
}

export function isChartValue(value: unknown): value is ChartValue {
  if (typeof value !== 'object' || value === null || Array.isArray(value)) return false;
  const chart = value as ChartValue;
  return Array.isArray(chart.categories) && Array.isArray(chart.series);
}

/**
 * Whether a value is what a text slot takes: a string, or a list of strings. A working part of
 * the checks and the renderer: the package does not export it.
 */
export function isTextValue(value: unknown): value is string | string[] {
  return typeof value === 'string' || (Array.isArray(value) && value.every((part) => typeof part === 'string'));
}

/**
 * One line of text as it can be written into a file: a cell of a table, a label of a chart, the
 * alternative text of a picture. A vertical tab, the line break of text copied out of a text
 * box, is a space; the other characters no file can hold are left out. A working part of the
 * checks and the renderer: the package does not export it.
 */
export function plainText(value: unknown): string {
  return writable(String(value ?? '').replace(/\u000B/g, ' '));
}

/**
 * Normalises a text value to paragraphs. A vertical tab breaks a line as a newline does, and
 * characters that no file can hold are left out. A part that is no string (a number, an object)
 * counts as empty here and never throws: validation is what says "expected text" of it.
 */
export function toParagraphs(value: string | string[]): string[] {
  const parts: unknown[] = Array.isArray(value) ? value : [value];
  return parts.flatMap((part) => (typeof part === 'string' ? part : '').split(/\r?\n|\u000B/)).map((line) => writable(line).trim());
}
