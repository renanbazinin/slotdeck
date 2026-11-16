import type { PatternDesign } from '../template/compose.ts';
import type { Direction } from '../template/manifest.ts';

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
  slides: SlideSpec[];
}

export interface SlideSpec {
  /** Id of a pattern in the template manifest. */
  pattern: string;
  /** For a pattern with a row of items: how many items this slide has. Defaults to the template's number. */
  items?: number;
  slots: Record<string, SlotValue>;
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

/** Normalises a text value to paragraphs. */
export function toParagraphs(value: string | string[]): string[] {
  const parts = Array.isArray(value) ? value : [value];
  return parts.flatMap((part) => part.split(/\r?\n/)).map((line) => line.trim());
}
