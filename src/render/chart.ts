// Builds a native PowerPoint chart: the chart part, and the small Excel workbook that holds its
// data so the chart can be edited with "Edit Data".
import { strToU8, zipSync } from 'fflate';
import { mix } from '../color/space.ts';
import { knownValues } from '../deck/spec.ts';
import type { ChartValue } from '../deck/spec.ts';
import { XML_HEADER, escapeXmlAttr } from '../pptx/xml.ts';
import type { ChartSlotDef, Direction } from '../template/manifest.ts';
import { writable } from '../util/chars.ts';
import { leftToRightTag } from '../util/language.ts';
import { marksRightToLeft, rightToLeftTag } from './text.ts';

const NS_C = 'http://schemas.openxmlformats.org/drawingml/2006/chart';
const NS_A = 'http://schemas.openxmlformats.org/drawingml/2006/main';
const NS_R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';

export const CHART_CONTENT_TYPE = 'application/vnd.openxmlformats-officedocument.drawingml.chart+xml';
export const CHART_REL_TYPE = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart';
export const WORKBOOK_REL_TYPE = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/package';
export const WORKBOOK_CONTENT_TYPE = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';
export const CHART_GRAPHIC_URI = NS_C;

/** The earliest date a zip entry can carry. */
const ZIP_EPOCH = new Date(1980, 0, 1);

const DEFAULT_COLORS = ['4472C4', 'ED7D31', 'A5A5A5', 'FFC000', '5B9BD5', '70AD47'];
const SHEET = 'Sheet1';

/** Text between two tags, in the chart and in its workbook alike. */
const text = (value: string) => writable(value).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
const hex = (value: string) => value.replace(/^#/, '').toUpperCase();

/** Spreadsheet column name for a zero-based index: 0 is A, 26 is AA. */
function columnName(index: number): string {
  let name = '';
  for (let n = index + 1; n > 0; n = Math.floor((n - 1) / 26)) name = String.fromCharCode(65 + ((n - 1) % 26)) + name;
  return name;
}

const fill = (color: string) => `<a:solidFill><a:srgbClr val="${hex(color)}"/></a:solidFill>`;
const line = (color: string, width = 9525) => `<a:ln w="${width}" cap="flat">${fill(color)}<a:round/></a:ln>`;
const noLine = '<a:ln><a:noFill/></a:ln>';

/** The workbook cell range of a series' values, of its name, and of the categories. */
const ranges = (chart: ChartValue) => {
  const last = chart.categories.length + 1;
  return {
    categories: `${SHEET}!$A$2:$A$${last}`,
    name: (series: number) => `${SHEET}!$${columnName(series + 1)}$1`,
    values: (series: number) => `${SHEET}!$${columnName(series + 1)}$2:$${columnName(series + 1)}$${last}`,
  };
};

function stringRef(formula: string, values: string[]): string {
  const points = values.map((value, index) => `<c:pt idx="${index}"><c:v>${text(value)}</c:v></c:pt>`).join('');
  return `<c:strRef><c:f>${text(formula)}</c:f><c:strCache><c:ptCount val="${values.length}"/>${points}</c:strCache></c:strRef>`;
}

/** The cached numbers of a series. A value that is not known has no point: the chart leaves a gap there, where a 0 would tell the reader it was nothing. */
function numberRef(formula: string, values: ReadonlyArray<number | null>, format: string): string {
  const points = values.map((value, index) => (value === null ? '' : `<c:pt idx="${index}"><c:v>${value}</c:v></c:pt>`)).join('');
  return `<c:numRef><c:f>${text(formula)}</c:f><c:numCache><c:formatCode>${text(format)}</c:formatCode><c:ptCount val="${values.length}"/>${points}</c:numCache></c:numRef>`;
}

function dataLabels(format: string, show: 'value' | 'none', position?: string): string {
  if (show === 'none') return '<c:dLbls><c:delete val="1"/></c:dLbls>';
  return (
    `<c:dLbls><c:numFmt formatCode="${escapeXmlAttr(format)}" sourceLinked="0"/><c:spPr><a:noFill/>${noLine}</c:spPr>` +
    (position ? `<c:dLblPos val="${position}"/>` : '') +
    `<c:showLegendKey val="0"/><c:showVal val="1"/><c:showCatName val="0"/><c:showSerName val="0"/><c:showPercent val="0"/><c:showBubbleSize val="0"/></c:dLbls>`
  );
}

// ---- How the numbers of a chart are written, and where its axis ends -------------------------
// For a chart whose place is known (`slot.drawn`: the charts of a free slide). A chart of a
// template's own layout is written as it always was.

/** How the numbers of a chart are shown: divided by so much, with a letter for the unit, to so many decimals. */
export interface NumberLook {
  divisor: 1 | 1000 | 1000000;
  suffix: '' | 'K' | 'M' | '%';
  decimals: number;
}

const exactAt = (value: number, decimals: number): boolean => Math.abs(value * 10 ** decimals - Math.round(value * 10 ** decimals)) < 1e-6 * Math.max(1, Math.abs(value));
/** The fewest decimals, none to two, at which every value is shown as it is; two when none does. */
const decimalsOf = (values: readonly number[]): number => [0, 1, 2].find((decimals) => values.every((value) => exactAt(value, decimals))) ?? 2;

/**
 * One way of writing numbers for some values: for a chart, or for the like charts of a slide.
 * Thousands are set apart. Every value keeps the decimals it has, up to two, so that a chart
 * shows the numbers it was given and not rounded ones. Where the largest is a million or more
 * every value is written in millions: with the decimals the values have there, up to two (1.25M,
 * 2M), and where that is not all of them, with two, or three where two would write two different
 * values alike (1,385,000 is 1.39M, as a reader says it). Thousands are written so only when that
 * loses nothing: the largest is ten thousand or more and every value is a whole number of thousands.
 */
export function numberLook(values: readonly number[], percent: boolean): NumberLook {
  if (percent) return { divisor: 1, suffix: '%', decimals: decimalsOf(values) };
  const largest = values.reduce((most, value) => Math.max(most, Math.abs(value)), 0);
  const millions = values.map((value) => value / 1000000);
  if (largest >= 1000000) {
    const exact = [0, 1, 2].find((decimals) => millions.every((value) => exactAt(value, decimals)));
    const apart = (decimals: number) => new Set(millions.map((value) => value.toFixed(decimals))).size === new Set(millions).size;
    return { divisor: 1000000, suffix: 'M', decimals: exact ?? (apart(2) ? 2 : 3) };
  }
  if (largest >= 10000 && values.every((value) => exactAt(value / 1000, 0))) return { divisor: 1000, suffix: 'K', decimals: 0 };
  return { divisor: 1, suffix: '', decimals: decimalsOf(values) };
}

/** The format code of a look, as a spreadsheet writes it. `decimals` is for an axis, whose steps can need other decimals than the values; `zero` writes a bare 0 for nothing. */
export function formatCode(look: NumberLook, decimals = look.decimals, zero = false): string {
  const digits = decimals > 0 ? `.${'0'.repeat(decimals)}` : '';
  if (look.suffix === '%') return `0${digits}"%"`;
  if (look.suffix === '') return `#,##0${digits}`;
  const code = `#,##0${digits}${look.divisor === 1000000 ? ',,' : ','}"${look.suffix}"`;
  return zero ? `${code};-${code};0` : code;
}

/** How wide a name of a category is counted under columns, in sizes of its text a character: about what PowerPoint sets ordinary text at. */
export const NAME_WIDE = 0.5;

/** A number as a chart with this look shows it: for counting how wide it is. */
export function shownAs(value: number, look: NumberLook, decimals = look.decimals): string {
  const scaled = Math.abs(value / look.divisor).toFixed(decimals);
  const [whole, part] = scaled.split('.') as [string, string | undefined];
  return `${value < 0 ? '-' : ''}${whole.replace(/\B(?=(\d{3})+(?!\d))/g, ',')}${part ? `.${part}` : ''}${look.suffix}`;
}

/**
 * A round end for a value axis that starts at nothing, and the step between its lines. A step
 * is 1, 2, 2.5 or 5 times a power of ten, and the axis reaches `most` in no more than `steps`
 * of them. Of the steps that do, the one whose axis ends nearest above `most` is taken, so that
 * the longest bar uses as much of the plot as a round axis allows; of two that end alike, the
 * smaller. `grain` is the smallest amount the numbers of the chart are written to (1 for whole
 * numbers, 0.1 for one decimal): a step is a whole number of grains, so that an axis of whole
 * numbers is not labelled 2.5, 7.5.
 */
export function roundAxis(most: number, steps: number, grain = 0): { max: number; step: number } {
  if (!(most > 0) || !Number.isFinite(most)) return { max: 1, step: 1 };
  const within = Math.max(1, Math.floor(steps));
  const tidy = (value: number) => Number(value.toPrecision(12));
  let best: { max: number; step: number } | undefined;
  let first: number | undefined;
  for (let power = Math.floor(Math.log10(most / within)) - 1; power < 400; power++) {
    // Every step from ten times the smallest that serves on ends further off than that one does.
    if (first !== undefined && power > first + 1) break;
    for (const base of [1, 2, 2.5, 5]) {
      const step = tidy(power >= 0 ? base * 10 ** power : base / 10 ** -power);
      if (grain > 0 && Math.abs(step / grain - Math.round(step / grain)) > 1e-6) continue;
      const count = Math.max(1, Math.ceil(most / step - 1e-9));
      if (count > within) continue;
      first ??= power;
      const max = tidy(step * count);
      if (!best || max < best.max - 1e-9) best = { max, step };
    }
  }
  return best ?? { max: most, step: most };
}

/** The most labels a value axis has: more are a ruler, and the eye does not read them. */
const MOST_LABELS = 5;

/**
 * A value axis for values from `low` to `high`: where it starts, where it ends, and the step
 * between its labels.
 *
 * Its labels stand at round values, three to five of them, and no more than `steps` steps where
 * the plot is short: a step is 1, 2 or 5 times a power of ten, or 2.5 times one where its labels
 * stay whole in the unit the numbers are written in (`unit`: 1, a thousand or a million; 25 and
 * 250 are whole, 2.5 is not). A step is also a whole number of `grain`, the smallest amount the
 * numbers are written to.
 *
 * The axis starts at nothing; at the whole step under the lowest value where a value is under
 * nothing; or, for a line that stays far above nothing, at the whole step at or under `from`.
 * Its labels start where it starts. It ends at the half step at or over the highest value, which
 * may be over its last label: so the longest bar uses as much of its plot as a round axis allows,
 * and a value just over a label does not cost the chart a whole step of empty plot. Where that
 * half step would take one label too many, it ends at the tenth of a step over the highest
 * value, and where the axis would then have fewer than three labels, at its third.
 *
 * Where values are under nothing and numbers stand at the ends of their bars, `under` is the share
 * of the axis that the number of the lowest bar takes: the axis starts a step lower where the
 * room under the lowest value is less, so that the number stands clear of the names.
 *
 * Of the steps that serve, the one whose axis the values fill best is taken. A step that writes a
 * decimal on every label is taken only where it fills the axis clearly better than a whole one
 * (0, 2, 4, 6 and not 0.0, 2.5, 5.0); of two that fill it alike, the one with fewer labels, and the
 * one whose last label is its end.
 */
export function roundSpan(low: number, high: number, options: { steps: number; grain?: number; unit?: number; from?: number | undefined; under?: number }): { min: number; max: number; step: number } {
  const tidy = (value: number) => Number(value.toPrecision(12));
  const { grain = 0, unit = 1, from, under = 0 } = options;
  if (!Number.isFinite(low) || !Number.isFinite(high)) return { min: 0, max: 1, step: 1 };
  const start = from ?? Math.min(0, low);
  const end = Math.max(0, high);
  const span = end - start;
  if (!(span > 0)) return { min: 0, max: 1, step: 0.5 };
  const most = clampNumber(Math.floor(options.steps), 2, MOST_LABELS - 1);
  const decimalsOfStep = (step: number) => [0, 1, 2, 3].find((decimals) => exactAt(step / unit, decimals)) ?? 4;
  let best: { min: number; max: number; step: number; cost: number } | undefined;
  for (let power = Math.floor(Math.log10(span / (2 * MOST_LABELS))) - 1; power <= Math.ceil(Math.log10(span)) + 1; power++) {
    for (const base of [1, 2, 2.5, 5]) {
      const step = tidy(power >= 0 ? base * 10 ** power : base / 10 ** -power);
      if (grain > 0 && Math.abs(step / grain - Math.round(step / grain)) > 1e-6) continue;
      const decimals = decimalsOfStep(step);
      let min = start < 0 || from !== undefined ? tidy(Math.floor(start / step + 1e-9) * step) : 0;
      const over = (part: number) => tidy(end > 0 ? Math.ceil(end / part - 1e-9) * part : 0);
      let max = over(step / 2);
      if (Math.floor((max - min) / step + 1e-9) > most) max = over(step / 10);
      max = tidy(Math.max(max, min + 2 * step));
      // Room under the lowest value, where the number at the end of a bar under nothing stands.
      for (let lower = 0; low < 0 && under > 0 && lower < 3 && (low - min) / (max - min) < under; lower++) min = tidy(min - step);
      const steps = Math.floor((max - min) / step + 1e-9);
      if (steps > most) continue;
      const labelled = Math.abs(min + steps * step - max) < 1e-9 * Math.max(1, Math.abs(max));
      const cost = 1 - span / (max - min) + 0.1 * decimals + 0.01 * (steps + 1) + (labelled ? 0 : 0.01) + (base === 2.5 ? 0.005 : 0);
      if (!best || cost < best.cost - 1e-9) best = { min, max, step, cost };
    }
  }
  if (best) return { min: best.min, max: best.max, step: best.step };
  // No step gives three labels (a span of one grain): the axis has two, at its ends.
  const { max, step } = roundAxis(span, 1, grain);
  const min = start < 0 || from !== undefined ? tidy(Math.floor(start / step + 1e-9) * step) : 0;
  return { min, max: tidy(Math.max(min + max, end)), step };
}
const clampNumber = (value: number, low: number, high: number) => Math.min(high, Math.max(low, value));

/** The labels of a value axis: its start and every step up to its end. */
export function axisLabels(axis: { min?: number | undefined; max: number; step?: number | undefined }): number[] {
  const min = axis.min ?? 0;
  if (axis.step === undefined || !(axis.step > 0)) return [min, axis.max];
  const step = axis.step;
  const count = Math.min(50, Math.floor((axis.max - min) / step + 1e-9));
  return Array.from({ length: count + 1 }, (_, at) => Number((min + at * step).toPrecision(12)));
}

/**
 * The room a chart keeps free around its plot, in points, counted from its text size: for the
 * names of its categories, for its value axis when one is shown (`axis`: how many characters
 * its longest label has), for the numbers at its bars when they are written (`numbers`: how
 * many characters the longest has), and for a legend of its own. Left to itself PowerPoint
 * keeps more than a small chart can spare, and the plot is under half of its frame.
 */
export function plotMargins(chart: { horizontal: boolean; size: number; names: number; axis?: number | undefined; numbers?: number | undefined; legend: boolean }): { left: number; right: number; top: number; bottom: number } {
  const s = chart.size;
  const legend = chart.legend ? 1.9 * s : 0;
  if (chart.horizontal) {
    return {
      // PowerPoint breaks a name it is not given its own width and a size and four fifths more:
      // with less, Q1 is set as Q over 1.
      left: (chart.names * 0.55 + 1.8) * s,
      right: 0.5 * s + (chart.numbers !== undefined ? (chart.numbers * 0.55 + 0.6) * s : 0),
      top: 0.4 * s,
      bottom: (chart.axis !== undefined ? 1.8 * s : 0.4 * s) + legend,
    };
  }
  return {
    left: chart.axis !== undefined ? (chart.axis * 0.55 + 1) * s : 0.4 * s,
    right: 0.5 * s,
    // A number over a column takes a line and a little air; a name under it the same, as PowerPoint
    // sets them: with less, the names are cut at the foot of the frame. A legend under the names
    // leaves them its own air; without one the names keep a little more, or their descenders
    // (the g of Spring) are cut.
    top: chart.numbers !== undefined ? 1.3 * s : 0.7 * s,
    bottom: chart.legend ? 1.7 * s + legend : 2.1 * s,
  };
}

/** The colour of a series or a slice: the template's chart colours in turn, and past their end the same colours again as tints towards the ground, so that no two neighbours are alike. */
export function seriesColor(colors: readonly string[], index: number, ground: string | undefined): string {
  const base = colors[index % colors.length] as string;
  const round = Math.floor(index / colors.length);
  return ground === undefined || round === 0 ? base : mix(hex(ground), hex(base), Math.min(0.75, 0.45 * round));
}

/**
 * The language a chart's text is marked with, from its names of categories and series. In a deck
 * that reads right to left, names with a right-to-left letter are marked with that language
 * (`he-IL`), so that PowerPoint places their punctuation and brackets as Hebrew; any other chart is
 * marked with the deck's left-to-right language (`en-US` for a Hebrew deck, or for one that states
 * no language), as every chart was before 0.7.0. One mark serves the whole chart: its text has one
 * paragraph style for all its labels.
 */
export function chartLanguage(chart: Pick<ChartValue, 'categories' | 'series'>, direction: Direction, language?: string): string {
  const names = [...chart.categories, ...chart.series.map((series) => series.name)].join(' ');
  return marksRightToLeft(names, direction) ? rightToLeftTag(names) : leftToRightTag(language);
}

/** How the text and the order of a chart are written, beyond its slot. */
export interface ChartWriting {
  /**
   * The chart reads right to left: the first category stands on the right of columns and lines,
   * with their value axis on the right, and horizontal bars grow from the right, with their names
   * on the right. The room kept around the plot, and the legend, are mirrored with it. A pie is
   * drawn as it is. Left out, a chart reads left to right, whatever direction its text reads.
   */
  mirrored?: boolean;
  /** The deck's language, which names the language of a chart whose names are left to right (see {@link chartLanguage}). */
  language?: string;
}

/** The chart part. `workbookRelId` is the relationship from the chart to its workbook; `direction` is the direction its text reads. */
export function buildChartXml(chart: ChartValue, slot: ChartSlotDef, workbookRelId: string, direction: Direction = 'ltr', writing: ChartWriting = {}): string {
  const colors = slot.colors && slot.colors.length > 0 ? slot.colors : DEFAULT_COLORS;
  const drawn = slot.drawn;
  const color = (index: number) => (drawn ? seriesColor(colors, index, drawn.ground ?? 'FFFFFF') : (colors[index % colors.length] as string));
  const lineColor = slot.lineColor ?? 'D9D9D9';
  const percent = chart.format === 'percent';
  const values = knownValues(chart);
  const mirrored = writing.mirrored === true && chart.type !== 'pie';
  // A chart in English is marked as it always was: at the end of its paragraph only.
  const tag = chartLanguage(chart, direction, writing.language);
  const ownTag = tag === 'en-US' ? '' : ` lang="${escapeXmlAttr(tag)}"`;
  const look = drawn ? (drawn.look ?? numberLook(values, percent)) : undefined;
  const format = look ? formatCode(look) : percent ? '0"%"' : 'General';
  const range = ranges(chart);
  const several = chart.series.length > 1;
  // A small chart shows less: `compact` drops the numbers on the bars and the gridlines, `spark`
  // also the axes and the legend, and draws thinner. What is left is the shape of the data.
  const detail = slot.detail ?? 'full';
  const full = detail === 'full';
  const spark = detail === 'spark';

  const seriesHead = (index: number) =>
    `<c:idx val="${index}"/><c:order val="${index}"/><c:tx>${stringRef(range.name(index), [chart.series[index]?.name ?? ''])}</c:tx>`;
  const seriesData = (index: number) =>
    `<c:cat>${stringRef(range.categories, chart.categories)}</c:cat>` +
    `<c:val>${numberRef(range.values(index), chart.series[index]?.values ?? [], format)}</c:val>`;

  const axisTicks = '<c:majorTickMark val="none"/><c:minorTickMark val="none"/>';
  const axisText = `${axisTicks}<c:tickLblPos val="nextTo"/>`;
  const axisLine = `<c:spPr>${line(lineColor)}</c:spPr>`;
  const gridlines = full ? `<c:majorGridlines><c:spPr>${line(lineColor)}</c:spPr></c:majorGridlines>` : '';
  // An axis that is not shown is still there: the plot names both of its axes.
  const shown = `<c:delete val="${spark ? 1 : 0}"/>`;
  const [categoryAxisId, valueAxisId] = [50010001, 50010002];
  // Bars are read by their length, so their axis starts at zero; left to itself PowerPoint
  // starts it near the smallest value, which exaggerates differences.
  const zeroBased = chart.type !== 'line' && values.every((value) => value >= 0);
  // Where its place is known, a chart says more by itself. Its value axis ends where the plan
  // says, the same for like charts of a slide, and is left out where the numbers stand at the
  // bars. The names of its categories stand upright, every one of them beside a horizontal bar
  // and every so many under columns; under a plot with values below nothing they stand at its foot.
  const numbers = drawn?.numbers ?? full;
  const size = slot.text?.size ?? 10;
  const below = values.some((value) => value < 0);
  const end = drawn?.axis;
  const valueShown = drawn && !spark ? `<c:delete val="${drawn.valueAxis === false ? 1 : 0}"/>` : shown;
  const scaling = end ? `<c:max val="${end.max}"/><c:min val="${end.min ?? 0}"/>` : zeroBased ? '<c:min val="0"/>' : '';
  const stepDecimals = look && end?.step !== undefined ? ([0, 1, 2, 3].find((decimals) => exactAt((end.step as number) / look.divisor, decimals)) ?? 2) : undefined;
  const axisFormat = look ? formatCode(look, Math.max(stepDecimals ?? look.decimals, 0), true) : format;
  const upright = `<c:txPr><a:bodyPr rot="0" vert="horz"/><a:lstStyle/><a:p><a:pPr${direction === 'rtl' ? ' rtl="1"' : ''}><a:defRPr${ownTag}/></a:pPr><a:endParaRPr lang="${escapeXmlAttr(tag)}"/></a:p></c:txPr>`;
  // A chart whose slide shows one legend for all its charts draws none of its own.
  const ownLegend = !spark && drawn?.legend !== false && (several || chart.type === 'pie');
  const widestOf = (list: readonly number[], decimals?: number) => (look ? list.reduce((most, value) => Math.max(most, shownAs(value, look, decimals).length), 1) : 1);
  /** The room around the plot of a chart whose place is known: see `plotMargins`. A chart that reads right to left has it mirrored, and its legend follows. */
  const room =
    drawn && chart.type !== 'pie'
      ? plotMargins({
          horizontal: chart.type === 'bar',
          size,
          names: chart.categories.reduce((most, category) => Math.max(most, category.length), 1),
          // The longest label of the axis: of those the plan gave it, and a character more than the longest value where PowerPoint picks them.
          axis: drawn.valueAxis !== false ? (end ? widestOf(axisLabels(end), stepDecimals) : widestOf(values) + 1) : undefined,
          numbers: numbers && chart.type !== 'line' ? widestOf(values) : undefined,
          legend: ownLegend,
        })
      : undefined;
  const margins = room && mirrored ? { ...room, left: room.right, right: room.left } : room;
  /**
   * Every how many categories one is named under columns: as many as have room side by side.
   * Each has the width of its place in the plot, which is what PowerPoint gives a name before
   * it breaks it.
   */
  const skip = (horizontal: boolean): number => {
    if (!drawn || horizontal) return 1;
    const longest = chart.categories.reduce((most, category) => Math.max(most, category.length), 1);
    const across = margins ? drawn.width - margins.left - margins.right : 0.9 * drawn.width;
    const each = Math.max(1, across) / Math.max(1, chart.categories.length);
    return Math.max(1, Math.ceil(((longest * NAME_WIDE + 0.5) * size) / each));
  };
  /**
   * Category and value axes. Horizontal bars list their categories top to bottom, as a table would.
   * A chart that reads right to left runs its categories from the right, which brings its value
   * axis to the right (that axis crosses at the first category); its horizontal bars grow from the
   * right, where their names then stand (that axis crosses at nothing).
   */
  const axes = (horizontal: boolean) =>
    `<c:catAx><c:axId val="${categoryAxisId}"/><c:scaling><c:orientation val="${horizontal || mirrored ? 'maxMin' : 'minMax'}"/></c:scaling>` +
    `${shown}<c:axPos val="${horizontal ? (mirrored ? 'r' : 'l') : 'b'}"/><c:numFmt formatCode="General" sourceLinked="0"/>${drawn ? `${axisTicks}<c:tickLblPos val="${below ? 'low' : 'nextTo'}"/>` : axisText}${axisLine}${drawn ? upright : ''}` +
    `<c:crossAx val="${valueAxisId}"/><c:crosses val="autoZero"/><c:auto val="1"/><c:lblAlgn val="ctr"/><c:lblOffset val="100"/>${drawn ? `<c:tickLblSkip val="${skip(horizontal)}"/>` : ''}<c:noMultiLvlLbl val="0"/></c:catAx>` +
    `<c:valAx><c:axId val="${valueAxisId}"/><c:scaling><c:orientation val="${horizontal && mirrored ? 'maxMin' : 'minMax'}"/>${scaling}</c:scaling>` +
    `${valueShown}<c:axPos val="${horizontal ? 'b' : mirrored ? 'r' : 'l'}"/>${gridlines}<c:numFmt formatCode="${escapeXmlAttr(axisFormat)}" sourceLinked="0"/>${axisText}` +
    `<c:spPr>${noLine}</c:spPr><c:crossAx val="${categoryAxisId}"/><c:crosses val="${horizontal ? 'max' : 'autoZero'}"/><c:crossBetween val="between"/>${end?.step !== undefined ? `<c:majorUnit val="${end.step}"/>` : ''}</c:valAx>`;
  const axisIds = `<c:axId val="${categoryAxisId}"/><c:axId val="${valueAxisId}"/>`;
  /**
   * The space between two groups of bars, as a share of a bar's width. Where the place is known
   * a bar is never thicker than it reads: a column takes at most 11% of the plot's width, and a
   * horizontal bar is at most 2.6 times as thick as the chart's text is high.
   */
  const gapWidth = (horizontal: boolean): number => {
    const usual = spark ? 40 : several ? 80 : 60;
    if (!drawn || spark) return usual;
    const count = Math.max(1, chart.categories.length);
    const along = horizontal ? 0.8 * drawn.height : 0.9 * drawn.width;
    const thickest = horizontal ? 2.6 * size : 0.11 * along;
    return Math.min(300, Math.max(usual, Math.ceil(100 * (along / (count * thickest) - chart.series.length))));
  };

  let plot: string;
  if (chart.type === 'pie') {
    const slices = chart.categories
      .map((_, index) => `<c:dPt><c:idx val="${index}"/><c:bubble3D val="0"/><c:spPr>${fill(color(index))}<a:ln w="${spark ? 6350 : 19050}">${fill('FFFFFF')}</a:ln></c:spPr></c:dPt>`)
      .join('');
    plot =
      `<c:pieChart><c:varyColors val="1"/><c:ser>${seriesHead(0)}${slices}${dataLabels(format, numbers ? 'value' : 'none', 'outEnd')}${seriesData(0)}</c:ser>` +
      `<c:firstSliceAng val="0"/></c:pieChart>`;
  } else if (chart.type === 'line') {
    const series = chart.series
      .map(
        (_, index) =>
          `<c:ser>${seriesHead(index)}<c:spPr><a:ln w="${full ? 28575 : spark ? 12700 : 19050}" cap="rnd">${fill(color(index))}<a:round/></a:ln></c:spPr>` +
          (spark ? '<c:marker><c:symbol val="none"/></c:marker>' : `<c:marker><c:symbol val="circle"/><c:size val="${full ? 6 : 4}"/><c:spPr>${fill(color(index))}${noLine}</c:spPr></c:marker>`) +
          `${seriesData(index)}<c:smooth val="0"/></c:ser>`,
      )
      .join('');
    plot = `<c:lineChart><c:grouping val="standard"/><c:varyColors val="0"/>${series}${dataLabels(format, 'none')}<c:marker val="${spark ? 0 : 1}"/>${axisIds}</c:lineChart>${axes(false)}`;
  } else {
    const horizontal = chart.type === 'bar';
    const series = chart.series
      .map((_, index) => `<c:ser>${seriesHead(index)}<c:spPr>${fill(color(index))}</c:spPr><c:invertIfNegative val="0"/>${seriesData(index)}</c:ser>`)
      .join('');
    plot =
      `<c:barChart><c:barDir val="${horizontal ? 'bar' : 'col'}"/><c:grouping val="clustered"/><c:varyColors val="0"/>${series}` +
      `${dataLabels(format, numbers ? 'value' : 'none', 'outEnd')}<c:gapWidth val="${gapWidth(horizontal)}"/>${axisIds}</c:barChart>${axes(horizontal)}`;
  }

  const share = (value: number) => Math.round(value * 10000) / 10000;
  const inner = (x: number, y: number, w: number, h: number): string =>
    `<c:layout><c:manualLayout><c:layoutTarget val="inner"/><c:xMode val="edge"/><c:yMode val="edge"/><c:x val="${share(x)}"/><c:y val="${share(y)}"/><c:w val="${share(w)}"/><c:h val="${share(h)}"/></c:manualLayout></c:layout>`;
  /** Where the legend of a chart with a counted plot stands: on the line kept for it at the foot of the frame, under the middle of the plot. Left to itself PowerPoint sets it higher, over the names of the categories. */
  let legendAt = '';
  /** The plot of a chart whose place is known: the frame without the room its names, its axis, its numbers and its legend take. */
  const counted = (): string | undefined => {
    if (!drawn || !(drawn.width > 0 && drawn.height > 0)) return undefined;
    if (chart.type === 'pie') {
      // A small pie is a circle as large as its frame and its legend allow, at the top of the frame. A full one has its numbers around it, and is left to PowerPoint.
      if (numbers) return undefined;
      const side = Math.min(0.92 * drawn.width, 0.9 * (drawn.height - (ownLegend ? 3 * size : 0)));
      return side > 0 ? inner((drawn.width - side) / 2 / drawn.width, (0.04 * drawn.height) / drawn.height, side / drawn.width, side / drawn.height) : undefined;
    }
    if (!margins) return undefined;
    const w = drawn.width - margins.left - margins.right;
    const h = drawn.height - margins.top - margins.bottom;
    // A frame too small for what stands around the plot is left to PowerPoint.
    if (!(w >= 0.3 * drawn.width && h >= 0.25 * drawn.height)) return undefined;
    if (ownLegend) {
      // A key and a name for each series: about how wide the legend is, to set it under the middle of the plot.
      const wide = chart.series.reduce((sum, series) => sum + ((chart.type === 'line' ? 3.2 : 1.6) + series.name.length * 0.5 + 0.8) * size, 0);
      const x = Math.max(0, Math.min(drawn.width - wide, margins.left + (w - wide) / 2));
      legendAt = `<c:layout><c:manualLayout><c:xMode val="edge"/><c:yMode val="edge"/><c:x val="${share(x / drawn.width)}"/><c:y val="${share((drawn.height - 1.75 * size) / drawn.height)}"/></c:manualLayout></c:layout>`;
    }
    return inner(margins.left / drawn.width, margins.top / drawn.height, w / drawn.width, h / drawn.height);
  };
  // With no axis and no legend to leave room for, the plot takes the whole frame.
  const plotLayout = spark
    ? '<c:layout><c:manualLayout><c:layoutTarget val="inner"/><c:xMode val="edge"/><c:yMode val="edge"/><c:x val="0.04"/><c:y val="0.06"/><c:w val="0.92"/><c:h val="0.88"/></c:manualLayout></c:layout>'
    : (counted() ?? '<c:layout/>');
  const legend = ownLegend ? `<c:legend><c:legendPos val="b"/>${legendAt}<c:overlay val="0"/></c:legend>` : '';
  const style = slot.text ?? {};
  const font = style.font ? `<a:latin typeface="${escapeXmlAttr(style.font)}"/><a:ea typeface="${escapeXmlAttr(style.font)}"/><a:cs typeface="${escapeXmlAttr(style.font)}"/>` : '';
  const defaults =
    // The paragraph direction makes mixed labels such as "רבעון 1" read in the right order.
    `<c:txPr><a:bodyPr/><a:lstStyle/><a:p><a:pPr${direction === 'rtl' ? ' rtl="1"' : ''}><a:defRPr sz="${Math.round((style.size ?? 10) * 100)}"${style.bold ? ' b="1"' : ''}${ownTag}>` +
    `${style.color ? fill(style.color) : ''}${font}</a:defRPr></a:pPr><a:endParaRPr lang="${escapeXmlAttr(tag)}"/></a:p></c:txPr>`;

  return (
    `${XML_HEADER}<c:chartSpace xmlns:c="${NS_C}" xmlns:a="${NS_A}" xmlns:r="${NS_R}">` +
    `<c:date1904 val="0"/><c:roundedCorners val="0"/>` +
    `<c:chart><c:autoTitleDeleted val="1"/><c:plotArea>${plotLayout}${plot}<c:spPr><a:noFill/>${noLine}</c:spPr></c:plotArea>${legend}` +
    `<c:plotVisOnly val="1"/><c:dispBlanksAs val="gap"/></c:chart>` +
    `<c:spPr><a:noFill/>${noLine}</c:spPr>${defaults}` +
    `<c:externalData r:id="${escapeXmlAttr(workbookRelId)}"><c:autoUpdate val="0"/></c:externalData>` +
    `</c:chartSpace>`
  );
}

/**
 * The data as an Excel workbook: categories down column A, one series per column from B, names in
 * row 1. A value that is not known is an empty cell. `rightToLeft` sets the sheet right to left,
 * as Excel shows a sheet in Hebrew: for a chart that reads right to left.
 */
export function buildChartWorkbook(chart: ChartValue, options: { rightToLeft?: boolean } = {}): Uint8Array {
  const stringCell = (reference: string, value: string) => `<c r="${reference}" t="inlineStr"><is><t xml:space="preserve">${text(value)}</t></is></c>`;
  const header = `<row r="1">${chart.series.map((series, index) => stringCell(`${columnName(index + 1)}1`, series.name)).join('')}</row>`;
  const rows = chart.categories
    .map((category, row) => {
      const cells = chart.series
        .map((series, index) => {
          const value = series.values[row];
          return value === null || value === undefined ? '' : `<c r="${columnName(index + 1)}${row + 2}"><v>${value}</v></c>`;
        })
        .join('');
      return `<row r="${row + 2}">${stringCell(`A${row + 2}`, category)}${cells}</row>`;
    })
    .join('');

  const main = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';
  const packageRels = 'http://schemas.openxmlformats.org/package/2006/relationships';
  const files: Record<string, string> = {
    '[Content_Types].xml':
      `${XML_HEADER}<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">` +
      `<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>` +
      `<Default Extension="xml" ContentType="application/xml"/>` +
      `<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>` +
      `<Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>` +
      `</Types>`,
    '_rels/.rels':
      `${XML_HEADER}<Relationships xmlns="${packageRels}">` +
      `<Relationship Id="rId1" Type="${NS_R}/officeDocument" Target="xl/workbook.xml"/></Relationships>`,
    'xl/workbook.xml':
      `${XML_HEADER}<workbook xmlns="${main}" xmlns:r="${NS_R}">${options.rightToLeft ? '<bookViews><workbookView/></bookViews>' : ''}<sheets><sheet name="${SHEET}" sheetId="1" r:id="rId1"/></sheets></workbook>`,
    'xl/_rels/workbook.xml.rels':
      `${XML_HEADER}<Relationships xmlns="${packageRels}">` +
      `<Relationship Id="rId1" Type="${NS_R}/worksheet" Target="worksheets/sheet1.xml"/></Relationships>`,
    'xl/worksheets/sheet1.xml': `${XML_HEADER}<worksheet xmlns="${main}">${options.rightToLeft ? '<sheetViews><sheetView rightToLeft="1" workbookViewId="0"/></sheetViews>' : ''}<sheetData>${header}${rows}</sheetData></worksheet>`,
  };
  // A fixed timestamp, so the same chart always gives the same bytes.
  return zipSync(Object.fromEntries(Object.entries(files).map(([name, content]) => [name, strToU8(content)])), { mtime: ZIP_EPOCH });
}

/** The frame that puts a chart on a slide, in place of the sample picture. Sizes are in EMU. */
export function chartFrameXml(shapeId: string, relId: string, area: { x: number; y: number; cx: number; cy: number }, alt: string): string {
  return (
    `<p:graphicFrame xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="${NS_A}" xmlns:r="${NS_R}">` +
    `<p:nvGraphicFramePr><p:cNvPr id="${escapeXmlAttr(shapeId)}" name="Chart" descr="${escapeXmlAttr(alt)}"/>` +
    `<p:cNvGraphicFramePr><a:graphicFrameLocks noGrp="1"/></p:cNvGraphicFramePr><p:nvPr/></p:nvGraphicFramePr>` +
    `<p:xfrm><a:off x="${Math.round(area.x)}" y="${Math.round(area.y)}"/><a:ext cx="${Math.round(area.cx)}" cy="${Math.round(area.cy)}"/></p:xfrm>` +
    `<a:graphic><a:graphicData uri="${CHART_GRAPHIC_URI}"><c:chart xmlns:c="${NS_C}" r:id="${escapeXmlAttr(relId)}"/></a:graphicData></a:graphic>` +
    `</p:graphicFrame>`
  );
}
