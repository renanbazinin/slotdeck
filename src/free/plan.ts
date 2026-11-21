// A free slide as the rest of the software sees it: an ordinary pattern with slots, and the list
// of shapes to make for it. The blocks are set out (layout.ts), each kind is drawn inside its box
// by the rules below, and every text gets a size at which it fits. Validation, measuring, the
// writer and the renderer then work on the pattern as on any other.
//
// Lengths in this file are points on the slide.
import { contrast, mix, readable } from '../color/space.ts';
import { FREE_PATTERN } from '../deck/spec.ts';
import type { FreeArrange, FreeBlockKind, SlideSpec, SlotValue } from '../deck/spec.ts';
import { isChartValue, isTableValue, isTextValue, knownValues, plainText, toParagraphs } from '../deck/spec.ts';
import type { ChartValue } from '../deck/spec.ts';
import { headingPlaces, headingsOver } from '../deck/headings.ts';
import type { HeadingOver } from '../deck/headings.ts';
import { NAME_WIDE, numberLook, plotMargins, roundSpan, seriesColor, shownAs } from '../render/chart.ts';
import { headingRoom, runInChars, runsIn } from '../render/text.ts';
import { tableTextSize } from '../render/table.ts';
import type { ChartDrawn, ChartSlotDef, Direction, ImageSlotDef, PatternDef, SlideBox, SlotDef, TableSlotDef, TextSlotDef } from '../template/manifest.ts';
import { FREE_PARTS, describeBlocks, freeSlotName } from './blocks.ts';
import type { FreeKit, FreeTextStyle } from './kit.ts';
import { arrangeBlocks } from './layout.ts';
import type { FreeLayout, LaidBlock } from './layout.ts';

/** A shape made for a free slide. Text shapes, panels and rules are copies of the template's own. */
export interface MadeShape {
  id: string;
  name: string;
  kind: 'text' | 'panel' | 'rule' | 'picture' | 'anchor' | 'table' | 'key';
  /** The template shape it is a copy of. A table without one is drawn by the code. */
  from?: { slide: number; shapeId: string };
  box: SlideBox;
  /** Text: its size in points and what the copy is given. */
  size?: number;
  bold?: boolean;
  italic?: boolean;
  /** The colour of a text, or the fill of a key: the small square that stands for a series in the legend of a slide. */
  color?: string;
  /** `own`: the source's bullets, set for the new size. `add`: bullets the code makes. `none`: the source has bullets and the copy has none. */
  bullets?: 'own' | 'add' | 'none';
  align?: 'start' | 'ctr';
  anchor?: 't' | 'ctr' | 'b';
  /** Lines break at the box's edge. A big number does not break. */
  wrap?: boolean;
  /** Text in this many columns. */
  columns?: number;
  /** Space kept free at the sides of the text, in points. */
  inset?: number;
  /** Text the code writes itself: the number of a step. */
  text?: string;
  /** The text can take several lines: they are set at least as far apart as its size asks, whatever the source's own line spacing. */
  lines?: 'several';
  /** The line spacing the copy is drawn with where its spacing gave way (`GIVE`): a share of the usual in whole percents (1.01 for 101%), or a height in whole points. Left out, its source's own. */
  lineSpacing?: { percent: number } | { points: number };
  /** Picture: the template picture whose outline it takes. */
  outline?: { slide: number; shapeId: string };
  /** A table the code draws, on a ground its usual look does not suit: the fill of every second row (`''` for none) and the colour of its lines. */
  look?: { band: string; line: string };
}

/**
 * A text slot of a free slide: it says which template shape its text is styled as, and how the
 * copy is drawn where that differs from the shape, so that text is measured as it is set.
 */
export type FreeTextSlot = TextSlotDef & { styleOf: { slide: number; shapeId: string }; drawn: FreeDrawn };

/** How the copy of a template text is drawn on a free slide. Lengths in points. */
export interface FreeDrawn {
  /** The space kept free at each side of the text. Above and below there is none. */
  inset: number;
  /** Lines break at the box's edge. */
  wrap: boolean;
  /** How far the lines hang behind a bullet; 0 for a copy that has no bullet and no margin, whatever its source has. Left out, the source's own margins. */
  hang?: number;
  /** Text in several columns: how many, the space between two of them, and the height the text may take in each, which can be more than the box drawn for it. */
  columns?: number;
  gap?: number;
  height?: number;
  /** The line spacing the copy is drawn with where it is not its source's (`MadeShape.lineSpacing`). */
  lineSpacing?: { percent: number } | { points: number };
}

/**
 * The table slot of a free slide. It takes the box it was given, and states what its room holds
 * with the text at four fifths of the size the table's rows alone would have, and never under
 * the smallest size a table is set in. `rowHeights` is there when the rows are not all
 * alike: the height each needs for the lines of its fullest cell, in points, the heading row
 * first. The rows share the height of the box in these proportions.
 */
export type FreeTableSlot = TableSlotDef & {
  box: SlideBox;
  rowHeights?: number[];
  /** The width each column asks for by what is written in it, in points: the columns share the table's width in these proportions. Left out, they share it equally. */
  columnWidths?: number[];
  /** How the cells of each column are aligned where that is not at the start: `'r'` for a column of numbers in a deck read from the left. */
  columnAligns?: Array<'l' | 'r' | undefined>;
};

export interface FreePlan {
  pattern: PatternDef;
  shapes: MadeShape[];
  layout: FreeLayout;
  /** The smallest text on the slide, in points, and the least it should be to read well. */
  smallest: number;
  comfortable: number;
}

export interface FreeContext {
  /** The deck's reading direction. */
  direction: Direction;
  /** The deck's factor on body text. */
  textScale?: number;
  /**
   * No content yet: the slide is being planned. Every text is taken to be as long as texts of
   * its kind usually are (`USUAL`), the slide is set out and its sizes are found as for content
   * of that length, and every budget says what its box holds at the size that ends on. Optional
   * parts are counted in by rule: an icon when `icons` says cards get one, a heading over a text
   * when the slide has several blocks, a caption when it has several of that kind, the source of
   * a quotation. `withParts` names slots to count in besides: the optional parts a slide that is
   * being changed already has, such as `b2Caption`. `lengths` gives some texts the length a
   * person asked for, in characters, by slot name (`b1Text`): such a text is set out and sized
   * for that length in place of the usual one, and its budget says what its box then holds.
   * `paragraphs` gives some of those texts the number of paragraphs they will have, by slot name:
   * such a text is counted in that many paragraphs, each ending on a line of its own, in place of
   * paragraphs of about 300 characters. `asIs` gives, of a slide a person asked to have whole
   * (`arrange.dense`), the lengths of the paragraphs of a text code sets as it is (a document's,
   * its headings among them), by slot name: such a text is judged as it will be set. `heads` gives
   * of such a text the headings that stand over a paragraph (`headingsOver`): each is counted as it
   * is set, run in at the head of that paragraph or on a line of its own.
   */
  planning?: {
    icons: boolean;
    withParts?: readonly string[];
    lengths?: Readonly<Record<string, number>>;
    paragraphs?: Readonly<Record<string, number>>;
    asIs?: Readonly<Record<string, readonly number[]>>;
    heads?: Readonly<Record<string, readonly HeadingOver[]>>;
  };
  /**
   * For a deck that asks for larger text: the size every text of the slide has without that
   * wish, by the name of its shape, or of its slot for the text of a chart. A heading, a
   * running text, a label, a quotation, a caption or the text of a chart that would come out
   * smaller than that is set at that size. Whether the slide can be made so is for the caller
   * to say (resolve.ts): this only keeps a half point lost to rounding, or to a block that got
   * narrower, from costing the slide its larger text.
   */
  atLeast?: ReadonlyMap<string, number>;
  /**
   * Counts running text as 0.6.6 counted it: a document's headings each on a line of their own,
   * and a character no narrower than the estimate (`WIDE`, `DENSE_WIDE`) whatever width the
   * template states. For the sizes that a deck asking for larger text must not go under
   * (resolve.ts): counted as now, a text that gained size from the count made the wish cost the
   * rest of its slide what 0.6.6 gave it.
   */
  asThen?: boolean;
  /**
   * Counts as 0.6.8 counted: no spacing gives way (`GIVE`), and a short text is set down its room
   * as it was. For the plans among which a deck asking for larger text chooses how much of the
   * wish a slide takes (resolve.ts): chosen among plans whose texts closer spacing lifts, a slide
   * took the whole wish with a text lifted just to the least, smaller than 0.6.8 set it.
   */
  ownSpacing?: boolean;
}

const clamp = (value: number, low: number, high: number) => Math.min(high, Math.max(low, value));
const half = (size: number) => Math.round(size * 2) / 2;
const halfDown = (size: number) => Math.floor(size * 2) / 2;
/**
 * How text is counted without a font to measure it with: a character is about 0.55 of the size
 * wide, a line 1.2 of it high, and where lines break at words about 85% of each is used.
 */
const WIDE = 0.55;
const LINE = 1.2;
const FULL = 0.85;
/**
 * The last line of a text may hang under its box by this share of its size, the space under its
 * letters, and by no more than so many points on a slide 13.33 inches wide: a text that was
 * fitted never ends further than that under its box, whatever its size.
 */
const HANG = 0.25;
const MOST_HANG = 3;
/** How far a last line may hang under its box, in points, by the size of the text. */
type Hang = (size: number) => number;

const perLine = (width: number, size: number, wide = WIDE) => Math.max(1, Math.floor(width / (wide * size)));
/**
 * How high a line of a text is at a size, in points: 1.2 times the size, or more where the
 * template sets the lines of that text further apart. The copy keeps the line spacing of the
 * shape it is made from, so a text of a template with a line spacing of 130% takes 1.56 times
 * its size for every line, and is counted so. Lines set closer than usual are counted as usual.
 */
function lineOf(style: Pick<FreeTextStyle, 'leading'>, size: number, give = 0): number {
  const leading = style.leading;
  const own = !leading ? LINE * size : Math.max(LINE * size, 'points' in leading ? leading.points : LINE * size * leading.percent);
  if (give > 0) {
    // Closed up, never higher than its own: lines in whole points at about single have nothing to give.
    const drawn = drawnLeading(style, size, give);
    return Math.min(own, 'points' in drawn ? drawn.points : LINE * size * drawn.percent);
  }
  return own;
}
/**
 * 0.6.9: how far the spacing of running text and lists gives way before its size does. Text
 * that does not fit at a size is first set with the space between its paragraphs closed up, to
 * 0.7 of what it is (`GAP_GIVE`), and its lines by up to a tenth (`LINE_GIVE`), never closer
 * than single: a template that sets its lines wider (112%) gives some of that up, one at single
 * gives none. In PowerPoint's own rendering, paragraphs closer than 0.7 of the usual space
 * stopped reading as apart. `GIVE` are the steps, loosest first: a text is set at the largest
 * size at which one of them fits it, and at the loosest of them that does.
 */
const GAP_GIVE = 0.3;
const LINE_GIVE = 0.1;
const GIVE: readonly number[] = [0, 0.25, 0.5, 0.75, 1];
const NO_GIVE: readonly number[] = [0];
/**
 * The line spacing a copy is drawn with where its spacing gives way so far (`give`, 0 to 1): the
 * spacing it has, at least single, less up to a tenth, in whole percents or whole points, as
 * PowerPoint draws them, and never closer than single.
 */
function drawnLeading(style: Pick<FreeTextStyle, 'leading'>, size: number, give: number): { percent: number } | { points: number } {
  const leading = style.leading;
  if (leading && 'points' in leading) {
    const full = Math.max(leading.points, LINE * size);
    return { points: Math.max(Math.ceil(LINE * size - 1e-9), Math.ceil(full * (1 - LINE_GIVE * give) - 1e-9)) };
  }
  const full = Math.max(1, leading?.percent ?? 1);
  return { percent: Math.max(100, Math.ceil(100 * full * (1 - LINE_GIVE * give) - 1e-9)) / 100 };
}
/** The height of so many lines of a text at a size, with `more` lines of the usual height on top: the space kept around a text follows its size, whatever the line spacing. */
const linesHigh = (style: Pick<FreeTextStyle, 'leading'>, lines: number, size: number, more = 0): number => (style.leading ? lines * lineOf(style, size) + more * LINE * size : (lines + more) * LINE * size);
/**
 * How many characters a box holds at a size: whole lines, with room lost where lines break at
 * words. A box lower than a line holds no line: text is never fitted into one, and what is
 * written for it is reported. `line` is the height of a line where it is not the usual one, and
 * `wide` how wide a character is counted where it is not `WIDE` (a text set dense).
 */
function capacity(width: number, height: number, size: number, hang: Hang, paragraphs = 1, line = LINE * size, wide = WIDE): number {
  const across = perLine(width, size, wide);
  const down = linesHeld(height, size, line, hang);
  return Math.max(1, Math.floor(across * down * FULL - (Math.max(1, paragraphs) - 1) * across * 0.5));
}
/** How many lines a box holds at a size. */
const linesHeld = (height: number, size: number, line: number, hang: Hang): number => Math.max(0, Math.floor((height + hang(size)) / line));
/** Whether a box is high enough for one line of a text at a size. */
const holdsLine = (height: number, size: number, line: number, hang: Hang): boolean => linesHeld(height, size, line, hang) >= 1;
/**
 * How wide PowerPoint sets a character of ordinary text, as a share of its size: 0.40 to 0.43 in
 * its own rendering of two templates, one in Latin and one in Hebrew letters. The estimate above
 * (0.55) is wider on purpose: it leaves a text room for wide letters, and every budget is stated
 * with it.
 */
const SET_WIDE = 0.42;
/** How many characters PowerPoint sets on one line of a width. */
const setAcross = (width: number, size: number): number => Math.max(1, Math.floor(width / (SET_WIDE * size)));
/** How many characters so many lines hold when PowerPoint sets them: every line but the last 85% full, the last one as full as it gets. */
const setHolds = (lines: number, width: number, size: number): number => Math.floor(setAcross(width, size) * (FULL * (lines - 1) + 1));
/** The fewest lines PowerPoint can set a text of so many characters in: every line full. */
const setLines = (chars: number, width: number, size: number): number => Math.max(1, Math.ceil(chars / setAcross(width, size)));
/**
 * How many characters the text box of a card holds at a size. The text of cards and steps is
 * fitted with this. Every other text, and every budget that is stated before there is content,
 * is counted with `capacity` at the real height of its lines.
 *
 * On a template that sets its lines as far apart as usual the two are the same. Where the lines
 * are set further apart, the length of a line and its height are counted the same way here:
 * both as the estimate has them (a character 0.55 of the size wide, a line 1.2 of it high), and
 * both as PowerPoint sets them (0.42, and the height the template gives the line). A text has to
 * pass both counts. The estimate is generous in width and, on such a template, short in height,
 * and the two cancel: in PowerPoint's own rendering of a template with a line spacing of 130% a
 * card text is about as high as the estimate says. Counting the height as it is and the width
 * by the estimate set the same texts a fifth smaller, in cards a third empty. The second count
 * is there for what the estimate cannot know: a box holds no more than its lines, as high as
 * they really are, hold (`setHolds`), and a box lower than one such line holds none.
 */
function cardCapacity(width: number, height: number, size: number, style: Pick<FreeTextStyle, 'leading'>, hang: Hang): number {
  const usual = capacity(width, height, size, hang);
  const line = lineOf(style, size);
  if (!(line > LINE * size)) return usual;
  const down = linesHeld(height, size, line, hang);
  if (down < 1) return 1;
  return Math.max(1, Math.min(usual, setHolds(down, width, size)));
}
/**
 * How high the text of a card is at a size, without the space kept under it: the lines the
 * estimate gives it at the usual height and, where the template sets its lines further apart,
 * no less than the fewest lines PowerPoint can set it in, at the height the template sets
 * them. See `cardCapacity`.
 */
function cardTextHigh(chars: number, width: number, size: number, style: Pick<FreeTextStyle, 'leading'>): number {
  const lines = Math.max(1, Math.ceil(chars / (perLine(width, size) * FULL)));
  const line = lineOf(style, size);
  if (!(line > LINE * size)) return lines * LINE * size;
  return Math.max(lines * LINE * size, setLines(chars, width, size) * line);
}

/**
 * The lines a short text takes when its words are set one after another on lines of so many
 * characters: a word that does not fit the rest of a line starts the next, and a word longer
 * than a line is broken. A heading of three long words takes three lines where its length
 * alone comes to two.
 */
function wrapLines(text: string, across: number): number {
  const most = Math.max(1, across);
  let lines = 1;
  let used = 0;
  for (const word of text.split(/\s+/).filter((part) => part !== '')) {
    if (used > 0 && used + 1 + word.length <= most) {
      used += 1 + word.length;
      continue;
    }
    if (used > 0) lines++;
    const more = Math.ceil(word.length / most) - 1;
    lines += more;
    used = word.length - more * most;
  }
  return lines;
}

/** A text to fit: how long it is, the box it has, and the style whose line height it is set at. */
interface Fitted {
  chars: number;
  paragraphs: number;
  width: number;
  height: number;
  style: FreeTextStyle;
  /** The text of a card or a step: counted with `cardCapacity`. */
  card?: boolean;
  /** How long each paragraph is, for a text of several paragraphs or with bullets: its lines are counted paragraph by paragraph. */
  each?: readonly number[];
  /** Its lines hang behind a bullet, 1.1 times the size in. */
  bullets?: boolean;
  /** How wide a character of it is counted, as a share of its size: `WIDE` where left out, `DENSE_WIDE` for a text set dense. */
  wide?: number;
  /**
   * A text a person asked to have whole (`arrange.dense`), set in so many columns, one or more,
   * each `width` wide and `height` high: it is judged by `judged`, and holds what `heldIn` says.
   * Left out, a text in columns is counted as one column as high as all of them.
   */
  columns?: number;
  /**
   * Of such a text: its paragraphs are of any length, as a writer told a length and a number of
   * paragraphs may write them. It fits where every text of `chars` characters in `paragraphs`
   * paragraphs does (`heldIn`). Left out, its paragraphs are as long as `each` says.
   */
  split?: boolean;
  /**
   * The paragraphs, by position among `each`, that are headings set run in at a size: each takes
   * no line of its own, and its letters (counted bold) and the em space go to the head of the
   * paragraph under it, as the renderer sets them (`setHeadings`). Left out, none.
   */
  runIn?: (size: number) => ReadonlySet<number>;
  /** How far its spacing gives way (`GIVE`): 0 or left out, not at all. Every count of it reads this. */
  give?: number;
}
const NONE: ReadonlySet<number> = new Set();
/**
 * How long each paragraph of a text is as it is drawn at a size: a heading set run in
 * (`Fitted.runIn`) is 0, and the paragraph under it is longer by what the heading takes of its
 * line (`runInChars`).
 */
function drawnLengths(text: Fitted, size: number): { lengths: number[]; joined: ReadonlySet<number> } {
  const each = text.each ?? [text.chars];
  const joined = text.runIn?.(size) ?? NONE;
  if (joined.size === 0) return { lengths: [...each], joined };
  return { lengths: each.map((chars, at) => (joined.has(at) ? 0 : chars + (joined.has(at - 1) ? runInChars(each[at - 1] as number) : 0))), joined };
}
/** The characters and the paragraphs a text has as it is drawn at a size (`drawnLengths`): what its length is measured by. */
function drawnCount(text: Fitted, size: number): { chars: number; paragraphs: number } {
  if (!text.runIn) return { chars: text.chars, paragraphs: text.paragraphs };
  const { lengths, joined } = drawnLengths(text, size);
  if (joined.size === 0) return { chars: text.chars, paragraphs: text.paragraphs };
  return { chars: Math.ceil(lengths.reduce((sum, chars) => sum + chars, 0)), paragraphs: Math.max(1, text.paragraphs - joined.size) };
}
/** How many characters the box of a text holds at a size: of a text a person asked to have whole, in any `paragraphs` paragraphs (`heldIn`). */
const charsHeld = (text: Fitted, size: number, hang: Hang): number => {
  if (text.card) return cardCapacity(text.width, text.height, size, text.style, hang);
  if (text.columns !== undefined) return heldIn(text, size, hang, text.paragraphs);
  return capacity(text.width, text.height, size, hang, drawnCount(text, size).paragraphs, lineOf(text.style, size, text.give), text.wide);
};
/** How far the lines of a bulleted text hang behind the bullet, and the space between two paragraphs, at a size: what the copy is drawn with (build.ts) and what its slot states. */
const hangAt = (size: number): number => 1.1 * size;
const gapAt = (size: number, give = 0): number => {
  const full = Math.round(0.5 * size);
  // Closed up, in whole points as PowerPoint draws it, and never under 0.7 of what it is.
  return give > 0 ? Math.max(Math.ceil(0.7 * full - 1e-9), Math.round(full * (1 - GAP_GIVE * give))) : full;
};
/**
 * How many characters of a text a line holds at a size, in the width its bullet leaves, as its
 * lines are counted: by length, with room lost where lines break at words. A paragraph of one line
 * is counted so too: counted as a whole line of its length, 20 paragraphs of capitals or of a wide
 * face, each a little shorter than a line, were set in 40 lines and ran 150 pt past their box.
 */
const acrossOf = (text: Fitted, size: number): number => perLine(text.width - (text.bullets ? hangAt(size) : 0), size, text.wide) * FULL;
/** How many lines each paragraph of a text takes at a size: by its length (`acrossOf`), and at least one; a heading set run in takes none (`drawnLengths`). */
function linesEach(text: Fitted, size: number): number[] {
  const across = acrossOf(text, size);
  const { lengths, joined } = drawnLengths(text, size);
  return lengths.map((chars, at) => (joined.has(at) ? 0 : Math.max(1, Math.ceil(chars / across))));
}
/**
 * The height the paragraphs of a text take at a size: every paragraph in whole lines, in the
 * width its bullet leaves, with the space between two of them. A list of six points of a line
 * and a half each takes twelve lines, not the nine that its length alone comes to.
 */
function paragraphsHigh(text: Fitted, size: number): number {
  const each = linesEach(text, size);
  const lines = each.reduce((sum, count) => sum + count, 0);
  return lines * lineOf(text.style, size, text.give) + Math.max(0, each.length - 1) * gapAt(size, text.give);
}
/**
 * The height every line of a text needs in a column, in order, at a size: a line, and under the
 * line that ends a paragraph the space after it (`gapAt`), as PowerPoint sets them. The space after
 * a heading set run in goes before the paragraph it heads, and is counted with its first line.
 */
function lineNeeds(text: Fitted, size: number): number[] {
  const line = lineOf(text.style, size, text.give);
  const gap = gapAt(size, text.give);
  const needs: number[] = [];
  let before = 0;
  for (const lines of linesEach(text, size)) {
    if (lines === 0) {
      before += gap;
      continue;
    }
    for (let at = 0; at < lines; at++) needs.push(line + (at === 0 ? before : 0) + (at === lines - 1 ? gap : 0));
    before = 0;
  }
  return needs;
}
/**
 * How many columns of a height the lines of a text take (`lineNeeds`), as PowerPoint fills them:
 * a column takes the next line where it fits under the lines it has, and the line that ends a
 * paragraph only where the space after it fits too; a paragraph goes on in the next column, line
 * by line. In PowerPoint's own rendering of 96 decks of text in two and three columns, all 215
 * column breaks fell so. A text that its columns do not hold is set by PowerPoint in columns as
 * high as it needs, the first as long as the next, and runs under its box from the first column
 * on. `Infinity` where a line is higher than the column.
 */
function columnsTaken(needs: readonly number[], height: number): number {
  let columns = 1;
  let used = 0;
  for (const need of needs) {
    if (need > height + 1e-6) return Number.POSITIVE_INFINITY;
    if (used + need > height + 1e-6) {
      columns++;
      used = 0;
    }
    used += need;
  }
  return columns;
}
/**
 * The least height of the columns of a text in which they hold it (see `columnsTaken`), so that
 * they come out alike: one of the heights its runs of lines take.
 */
function columnsHigh(needs: readonly number[], columns: number): number {
  const total = needs.reduce((sum, need) => sum + need, 0);
  const least = Math.max(0, ...needs, total / columns);
  const heights = new Set<number>();
  for (let from = 0; from < needs.length; from++) {
    let sum = 0;
    for (let to = from; to < needs.length; to++) {
      sum += needs[to] as number;
      if (sum >= least - 1e-6) heights.add(sum);
    }
  }
  const sorted = [...heights].sort((a, b) => a - b);
  return sorted.find((height) => columnsTaken(needs, height) <= columns) ?? total;
}
/**
 * The room of a column of a text a person asked to have whole (`arrange.dense`), at a size: its
 * height less one whole line. PowerPoint does not always set its columns as they are counted: in
 * its own rendering, a text of two columns that fitted by a point or more came out now and then
 * with its first column a line longer, so that its last ran under its box, and words of wide
 * letters take more of a line than their length says. That line is kept free in every column.
 */
const columnRoom = (text: Fitted, size: number): number => text.height - lineOf(text.style, size, text.give);
/**
 * The one judge of a text a person asked to have whole (`arrange.dense`), in so many columns
 * (`Fitted.columns`, one or more): whether its paragraphs fit its box at a size, and how high its
 * columns come out. It counts by length, as every other text is counted: each paragraph takes
 * whole lines (`linesEach`), the lines go into the columns one after another as PowerPoint fills
 * them (`columnsTaken`), and each column holds its lines and the space after each paragraph in it
 * (`lineNeeds`) within its room, a line less than its height (`columnRoom`). It holds no more
 * characters than the box holds by its length alone (`capacity`), as was always so. No measure of
 * the words, of the kind of their letters or of a language counts: the face a template sets them
 * in is not known, and every such measure let some text pass that ran past its box in another.
 * `high` is how high its columns are where they come out alike (`columnsHigh`): its box is drawn
 * that high and a line more.
 */
function judged(text: Fitted, size: number, hang: Hang): { fits: boolean; high: number } {
  const columns = text.columns ?? 1;
  const needs = lineNeeds(text, size);
  const room = columnRoom(text, size);
  const drawn = drawnCount(text, size);
  const byLength = drawn.chars <= capacity(text.width, text.height * columns, size, hang, drawn.paragraphs, lineOf(text.style, size, text.give), text.wide);
  return { fits: room > 0 && byLength && columnsTaken(needs, room) <= columns, high: columnsHigh(needs, columns) };
}
/**
 * How many characters a text a person asked to have whole (`arrange.dense`) holds in its box at a
 * size, in any `paragraphs` paragraphs of any length: every such text passes `judged`. Each
 * paragraph is taken to lose up to a whole line where it ends (its lines are no more than its
 * length takes and one more each), with the space after it, and each column after the first to
 * lose a line and that space where it starts (`columnsTaken` moves a line that does not fit to
 * the next column). What a writer is told of such a text, with the number of paragraphs it is
 * told, is so much at most: a writer inside it is not sent back, however its paragraphs fall.
 */
function heldIn(text: Fitted, size: number, hang: Hang, paragraphs: number): number {
  const columns = text.columns ?? 1;
  const line = lineOf(text.style, size, text.give);
  const step = line + gapAt(size, text.give);
  const room = columnRoom(text, size);
  const count = Math.max(1, paragraphs);
  if (room < step) return 0;
  const lines = Math.floor((columns * room - (columns - 1) * step - count * step) / line);
  if (lines < 1) return 0;
  return Math.min(Math.floor(acrossOf(text, size) * lines), capacity(text.width, text.height * columns, size, hang, count, line, text.wide));
}
/**
 * The most paragraphs of a text a person asked to have whole (`arrange.dense`) that leave half the
 * room of its columns for its lines, after what each may lose where it ends (see `heldIn`), and
 * no fewer than one. A writer is told no more (`toldParagraphs`), and once written, a text of more
 * that does not fit is told so (`text has N paragraphs`).
 */
function roomyParagraphs(text: Fitted, size: number): number {
  const columns = text.columns ?? 1;
  const step = lineOf(text.style, size, text.give) + gapAt(size, text.give);
  const half = (columns * columnRoom(text, size) - (columns - 1) * step) / 2;
  return Math.max(1, Math.floor(half / step));
}
/**
 * The paragraphs a writer of a text a person asked to have whole (`arrange.dense`) is told where
 * its slide does not hold it as it is, no more than it has and than leave half its room for its
 * lines (`roomyParagraphs`). Each paragraph a writer is told may lose a line where it ends, and its
 * budget (`heldIn`) keeps that room free, so fewer paragraphs hold more. Where the slide holds all
 * of the text's characters in some number of paragraphs over twelve, the most such: its writer
 * joins as few as it can, as seventy short lines of a list are told 33. Where it holds them in no
 * more than twelve, the text is cut, and its writer is told as many as any writer of such a text
 * (`DENSE_PARAGRAPHS`): a slide of the 21 paragraphs of a document, its headings among them, holds
 * about a fifth less told 21 than told twelve.
 */
function toldParagraphs(text: Fitted, size: number, hang: Hang, paragraphs: number): number {
  const most = Math.min(paragraphs, roomyParagraphs(text, size));
  for (let count = most; count > DENSE_PARAGRAPHS; count--) if (tensDown(heldIn(text, size, hang, count)) >= text.chars) return count;
  return clamp(Math.min(most, DENSE_PARAGRAPHS), 1, Math.max(1, paragraphs));
}
/**
 * Whether a text fits its box at a size: by its length, and, where it has several paragraphs or
 * bullets, by the lines they take one by one. A text a person asked to have whole
 * (`Fitted.columns`) fits where `judged` says so, or, of any paragraphs (`Fitted.split`), where
 * its length is within what it holds in the paragraphs a writer is told (`heldIn`).
 */
function fitsAt(text: Fitted, size: number, hang: Hang): boolean {
  if (text.columns !== undefined) return text.split ? text.chars <= heldIn(text, size, hang, text.paragraphs) : judged(text, size, hang).fits;
  if (drawnCount(text, size).chars > charsHeld(text, size, hang)) return false;
  if (!text.each || (text.each.length < 2 && !text.bullets)) return true;
  return paragraphsHigh(text, size) <= text.height + hang(size);
}
/** A text with its spacing given way so far (`Fitted.give`). */
const giving = (text: Fitted, give: number): Fitted => (give > 0 ? { ...text, give } : text);
/** The loosest of the spacings `give` at which every text fits its box at a size; `undefined` where none does. */
function spacingAt(texts: Fitted[], size: number, hang: Hang, give: readonly number[] = NO_GIVE): number | undefined {
  for (const level of give) if (texts.every((text) => fitsAt(giving(text, level), size, hang))) return level;
  return undefined;
}
/** The largest size, in half points between `least` and `most`, at which every text fits its box at one of the spacings `give` (`GIVE`): by default its own. */
function fitSize(texts: Fitted[], most: number, least: number, hang: Hang, give: readonly number[] = NO_GIVE): number {
  for (let size = halfDown(most); size > least; size -= 0.5) {
    if (spacingAt(texts, size, hang, give) !== undefined) return size;
  }
  return Math.min(halfDown(most), half(least));
}

/**
 * How many characters a list of so many points holds in a box at a size, when the points are
 * about alike: every point has the same number of whole lines behind its bullet, the space
 * between two points is taken off the height, and a character is kept back for each point,
 * because points are never quite alike. A box with fewer lines than points holds one line each
 * for as many points as it has lines. It is never more than the box holds by the length of
 * the text alone (`capacity`): a list is fitted by both counts.
 */
function listCapacity(width: number, height: number, size: number, points: number, style: Pick<FreeTextStyle, 'leading'>, hang: Hang, wide = WIDE): number {
  const across = perLine(width - hangAt(size), size, wide) * FULL;
  const lines = linesHeld(height - Math.max(0, points - 1) * gapAt(size), size, lineOf(style, size), hang);
  const each = Math.floor(lines / points);
  const byPoints = each >= 1 ? points * (Math.floor(across * each) - 1) : Math.floor(across * lines);
  return Math.max(1, Math.min(byPoints, capacity(width, height, size, hang, points, lineOf(style, size), wide)));
}

/**
 * While a slide is planned there is no content. Every text is then taken to be as long as texts
 * of its kind usually are, in characters, paragraph by paragraph: a card or a step says about 70
 * characters, a quotation about 100, a list has about four points of 35, running text about 300,
 * a note 110, a figure five. The parts that are not named here are short and fit by their width
 * alone.
 */
const USUAL: Partial<Record<FreeBlockKind, Record<string, readonly number[]>>> = {
  card: { Text: [70] },
  step: { Text: [70] },
  quote: { Quote: [100] },
  list: { Items: [35, 35, 35, 35] },
  text: { Text: [300] },
  note: { Text: [110] },
  figure: { Value: [5] },
};
/** What a writer is told at the least for a text whose box holds no line: nobody can write to one character. */
const LEAST_SAID = 12;

/**
 * A length a person asked for, in characters, as the paragraphs of a usual text: paragraphs of
 * about 300 characters, as running text usually has them, and one for a shorter text. `count`
 * makes it so many paragraphs, as alike as they can be.
 */
function askedParagraphs(chars: number, count = Math.max(1, Math.round(chars / 300))): number[] {
  return Array.from({ length: count }, (_, at) => Math.floor(chars / count) + (at < chars % count ? 1 : 0));
}

/**
 * How wide a word is set, in sizes of its text: about half a size a letter, more for capitals and for
 * the widest letters, less for the narrowest. Generous, so that a word counted to fit a column does.
 */
const wordWide = (word: string): number =>
  [...word].reduce((sum, char) => sum + (/[MWmw@%]/.test(char) ? 0.88 : /[A-Z]/.test(char) ? 0.68 : /[ijlt.,:;'!|I]/.test(char) ? 0.32 : /[fr()]/.test(char) ? 0.4 : 0.56), 0);
/**
 * How wide PowerPoint sets text of ordinary weight, as a share of `wordWide`: 0.77 to 0.81 of it
 * on average in its own rendering of two templates, one in Latin and one in Hebrew letters, and
 * no more than this on nineteen lines of twenty. A heavy face (a heading, the title of a table)
 * is set at about the whole of it.
 */
const SET_SHARE = 0.85;
/**
 * The lines PowerPoint most likely sets a text in, in a width at a size: its words one after
 * another, each counted by the kind of its letters, and a word wider than a line broken. `share`
 * is how wide the face is set (`SET_SHARE` for ordinary text, 1 for a heavy one).
 */
function setLinesOf(text: string, width: number, size: number, share = SET_SHARE): number {
  const room = width / (share * size);
  const space = wordWide(' ');
  let lines = 1;
  let used = 0;
  for (const word of text.split(/\s+/).filter((part) => part !== '')) {
    const wide = wordWide(word);
    if (used > 0 && used + space + wide <= room + 1e-9) {
      used += space + wide;
      continue;
    }
    if (used > 0) lines++;
    const more = room > 0 ? Math.max(0, Math.ceil(wide / room - 1e-9) - 1) : 0;
    lines += more;
    used = wide - more * room;
  }
  return lines;
}

/**
 * A cell that is read as a number: one that starts with one, with a sign or a currency sign before
 * it (`4,120`, `-3%`, `6 days`). A column of them is aligned on its last characters, so that a
 * number in it does not run into the start of the next column.
 */
const NUMBER_CELL = /^[+\-\u2212]?[\u20AC$\u20AA\u00A3]?\s?\d/;
/** A number by itself, with its sign, separators and a percent sign. */
const NUMBER_WORD = /^[+\-\u2212]?[\u20AC$\u20AA\u00A3]?\d[\d.,:]*%?$/;
/**
 * The parts of a cell that do not break: its words, with a number and the short unit after it
 * taken as one (`2.4 minutes` does not stand as the number over its unit).
 */
const unbroken = (cell: string): string[] => {
  const words = cell.split(/\s+/).filter((word) => word !== '');
  const out: string[] = [];
  for (let at = 0; at < words.length; at++) {
    const word = words[at] as string;
    const next = words[at + 1];
    if (NUMBER_WORD.test(word) && next !== undefined && next.length <= 8 && !NUMBER_WORD.test(next)) {
      out.push(`${word} ${next}`);
      at++;
    } else out.push(word);
  }
  return out;
};

/** A slide of more charts than this draws them all bare: so many are a texture, and nobody reads their names. */
const MOST_DRAWN = 24;
/** Like charts of a slide share one value axis when the largest value of one is no more than this many times the largest of another. Further apart, the small one would be a row of stubs. */
const ALIKE = 5;

/** A chart value that can be drawn: names for its categories, and a number for each in every series. What is not is told by the deck's check. */
const drawable = (value: SlotValue | undefined): ChartValue | undefined => {
  if (!isChartValue(value) || value.categories.length === 0 || value.series.length === 0) return undefined;
  const sound = value.series.every((series) => Array.isArray(series?.values) && series.values.length === value.categories.length && series.values.every((number) => number === null || (typeof number === 'number' && Number.isFinite(number))));
  return sound && knownValues(value).length > 0 && value.categories.every((category) => typeof category === 'string') ? value : undefined;
};

/** The paragraphs of a text value. A value that is no text counts as empty: the deck's check reports it. */
const paragraphsOf = (value: SlotValue | undefined): string[] => (isTextValue(value) ? toParagraphs(value).filter((paragraph) => paragraph !== '') : []);
const present = (value: SlotValue | undefined): boolean => value !== undefined && value !== null && !(typeof value === 'string' && value.trim() === '') && !(Array.isArray(value) && value.length === 0);

/**
 * What to write for a budget, in round numbers: "Write 70 to 120 characters". The upper end is
 * never above the budget itself: text of that length is not reported as too long. A budget under
 * 40 characters has no such sentence: the limit of the slot says all there is to say.
 */
const target = (most: number): string | undefined => {
  if (most < 40) return undefined;
  const tens = (value: number) => Math.round(value / 10) * 10;
  return `Write ${tens(Math.round(most * 0.5))} to ${Math.min(tens(Math.floor(most * 0.95)), Math.floor(most / 10) * 10)} characters`;
};
/**
 * The description of a slot from its sentences. It says what the part is for and, for a text,
 * how much to write. It never repeats a limit that a field of the slot states (its characters,
 * its paragraphs, the categories of a chart, the rows of a table), and it has no full stop at
 * its end: whoever shows it to a writer puts the limits after it.
 */
const say = (...sentences: Array<string | undefined>): string | undefined => sentences.filter((sentence) => sentence !== undefined && sentence !== '').join('. ') || undefined;

/**
 * The lines a free slide's title of so many characters takes, as the plan counts them: at half
 * its size a character, the width a title's heavy face has, one to three.
 */
export function freeTitleLines(kit: FreeKit, chars: number): number {
  if (!kit.title) return 1;
  return clamp(Math.ceil((chars * 0.5 * kit.title.size) / (kit.title.box.w * (kit.slideWidth / 12700))), 1, 3);
}

/**
 * The least size running text is set at to read well, in points, on a slide this wide (EMU): 9 pt
 * on a slide 13.33 inches wide, and less on a narrower one, by its width; never under 5. A slide
 * whose smallest text is under it is told (`FreePlan.comfortable`), and a long text a person asked
 * to have whole is set no smaller (`FreeArrange.dense`).
 */
export const readableSize = (slideWidth: number): number => half(Math.max(5, 9 * (slideWidth / 12700 / 960)));
/** The least size the running text of a free slide is set at to read well on a template: its own (`TemplateManifest.leastTextSize`), or by its width (`readableSize`). */
export const leastSizeOf = (kit: Pick<FreeKit, 'leastTextSize' | 'slideWidth'>): number => kit.leastTextSize ?? readableSize(kit.slideWidth);
/**
 * How wide a character of a text set dense is counted, as a share of its size: as PowerPoint sets
 * a letter of ordinary text (`wordWide` of a small or a Hebrew letter, 0.56, by `SET_SHARE`), not
 * with the room to spare that `WIDE` keeps. In PowerPoint's own rendering of 3,035 lines of such
 * text at 6 to 17.5 pt, in Hebrew and in English, a character took 0.41 of its size on a middling
 * line, 0.44 on one line of twenty, and never more than 0.47. Lines still break at words (`FULL`).
 */
const DENSE_WIDE = 0.48;
/** A text set dense keeps its lines to about so many characters as PowerPoint sets them (`SET_WIDE`): two columns where they are no longer, three where they would be. */
const DENSE_LINE = 95;
/**
 * A text set dense that two columns do not hold at a size takes three where their lines hold so
 * many characters as PowerPoint sets them (`SET_WIDE`), as a column of a newspaper does: in two,
 * a text of a document was set a size smaller, in three columns with the last a third full.
 */
const DENSE_SHORT = 45;
/** The most paragraphs a writer is told for a text set dense: it is planned for any number up to it. */
const DENSE_PARAGRAPHS = 12;
/** The top of the range a writer is told for an asked length, as a share of it: `withLength` tells a text up to 1.15 times its share. */
const DENSE_TOP = 1.15;
/** A number of characters in round tens, rounded down, as a writer is told it (`tens` in length.ts). */
const tensDown = (chars: number): number => Math.max(10, Math.floor(chars / 10) * 10);

/**
 * The size the running text of a free slide has on a template when nothing makes it smaller: the
 * template's own, with the template's and the deck's factor on body text, and never over 0.65 of
 * the size of the title.
 */
export function freeBodySize(kit: FreeKit, deckScale = 1): number {
  const style = kit.styles.body;
  return Math.min(half((style.set?.size ?? style.size) * deckScale * (style.size <= 12 ? kit.textScale : 1)), 0.65 * (kit.title?.size ?? 28));
}

/** One legend for all the charts of a slide: the names of their series, and the size its text is set at, in points. */
interface SlideLegend {
  names: string[];
  size: number;
}

/**
 * Works a free slide out: where its blocks go, the shapes to make, and the pattern whose slots
 * hold its content. The same slide and template always give the same plan.
 */
export function freePlan(kit: FreeKit, slide: Pick<SlideSpec, 'blocks' | 'arrange' | 'slots' | 'headings'>, context: FreeContext): FreePlan {
  let plan = planWith(kit, slide, context, undefined, false) as Planned;
  // Card text longer than its budget that had to be drawn smaller to stay inside its card makes
  // the slide crowded: it is set out again in the larger room, where the cards are higher, and
  // that is taken where it gives some block more room. So is a long text a person asked to have
  // whole that is too long for four fifths of the template's size (`cramped`): it is set smaller
  // only there, in the larger room where the frame has one.
  let crowded = false;
  if ((plan.squeezed || plan.cramped) && !plan.layout.dense) {
    const again = planWith(kit, slide, context, undefined, true);
    if (again && (again.layout.dense || plan.cramped)) {
      plan = again;
      crowded = true;
    }
  }
  // A slide of like charts that all show the same series says once what the colours stand for:
  // one legend over the charts, and none in any of them, so that every plot has the room its
  // own legend took. The slide is set out again with a strip kept for it. Should a chart be
  // drawn bare in the room that leaves, the slide stays as it was, each chart with its legend.
  const legend = sharedLegend(kit, slide, context, plan);
  const chosen: Planned = (legend && planWith(kit, slide, context, legend, crowded)) ?? plan;
  delete chosen.squeezed;
  delete chosen.cramped;
  return chosen;
}

/** A plan, whether card text on it that is longer than its budget had to be drawn smaller, and whether a text set dense needs the larger room. */
type Planned = FreePlan & { squeezed?: boolean; cramped?: boolean };

/** The legend that the charts of a slide share, when they can share one: see `freePlan`. */
function sharedLegend(kit: FreeKit, slide: Pick<SlideSpec, 'blocks' | 'slots'>, context: FreeContext, plan: FreePlan): SlideLegend | undefined {
  const kinds = (slide.blocks ?? []) as FreeBlockKind[];
  if (context.planning || kinds[0] === 'note') return undefined;
  const charts = kinds.map((kind, position) => (kind === 'chart' ? drawable(slide.slots?.[freeSlotName(position + 1, 'Chart')]) : undefined));
  // Only a slide of charts, with or without a bottom line: among other blocks a legend at the top of the slide would stand far from its charts.
  if (kinds.some((kind, position) => kind !== 'note' && (kind !== 'chart' || charts[position] === undefined))) return undefined;
  const shown = charts.filter((chart): chart is ChartValue => chart !== undefined);
  const names = shown[0]?.series.map((series) => series.name.trim()) ?? [];
  if (shown.length < 2 || names.length < 2 || names.some((name) => name === '')) return undefined;
  if (shown.some((chart) => chart.type === 'pie' || chart.series.length !== names.length || chart.series.some((series, at) => series.name.trim() !== names[at]))) return undefined;
  const slots = plan.pattern.slots.filter((slot): slot is ChartSlotDef => slot.kind === 'chart');
  if (slots.some((slot) => slot.detail === 'spark')) return undefined;
  const k = kit.slideWidth / 12700 / 960;
  return { names, size: Math.max(half(8 * k), ...slots.map((slot) => slot.text?.size ?? 0)) };
}

/** The plan of a slide, with a strip kept at the top of its room for the legend of its charts when one is given. `undefined` when a chart would be drawn bare in what is left: such a slide has no legend. */
function planWith(kit: FreeKit, slide: Pick<SlideSpec, 'blocks' | 'arrange' | 'slots' | 'headings'>, context: FreeContext, legend: SlideLegend | undefined, crowded: boolean): Planned | undefined {
  const kinds = (slide.blocks ?? []) as FreeBlockKind[];
  const slots = slide.slots ?? {};
  const planning = context.planning;
  const rtl = context.direction === 'rtl';
  const Wpt = kit.slideWidth / 12700;
  const Hpt = kit.slideHeight / 12700;
  /** Sizes that are fixed in points are for a slide 13.33 inches wide; a narrower slide takes less. */
  const k = Wpt / 960;
  const deckScale = context.textScale ?? 1;
  // The floor of every size, on the half points a file is written in.
  const HARD = half(Math.max(4, 5 * k));
  const blocksOnSlide = kinds.filter((kind) => kind !== 'note').length;

  const at = (index: number, part: string) => slots[freeSlotName(index, part)];
  /** While planning, the number of paragraphs a text that was asked a length of will have, where the planner gives one (`planning.paragraphs`). */
  const countOf = (index: number, part: string): number | undefined => {
    const asked = planning?.lengths?.[freeSlotName(index, part)];
    const count = planning?.paragraphs?.[freeSlotName(index, part)];
    return asked !== undefined && asked > 0 && count !== undefined && Number.isInteger(count) && count >= 1 ? Math.min(count, asked) : undefined;
  };
  /** While planning a slide a person asked to have whole (`arrange.dense`), the paragraphs of a text code sets as it is, by their lengths (`planning.asIs`). */
  const asIsOf = (index: number, part: string): readonly number[] | undefined => {
    const each = slide.arrange?.dense === true && planning?.lengths?.[freeSlotName(index, part)] !== undefined ? planning.asIs?.[freeSlotName(index, part)] : undefined;
    return each && each.length > 0 ? each : undefined;
  };
  /** How long the paragraphs of a text part are: as written, or, while planning, as code sets them (`asIsOf`), as long as was asked or as texts of the kind usually are. */
  const lengthsOf = (index: number, part: string): readonly number[] => {
    if (!planning) return paragraphsOf(at(index, part)).map((paragraph) => paragraph.length);
    const own = asIsOf(index, part);
    if (own) return own;
    const asked = planning.lengths?.[freeSlotName(index, part)];
    return asked !== undefined && asked > 0 ? askedParagraphs(asked, countOf(index, part)) : (USUAL[kinds[index - 1] as FreeBlockKind]?.[part] ?? []);
  };
  const said = (index: number, part: string): number => lengthsOf(index, part).reduce((sum, chars) => sum + chars, 0);
  /**
   * The headings of a text part that stand over a paragraph (`headingsOver`): of a text as written,
   * by the slide's headings (`SlideSpec.headings`); while planning, of a text code sets as it is
   * (`planning.heads`).
   */
  const headsOf = (index: number, part: string): readonly HeadingOver[] => {
    if (context.asThen) return [];
    if (planning) return asIsOf(index, part) ? (planning.heads?.[freeSlotName(index, part)] ?? []) : [];
    const paragraphs = paragraphsOf(at(index, part));
    return headingsOver(paragraphs, headingPlaces(paragraphs, slide.headings));
  };
  /** The least size a part of a block is set at: what it has without a deck's wish for larger text, or 0. */
  const floorOf = (index: number | undefined, label: string): number => (index === undefined ? 0 : (context.atLeast?.get(`Block ${index} ${label}`) ?? 0));
  /** Whether an optional part is there: by its content, or, while planning, by what is offered. */
  const has = (index: number, kind: FreeBlockKind, part: string): boolean => {
    if (!planning) return present(at(index, part));
    const entry = FREE_PARTS[kind].find((candidate) => candidate.name === part);
    if (!entry || entry.required) return true;
    if (planning.withParts?.includes(freeSlotName(index, part))) return true;
    if (part === 'Icon') return planning.icons;
    if (part === 'Heading') return blocksOnSlide >= 2;
    // A caption tells like things apart: charts, tables, photos and graphics are offered one when the slide has several of the kind.
    if (part === 'Caption') return kinds.filter((other) => other === kind).length >= 2;
    return part === 'Source';
  };
  const hasTitle = planning ? true : present(slots.title);
  /** How far a last line may hang under its box: a quarter of its size, and no more than 3 pt on a slide 13.33 inches wide. */
  const hang: Hang = (size) => Math.min(HANG * size, MOST_HANG * (kit.slideWidth / 12700 / 960));

  // Natural sizes: the template's own, with the template's and the deck's factor on body text.
  const natural = (style: FreeTextStyle, scaled = true, deck = deckScale): number => {
    const size = style.set?.size ?? style.size;
    return half(size * (scaled ? deck * (style.size <= 12 ? kit.textScale : 1) : 1));
  };
  const T = kit.title?.size ?? 28;
  const bodyN = freeBodySize(kit, deckScale);
  const headN = Math.min(Math.max(natural(kit.styles.heading), bodyN), 0.8 * T);
  // The same two as the template alone asks for them, without the deck's factor. What a block's
  // width allows, and how much larger a heading is than its text, are worked out from these, so
  // that a deck which asks for larger text never gets smaller text for it.
  const body1 = freeBodySize(kit);
  const head1 = Math.min(Math.max(natural(kit.styles.heading, true, 1), body1), 0.8 * T);
  const captionN = Math.min(natural(kit.styles.caption), bodyN);
  const figureN = natural(kit.styles.figure, false);
  // A quotation is set clearly larger than running text, whatever the template's own style gives.
  const quoteN = Math.min(Math.max(natural(kit.styles.quote), half(1.6 * bodyN)), 0.9 * T);
  const noteStyle = kit.styles.note ?? kit.notePanel?.text ?? kit.styles.heading;
  // A note is the bottom line of the slide, not its loudest text: it is set no larger than nine
  // tenths of a heading, whatever the text was that the template set on its panel.
  const noteN = Math.min(kit.styles.note || kit.notePanel ? natural(noteStyle) : bodyN, 0.65 * T, Math.max(bodyN, halfDown(0.9 * headN)));

  // A note is a band across the slide, one to three lines high. `lower` is under 1 on a slide
  // whose notes are set smaller to share their room (layout.ts): the factor on the size of the
  // text, and on the least space kept around it.
  const noteInset = 0.019 * Wpt;
  const noteSize = (width: number, lower = 1) => {
    const size = Math.max(HARD, Math.min(noteN, halfDown((width - 2 * noteInset) / (WIDE * 30))));
    return lower < 1 ? Math.max(HARD, halfDown(size * lower)) : size;
  };
  /**
   * Whether a note that is written is set on one line: counted at half its size a character, the
   * width a note's heavy face has in PowerPoint, and not with the room to spare that budgets are
   * stated with. A band for one line keeps 0.6 of a line free above and below the text, so a
   * note that takes a second line after all still ends inside it.
   */
  const oneLine = (index: number, width: number, size: number) => !planning && said(index, 'Text') <= Math.floor((width - 2 * noteInset) / (0.5 * size));
  // A note that fits one line is one line: nothing is lost at a break that is not there.
  const noteLines = (index: number, width: number, size: number) => (said(index, 'Text') <= perLine(width - 2 * noteInset, size) || oneLine(index, width, size) ? 1 : clamp(Math.ceil(said(index, 'Text') / (perLine(width - 2 * noteInset, size) * FULL)), 1, 3));
  const noteAround = (size: number, lower: number) => 2 * Math.max(0.011 * Wpt * lower, 0.6 * LINE * size);
  const noteHeight = (index: number, width: number, lower = 1): number => {
    const size = noteSize(width * Wpt, lower);
    return (linesHigh(noteStyle, noteLines(index, width * Wpt, size), size) + noteAround(size, lower)) / Wpt;
  };

  // A slide without a title has its blocks start where the title would: never over what stays above the content.
  const untitled = kit.title ? Math.max(kit.title.box.y, kit.area.y) : kit.band.top;
  // A title longer than its line takes a second line, or a third, under the first: the blocks
  // start under those, half a gap lower, and not under a title that lies across them. Its
  // lines are counted at half its size a character, the width a title's heavy face has. A title
  // that is set from its foot upwards only gets more room than it needs: erring low is harmless.
  const titleChars = planning || !kit.title ? 0 : paragraphsOf(slots.title).reduce((sum, paragraph) => sum + paragraph.length, 0);
  const titleLines = freeTitleLines(kit, titleChars);
  const underTitle = kit.title && titleLines > 1 ? kit.title.box.y + (titleLines * LINE * kit.title.size + 0.5 * kit.gap * Wpt) / Hpt : 0;
  // Under a title of two lines the blocks keep the distance they have from a title of one: the
  // room starts lower by the lines the title has more. Otherwise the blocks would hang from its last line.
  const titleDrop = kit.title && titleLines > 1 ? ((titleLines - 1) * LINE * kit.title.size) / Hpt : 0;
  const top = Math.max((hasTitle || !kit.title ? kit.band.top : untitled) + titleDrop, underTitle);
  /**
   * The lines the heading of a card or a step takes on lines of so many characters: one when it
   * is no longer than a line, otherwise two, or three when its words do not go into two. Before
   * there is content a heading is one line.
   */
  const headingLines = (index: number, across: number): number => {
    if (planning || said(index, 'Heading') <= across) return 1;
    return clamp(wrapLines(paragraphsOf(at(index, 'Heading')).join(' '), across), 2, 3);
  };
  // A card is as high as its heading and text need, so that a row of cards is not a row of half-empty boxes.
  // `set` is for how far down the card what is written on it most likely reaches once PowerPoint
  // has set it, which is less: its lines are then counted as PowerPoint fills them, and the card's
  // padding is counted above them only, not under the text.
  const cardHeight = (index: number, kind: FreeBlockKind, widthU: number, set = false): number | undefined => {
    if (kind !== 'card' && kind !== 'step') return undefined;
    const w = widthU * Wpt;
    const p = kit.panel ? clamp(0.09 * w, 0.006 * Wpt, 0.03 * Wpt) : 0;
    const inner = w - 2 * p;
    const textSize = Math.max(HARD, Math.min(bodyN, halfDown(inner / (WIDE * 20))));
    const headSize = Math.max(HARD, Math.min(headN, Math.max(LINE * textSize, (head1 * textSize) / body1), halfDown(inner / (WIDE * 12))));
    const everyIcon = kind === 'card' && kinds.every((other, position) => other !== 'card' || has(position + 1, 'card', 'Icon'));
    const lead = kind === 'step' ? 1.3 * Math.min(figureN, 1.6 * headSize) : everyIcon ? clamp(0.175 * w, 0.018 * Wpt, 0.0375 * Wpt) + 0.75 * Math.max(p, 0.008 * Wpt) : 0;
    if (set && !planning) {
      const headLines = clamp(setLines(said(index, 'Heading'), inner, headSize), 1, 3);
      return (p + lead + linesHigh(kit.styles.heading, headLines, headSize) + 0.5 * headSize + setLines(said(index, 'Text'), inner, textSize) * lineOf(kit.styles.body, textSize)) / Wpt;
    }
    const headLines = headingLines(index, perLine(inner, headSize));
    // Under the text 0.6 of a line of the usual height is kept free.
    return (2 * p + lead + linesHigh(kit.styles.heading, headLines, headSize) + 0.5 * headSize + cardTextHigh(said(index, 'Text'), inner, textSize, kit.styles.body) + 0.6 * LINE * textSize) / Wpt;
  };
  /**
   * How high a text, a list or a quotation needs to be at a width where it stands over or under
   * blocks of other kinds: its lines at the size it has there, a heading or a source where it
   * has one, and one line more. What it has no use for goes to its neighbours.
   */
  const proseHeight = (index: number, kind: FreeBlockKind, widthU: number): number | undefined => {
    if (kind !== 'text' && kind !== 'list' && kind !== 'quote') return undefined;
    const w = widthU * Wpt;
    if (kind === 'quote') {
      const size = Math.max(HARD, Math.min(quoteN, halfDown(w / (WIDE * 18))));
      const lines = Math.max(1, Math.ceil(said(index, 'Quote') / (perLine(Math.min(w, 50 * WIDE * size), size) * FULL)));
      const source = has(index, 'quote', 'Source') ? LINE * Math.min(bodyN, 0.5 * size + 4) + 0.7 * size : 0;
      return (linesHigh(kit.styles.quote, lines, size, 1) + source) / Wpt;
    }
    const style = kind === 'text' ? kit.styles.body : kit.styles.list;
    const size = Math.max(HARD, Math.min(bodyN, halfDown(w / (WIDE * (kind === 'text' ? 28 : 24)))));
    const each = lengthsOf(index, kind === 'text' ? 'Text' : 'Items');
    const across = perLine(w - (kind === 'list' ? hangAt(size) : 0), size) * FULL;
    const lines = each.reduce((sum, chars) => sum + Math.max(1, Math.ceil(chars / across)), 0);
    const heading = has(index, kind, 'Heading') ? (LINE + 0.65) * Math.min(headN, LINE * size) : 0;
    return (heading + linesHigh(style, Math.max(1, lines), size, 1) + Math.max(0, each.length - 1) * gapAt(size)) / Wpt;
  };
  // The legend of a slide of charts has a strip twice its text size high, at the top of the room.
  const legendStrip = legend ? (2 * legend.size) / Hpt : 0;
  const denseTop = Math.max((hasTitle || !kit.title ? kit.dense.top : untitled) + titleDrop, underTitle);
  const layout = arrangeBlocks(kinds, {
    frame: { x: kit.area.x, w: kit.area.w, top: top + legendStrip, bottom: kit.band.bottom, denseTop: denseTop + legendStrip, denseBottom: kit.dense.bottom, gap: kit.gap, aspect: Wpt / Hpt, card: kit.panel ? clamp((kit.panel.drawn.h * Hpt) / (kit.panel.drawn.w * Wpt), 0.6, 1.6) : 0 },
    arrange: slide.arrange as FreeArrange | undefined,
    rtl,
    noteHeight,
    needed: cardHeight,
    likely: (index, kind, width) => cardHeight(index, kind, width, true),
    prose: proseHeight,
    crowded,
  });

  // ---- Colours on a ground the template's own look was not made for ---------------------
  /** What a chart or a table is drawn on: the fill of the panel under a tile, or the slide's background. */
  const groundOf = (onPanel: boolean): string => (onPanel && kit.panel?.fill ? kit.panel.fill : kit.background);
  /**
   * The colour a text is given so that it reads on `ground`, when the one it would have does
   * not: `stated`, or the theme's text colour where none is stated. It is the colour of the
   * slide's running text, which was chosen because it reads on this frame. `undefined` when the
   * text reads as it is.
   */
  const legible = (stated: string | undefined, ground: string): string | undefined => (contrast(stated ?? kit.ink, ground) >= 3 ? undefined : readable(kit.styles.body.color, ground, 3));
  /** A line that shows on `ground` and stays quiet: a third of the text's colour in the ground's own. */
  const faint = (text: string, ground: string): string => mix(text, ground, 1 / 3);

  // ---- Shapes and slots -----------------------------------------------------------------
  const shapes: MadeShape[] = [];
  const made: SlotDef[] = [];
  let nextId = kit.firstId;
  let smallest = Infinity;
  // A size that is there stays there when it is rounded: a chart or a picture cannot be drawn in no area at all.
  const extent = (size: number, whole: number): number => (size > 0 ? Math.max(0.0001, Math.round((size / whole) * 10000) / 10000) : 0);
  const toBox = (x: number, y: number, w: number, h: number): SlideBox => ({ x: Math.round((x / Wpt) * 10000) / 10000, y: Math.round((y / Hpt) * 10000) / 10000, w: extent(w, Wpt), h: extent(h, Hpt) });
  /** A box counted from the start edge of its block: for a deck read from the right it is mirrored inside the block. */
  const inBlock = (block: Rect, x: number, y: number, w: number, h: number): SlideBox => toBox(rtl ? block.x + block.w - (x - block.x) - w : x, y, w, h);

  interface Rect {
    x: number;
    y: number;
    w: number;
    h: number;
  }
  const rectOf = (block: LaidBlock): Rect => ({ x: block.box.x * Wpt, y: block.box.y * Hpt, w: block.box.w * Wpt, h: block.box.h * Hpt });
  const pad = (w: number, h: number) => Math.min(clamp(0.09 * w, 0.006 * Wpt, 0.03 * Wpt), 0.14 * h);

  const panel = (block: Rect, index: number): void => {
    if (!kit.panel) return;
    shapes.push({ id: String(nextId++), name: `Block ${index} panel`, kind: 'panel', from: { slide: kit.panel.slide, shapeId: kit.panel.shapeId }, box: toBox(block.x, block.y, block.w, block.h) });
  };
  interface TextOptions {
    slot?: string;
    required?: boolean;
    paragraphs?: boolean;
    most?: number;
    describe?: string;
    fixed?: string;
    wrap?: boolean;
    align?: 'start' | 'ctr';
    anchor?: 't' | 'ctr' | 'b';
    inset?: number;
    columns?: number;
    gap?: number;
    lines?: number;
    color?: string;
    /** The box holds no line of the text, even at the smallest size. Left out, a slot whose budget is not given is asked. */
    none?: boolean;
    /** The paragraphs of a text that was fitted: its slot holds at least so many. */
    fitted?: number;
    /** The paragraphs its slot takes, exactly: of a text a person asked to have whole, the most its writer is told, or as many as it has. */
    told?: number;
    /** Whether the text can take several lines. Left out: running text, lists, quotations, notes and labels can, the short parts cannot. */
    several?: boolean;
    /** How far its spacing gave way (`GIVE`): it is drawn so. Its budgets are counted at its own spacing all the same. */
    give?: number;
  }
  /** Makes a text shape as a copy of a style, and its slot when it holds content of the deck. */
  const text = (index: number, label: string, style: FreeTextStyle, size: number, box: SlideBox, width: number, height: number, options: TextOptions = {}): void => {
    const id = String(nextId++);
    smallest = Math.min(smallest, size);
    const shape: MadeShape = { id, name: `Block ${index} ${label}`, kind: 'text', from: { slide: style.slide, shapeId: style.shapeId }, box, size, align: options.align ?? 'start', anchor: options.anchor ?? 't', wrap: options.wrap ?? true, inset: options.inset ?? 0 };
    if (style.set?.bold !== undefined) shape.bold = style.set.bold;
    if (style.set?.italic !== undefined) shape.italic = style.set.italic;
    const color = options.color ?? style.set?.color;
    if (color) shape.color = color;
    if (style.set?.bullets) shape.bullets = 'add';
    else if (style.set?.bullets === false) shape.bullets = 'none';
    else if (style.bullet) shape.bullets = 'own';
    if (options.columns && options.columns > 1) shape.columns = options.columns;
    if (options.fixed !== undefined) shape.text = options.fixed;
    if (options.several ?? ['text', 'list', 'quote', 'note', 'label'].includes(label)) shape.lines = 'several';
    // Lines closed up are drawn so; lines that give nothing (a template at single) keep their source's spacing.
    if (options.give && lineOf(style, size, options.give) < lineOf(style, size) - 1e-9) shape.lineSpacing = drawnLeading(style, size, options.give);
    shapes.push(shape);
    if (!options.slot) return;
    const columns = options.columns && options.columns > 1 ? options.columns : 1;
    const inner = (width - 2 * (options.inset ?? 0) - (columns - 1) * kit.gap * Wpt) / columns;
    // A box that holds no line of its text holds none of it: with content that is what the deck's
    // check reports. Before there is content a writer is told one short line's worth: what the
    // width holds at the size the text is set at, which is then the smallest it is ever set at,
    // and never under a dozen characters. Nobody can write to one.
    const none = options.none ?? (options.most === undefined && !holdsLine(height * columns, size, lineOf(style, size), hang));
    const most = Math.max(1, none ? (planning ? Math.max(LEAST_SAID, Math.floor(perLine(inner, size) * FULL)) : 1) : (options.most ?? capacity(inner, height * columns, size, hang, options.paragraphs ? 2 : 1, lineOf(style, size))));
    const slot: FreeTextSlot = { name: options.slot, kind: 'text', shapeId: id, box, size, fontSize: size, scales: false, alignOnDirectionChange: 'keep', maxChars: most, styleOf: { slide: style.slide, shapeId: style.shapeId }, drawn: { inset: options.inset ?? 0, wrap: options.wrap ?? true } };
    if (shape.bullets) slot.drawn.hang = shape.bullets === 'none' ? 0 : hangAt(size);
    if (shape.lineSpacing) slot.drawn.lineSpacing = shape.lineSpacing;
    if (columns > 1) {
      slot.drawn.columns = columns;
      slot.drawn.gap = kit.gap * Wpt;
      slot.drawn.height = height;
    }
    if (options.paragraphs) {
      slot.maxParagraphs = options.told ?? Math.max(options.fitted ?? 0, clamp(Math.floor(((height * columns) / lineOf(style, size)) / (options.lines ?? 2)), 1, 12));
      slot.paragraphGap = gapAt(size, options.give);
    }
    if (!options.required) slot.empty = 'remove';
    const described = options.describe ?? target(most);
    if (described !== undefined) slot.description = described;
    made.push(slot);
  };

  /** With a legend for the slide: a chart came out bare in the room the legend leaves. */
  let bare = false;
  /** Card text longer than its budget had to be drawn smaller (see `freePlan`). */
  let squeezed = false;
  /** A long text a person asked to have whole is too long for four fifths of the template's size (see `freePlan`). */
  let cramped = false;
  /** A chart of the slide and its place: what its slot is worked out from. */
  interface ChartPlace {
    index: number;
    id: string;
    box: SlideBox;
    area: Rect;
    /** The colour it stands on. */
    ground: string;
    /** It stands on a panel, as a tile among other blocks. */
    tile: boolean;
  }
  /**
   * The slots of the charts of a slide. Each chart is drawn in as much detail as its place reads
   * at, and like charts are drawn alike: charts of one type, one format and one number of series
   * share their value axis and the way their numbers are written, so that a bar of one can be
   * read against a bar of the next. Whether numbers stand at the bars is decided once for the
   * whole slide.
   */
  const chartSlots = (places: ChartPlace[]): ChartSlotDef[] => {
    // No text of a chart is set under 7 pt on a slide 13.33 inches wide: nobody reads it on a
    // projected slide. A chart whose place carries no more is drawn bare.
    const least = 7 * k;
    const stated = kit.chart.text?.size ?? captionN;
    const own = half(stated * deckScale * (stated <= 12 ? kit.textScale : 1));
    const crowd = places.length > MOST_DRAWN;
    const round1 = (value: number) => Math.round(value * 10) / 10;
    const entries = places.map((place) => {
      const { area, index } = place;
      const value = planning ? undefined : drawable(at(index, 'Chart'));
      // A chart is drawn bare where nothing else can be read: among more than two dozen, in a
      // place too low or too narrow for a line of small text beside its bars, or where the text
      // its place carries would be under the least. A small chart keeps the names of its
      // categories, and its numbers or its value axis. On a tile among other blocks a chart is a
      // picture of its numbers more than a thing to read: it keeps its names only where the tile
      // leaves it 7.5% of the slide's width in height.
      const lowest = place.tile ? 0.075 * Wpt : 42.25 * k;
      let detail: NonNullable<ChartSlotDef['detail']> = crowd || area.w < 0.085 * Wpt || area.h < lowest ? 'spark' : area.w >= 0.27 * Wpt && area.h >= 0.165 * Wpt ? 'full' : 'compact';
      // Its text follows its frame: the template's own size and no more than the frame carries,
      // so that names and numbers leave the plot its room: 6% of its width and a tenth of its
      // height in a full chart, 4.5% and 7.5% in a small one, and never under what a small chart
      // carries there: up to 10 pt, by its width, and an eighth of its height, so that names
      // and numbers take no more than two fifths of a low chart and its bars have the rest. A
      // frame too low for that at the least size is set at the least where it carries it.
      const byHeight = Math.max(halfDown(area.h / 8), Math.min(halfDown(area.h / 6.5), Math.ceil(least * 2) / 2));
      const carries = Math.min(half(Math.min(0.042 * area.w, 10 * k)), byHeight);
      let size = Math.max(carries, detail === 'full' ? Math.min(own, halfDown(0.06 * area.w), halfDown(0.1 * area.h)) : Math.min(own, halfDown(0.045 * area.w), halfDown(0.075 * area.h)));
      // Beside horizontal bars every category is named: the text is as small as that takes.
      if (value?.type === 'bar' && detail !== 'spark') size = Math.min(size, halfDown((area.h - 2 * least) / (1.15 * value.categories.length)));
      // Under columns and lines the names are set up to a fifth smaller where that lets every one of
      // them stand side by side (`NAME_WIDE`, as the chart is drawn with), and are thinned out only beyond that.
      if (value && value.type !== 'bar' && value.type !== 'pie' && detail !== 'spark') {
        const longest = value.categories.reduce((most, category) => Math.max(most, category.length), 1);
        const named = (at: number) => (area.w - 4.5 * at) / value.categories.length >= (longest * NAME_WIDE + 0.5) * at;
        for (let at = halfDown(size); !named(size) && at >= Math.max(least, 0.8 * size) - 1e-9; at -= 0.5) {
          if (named(at)) size = at;
        }
      }
      size = Math.max(context.atLeast?.get(freeSlotName(index, 'Chart')) ?? 0, size);
      if (detail !== 'spark' && size < least - 1e-9) detail = 'spark';
      size = Math.max(HARD, size);
      const numbers = value ? knownValues(value) : [];
      return { place, value, detail, size, numbers, most: numbers.reduce((most, number) => Math.max(most, number), 0), below: numbers.some((number) => number < 0) };
    });
    type Entry = (typeof entries)[number];
    // The charts of a slide that are drawn in one detail set their names, axis and numbers at one
    // size: the smallest that any of them needs. A bar chart beside a column chart does not have
    // larger names than it.
    for (const detail of ['full', 'compact'] as const) {
      const alike = entries.filter((entry) => entry.detail === detail);
      const shared = alike.reduce((least, entry) => Math.min(least, entry.size), Infinity);
      for (const entry of alike) entry.size = Math.max(context.atLeast?.get(freeSlotName(entry.place.index, 'Chart')) ?? 0, shared);
    }

    // Charts that are alike, and whose sizes are near enough to be read on one scale, are one group.
    const groups = new Map<string, Entry[]>();
    for (const entry of entries) {
      const { value } = entry;
      if (!value || value.type === 'pie' || entry.below || !(entry.most > 0)) continue;
      const key = `${value.type}|${value.format === 'percent' ? 'percent' : 'number'}|${value.series.length}`;
      groups.set(key, [...(groups.get(key) ?? []), entry]);
    }
    const units: Entry[][] = [];
    const grouped = new Set<Entry>();
    for (const members of groups.values()) {
      const sizes = members.map((member) => member.most);
      if (members.length < 2 || Math.max(...sizes) > ALIKE * Math.min(...sizes)) continue;
      units.push(members);
      for (const member of members) grouped.add(member);
    }
    for (const entry of entries) if (!grouped.has(entry)) units.push([entry]);

    const oneLegend = legend !== undefined;
    if (oneLegend && entries.some((entry) => entry.detail === 'spark')) bare = true;
    const drawnOf = new Map<Entry, ChartDrawn>();
    /** What is known of a group before the slide decides about numbers: how its numbers are written, and how wide the widest is. */
    const ways = units.map((unit) => {
      const first = unit[0] as Entry;
      const value = first.value;
      const look = numberLook(unit.flatMap((entry) => entry.numbers), value?.format === 'percent');
      const widest = unit.reduce((most, entry) => entry.numbers.reduce((wide, number) => Math.max(wide, shownAs(number, look).length), most), 1);
      return { unit, first, value, shared: unit.length > 1, look, widest, horizontal: value?.type === 'bar', read: unit.filter((entry) => entry.detail !== 'spark') };
    });
    type Way = (typeof ways)[number];
    for (const { unit, shared, look } of ways) {
      for (const entry of unit) drawnOf.set(entry, { width: round1(entry.place.area.w), height: round1(entry.place.area.h), ground: entry.place.ground, ...(oneLegend ? { legend: false } : {}), ...(shared ? { look } : {}) });
    }
    /**
     * The plot of a chart: its frame without the room its names, its axis, its numbers and its
     * legend take (`plotMargins`, which the chart is then drawn with). `along` is the length
     * the categories share, `across` the length the values are drawn on.
     */
    const plotOf = (way: Way, entry: Entry, axis: boolean, numbers: boolean): { along: number; across: number } => {
      const value = entry.value as ChartValue;
      const margins = plotMargins({ horizontal: way.horizontal, size: entry.size, names: value.categories.reduce((most, category) => Math.max(most, category.length), 1), axis: axis ? way.widest : undefined, numbers: numbers ? way.widest : undefined, legend: value.series.length > 1 && !oneLegend });
      const w = Math.max(1, entry.place.area.w - margins.left - margins.right);
      const h = Math.max(1, entry.place.area.h - margins.top - margins.bottom);
      return way.horizontal ? { along: h, across: w } : { along: w, across: h };
    };
    /** Whether the numbers have room at the bars of a chart: each no wider than a column, or no higher than a horizontal bar is thick. */
    const holdsNumbers = (way: Way, entry: Entry, axis: boolean): boolean => {
      const value = entry.value as ChartValue;
      const plot = plotOf(way, entry, axis, true);
      const each = plot.along / (value.categories.length * (value.series.length + (value.series.length > 1 ? 0.8 : 0.6)));
      // A plot that numbers would leave too little of keeps its axis instead.
      const room = way.horizontal ? entry.place.area.w : entry.place.area.h;
      return plot.across >= 0.3 * room && (way.horizontal ? each >= 0.95 * entry.size : each >= 1.05 * way.widest * WIDE * entry.size);
    };
    // Numbers at the bars in place of a value axis is decided once for the slide, so that one
    // slide does not show small charts with numbers and no axis beside small charts with an axis
    // and no numbers: every small bar and column chart has its numbers at the bars where each of
    // them has room for them and no small line stands among them (a line carries no numbers), and
    // none has them otherwise. A small chart with a value under nothing keeps its axis, which says
    // where nothing is. A full chart always keeps its axis, so its numbers are an addition: like
    // charts that are drawn full have them where every one of them has room for them, whatever
    // an unlike chart beside them can hold.
    const read = ways.filter((way) => way.value && way.value.type !== 'pie' && way.read.length > 0);
    const small = read.filter((way) => way.read.some((entry) => entry.detail === 'compact'));
    const smallNumbered = small.length > 0 && small.every((way) => way.value?.type !== 'line' && !way.first.below && way.read.every((entry) => entry.detail !== 'compact' || holdsNumbers(way, entry, false)));
    const fullNumbered = new Map(read.map((way) => [way, way.value?.type !== 'line' && way.read.every((entry) => entry.detail !== 'full' || holdsNumbers(way, entry, true))]));
    const numbersOf = (way: Way, entry: Entry): boolean => (entry.detail === 'full' ? fullNumbered.get(way) === true : entry.detail === 'compact' && smallNumbered);
    const axisOf = (entry: Entry): boolean => entry.detail === 'full' || (entry.detail === 'compact' && !smallNumbered);
    for (const way of ways) {
      const { unit, first, value, shared, look } = way;
      if (!value) continue;
      if (value.type === 'pie') {
        for (const entry of unit) (drawnOf.get(entry) as ChartDrawn).numbers = entry.detail === 'full';
        continue;
      }
      for (const entry of unit) {
        const drawn = drawnOf.get(entry) as ChartDrawn;
        drawn.numbers = numbersOf(way, entry);
        if (entry.detail === 'compact' && smallNumbered) drawn.valueAxis = false;
      }
      // Where the axis starts and ends (`roundSpan`). Bars start at nothing; a line starts there
      // too, unless it stays far above nothing (its lowest value at 60% of its highest or more),
      // and then at a round value under its lowest. An axis that is shown has three to five round
      // labels, written with the decimals of its step, and ends at the half step over the largest
      // value. Where no axis is shown a bar chart ends at its largest value, so that the longest
      // bar uses the whole of its plot; the number at that bar has its room over the plot.
      const all = unit.flatMap((entry) => entry.numbers);
      const low = all.reduce((least, number) => Math.min(least, number), Infinity);
      const high = all.reduce((most, number) => Math.max(most, number), -Infinity);
      const withAxis = way.read.filter(axisOf);
      if (withAxis.length === 0) {
        if (!first.below && first.most > 0 && (value.type !== 'line' || shared)) for (const entry of unit) (drawnOf.get(entry) as ChartDrawn).axis = { max: unit.reduce((most, member) => Math.max(most, member.most), 0) };
        continue;
      }
      const stepsOf = (entry: Entry): number => {
        const plot = plotOf(way, entry, true, numbersOf(way, entry));
        const need = way.horizontal ? (way.widest * WIDE + 1.2) * entry.size : 1.8 * entry.size;
        return clamp(Math.floor(plot.across / need), 2, 4);
      };
      const from = value.type === 'line' && low > 0 && low >= 0.6 * high ? low : undefined;
      // A chart with values under nothing and numbers at its bars keeps room under its lowest bar
      // for the number at its end, so that it stands clear of the names of the categories: as
      // much of its plot as the number takes, and not a whole step more.
      const under = first.below
        ? way.read.filter((entry) => numbersOf(way, entry)).reduce((most, entry) => Math.max(most, (way.horizontal ? (way.widest * 0.55 + 0.6) * entry.size : 1.4 * entry.size) / plotOf(way, entry, axisOf(entry), true).across), 0)
        : 0;
      const span = roundSpan(low, high, { steps: withAxis.reduce((fewest, entry) => Math.min(fewest, stepsOf(entry)), Infinity), grain: look.divisor / 10 ** look.decimals, unit: look.divisor, from, under });
      for (const entry of unit) {
        const drawn = drawnOf.get(entry) as ChartDrawn;
        drawn.axis = { max: span.max, ...(axisOf(entry) ? { step: span.step } : {}), ...(span.min !== 0 ? { min: span.min } : {}) };
      }
    }

    return entries.map((entry) => {
      const { place, detail, size } = entry;
      if (detail !== 'spark') smallest = Math.min(smallest, size);
      // The template's chart was made for the ground its own layouts have. On another, such
      // as a dark frame, its text and its lines are set in colours that read there.
      const textColor = legible(kit.chart.text?.color, place.ground);
      const lineColor = contrast(kit.chart.lineColor ?? 'D9D9D9', place.ground) >= 1.1 ? kit.chart.lineColor : faint(textColor ?? kit.chart.text?.color ?? kit.ink, place.ground);
      const slot: ChartSlotDef = { name: freeSlotName(place.index, 'Chart'), kind: 'chart', shapeId: place.id, box: place.box, text: { ...(kit.chart.text ?? {}), size, ...(textColor ? { color: textColor } : {}) } };
      if (kit.chart.colors) slot.colors = kit.chart.colors;
      if (lineColor) slot.lineColor = lineColor;
      if (detail !== 'full') slot.detail = detail;
      slot.drawn = drawnOf.get(entry) as ChartDrawn;
      // What reads well at this size. The limits are soft: a chart over them is drawn, and the deck's check says so.
      slot.maxCategories = detail === 'full' ? 8 : detail === 'compact' ? clamp(Math.floor(place.area.w / Wpt / 0.034), 3, 8) : 12;
      slot.maxSeries = detail === 'full' ? 3 : detail === 'compact' ? 2 : 1;
      slot.description = `${detail === 'full' ? 'A chart' : 'A small chart'}. When the brief gives no numbers for it, leave it empty`;
      return slot;
    });
  };

  // Like parts of like blocks share one size: the smallest that any of them needs.
  const blocksOf = (kind: FreeBlockKind) => layout.blocks.filter((block) => block.kind === kind);
  const tiles = layout.tiles;
  const single = blocksOnSlide === 1;

  /** The size of some like texts: no larger than `most`, shrunk to fit by at most a fifth, grown when all are short. */
  /** The size texts are set at, between `most` and four fifths of `floorOf` (by default `most`), or larger where they are short. */
  const sized = (texts: Fitted[], most: number, grow = 1, floorOf = most, give: readonly number[] = NO_GIVE): number => {
    const least = Math.max(HARD, 0.8 * Math.min(most, floorOf));
    if (texts.every((entry) => entry.chars === 0)) return Math.max(HARD, halfDown(most));
    const short = texts.every((entry) => entry.chars <= 0.6 * charsHeld(entry, most, hang));
    return Math.max(HARD, fitSize(texts, short ? most * grow : most, least, hang, give));
  };

  // ---- running text and lists
  // They are worked out before cards and steps: short text on a card is set larger, but never
  // larger than the running text or the list of its slide (`proseSize`).
  let proseSize = Infinity;
  for (const kind of ['text', 'list'] as const) {
    const blocks = blocksOf(kind);
    if (blocks.length === 0) continue;
    const part = kind === 'text' ? 'Text' : 'Items';
    const bodyStyle = kind === 'text' ? kit.styles.body : kit.styles.list;
    const least = kind === 'text' ? 28 : 24;
    // How wide a character of running text is counted: as the template says PowerPoint sets it
    // (`characterWidth`), or by the estimate that leaves room for any face, and for a text set
    // dense by the narrower one of PowerPoint's own rendering of two templates.
    const bodyWide = context.asThen ? Math.max(WIDE, kit.characterWidth ?? WIDE) : (kit.characterWidth ?? WIDE);
    const denseWide = context.asThen ? Math.max(DENSE_WIDE, kit.characterWidth ?? DENSE_WIDE) : (kit.characterWidth ?? DENSE_WIDE);
    const rooms = blocks.map((block) => {
      const rect = rectOf(block);
      const onPanel = tiles && !!kit.panel;
      const p = onPanel ? pad(rect.w, rect.h) : 0;
      return { block, rect, onPanel, p, x: rect.x + p, y: rect.y + p, w: rect.w - 2 * p, h: rect.h - 2 * p, heading: has(block.index, kind, 'Heading') };
    });
    const alone = single && rooms.length === 1 ? (rooms[0] as (typeof rooms)[number]) : undefined;
    const chars = alone ? said(alone.block.index, part) : 0;
    let most = Math.max(HARD, Math.min(bodyN, ...rooms.map((room) => halfDown(room.w / (bodyWide * least)))));
    // Texts beside each other are set no smaller than four fifths of the top size as 0.6.6 counted
    // it. Counted with a template's narrower width the top is higher in a deck that asks for larger
    // text, and four fifths of it held texts that need smaller over their budget.
    const mostThen = Math.max(HARD, Math.min(bodyN, ...rooms.map((room) => halfDown(room.w / (Math.max(WIDE, bodyWide) * least)))));
    let columns = 1;
    /** A text alone keeps a line a person reads in one go: so many characters to a line. */
    let measure: number | undefined;
    let down = 0;
    if (alone && chars > 0) {
      if (chars <= 90 && kind === 'text') {
        // A sentence alone on a slide is set large, a little down the page.
        most = half(Math.min(1.6 * bodyN, 0.8 * T));
        measure = 40;
      } else if (chars <= 600) {
        most = half(Math.min(1.25 * bodyN, 0.65 * T));
        measure = 66;
      } else {
        most = bodyN;
        const lines = Math.ceil(chars / (perLine(alone.w, most) * FULL));
        if (lines > 9 && kind === 'text') columns = 2;
        else measure = 90;
      }
    } else if (alone) {
      measure = 66;
    }
    const headed = rooms.find((room) => room.heading);
    const headSize = Math.max(floorOf(headed?.block.index, 'heading'), HARD, Math.min(headN, Math.max(LINE * most, (head1 * most) / body1), ...rooms.map((room) => halfDown(room.w / (WIDE * 12)))));
    const headHeight = lineOf(kit.styles.heading, headSize) + 0.15 * headSize;
    // Under a heading, text is never set as large as the heading.
    const under = rooms.some((room) => room.heading) ? Math.max(HARD, halfDown(headSize / 1.05)) : Infinity;
    // The width of that line follows the size the text really gets: under a heading it is set
    // smaller than it would be alone, and a line made for the larger size would run on too long.
    const narrow = alone && measure !== undefined ? Math.min(alone.w, measure * WIDE * Math.min(most, Math.max(under, Math.min(most, bodyN)))) : undefined;
    /** The room of each text in so many columns: text in columns takes the whole width of its room. */
    const bodiesIn = (count: number) =>
      rooms.map((room) => {
        const width = count > 1 ? room.w : (narrow ?? room.w);
        const topOf = room.y + (room.heading ? headHeight + 0.5 * headSize : 0);
        const perColumn = (width - (count - 1) * kit.gap * Wpt) / count;
        // The text is drawn in a box at least a line of the smallest size high. What it has room for is counted from the room itself.
        return { room, width, perColumn, topOf, height: Math.max(LINE * HARD, room.y + room.h - topOf), left: room.y + room.h - topOf };
      });
    // The copy of a list has bullets, and its lines hang behind them: a line of it is shorter than its box is wide.
    const bulleted = bodyStyle.set?.bullets === true || (bodyStyle.set?.bullets !== false && bodyStyle.bullet);
    // A long text a person asked to have whole (`arrange.dense`) that its slide does not hold at
    // four fifths of the template's size is set out again in the larger room (`freePlan`). There
    // it is counted as PowerPoint sets it (`DENSE_WIDE`) and set as large as it fits, down to the
    // least size that reads (`readableSize`), in two columns, or three where two would make its
    // lines longer than a person reads in one go (`DENSE_LINE`). Under that size it is not set:
    // what the slide then holds is the budget of its text. Whether such a text fits, and what it
    // holds, is said by one judge (`judged`, `heldIn`), wherever it is asked.
    let wide = bodyWide;
    /** A text alone that a person asked to have whole (`arrange.dense`), long or short: its lines are counted paragraph by paragraph, and its slot states the box it is drawn in. */
    const dense = alone !== undefined && slide.arrange?.dense === true;
    // A short text of many paragraphs is one too: a list of fifteen short points that one column
    // did not hold at four fifths of the template's size was told six points in one column, and cut
    // by about a sixth. In two columns it is set whole.
    const whole = dense && kind === 'text';
    // Before there is content, such a text is set out for the most its writer may write: the top
    // of the range a writer is told (`DENSE_TOP`), in up to the most paragraphs it is told
    // (`DENSE_PARAGRAPHS`), or, where the planner gives the number of paragraphs (a document's),
    // the text in that many. Its paragraphs are then of any length (`Fitted.split`): a text set
    // out for paragraphs all as long as the next overflowed its box when its writer wrote the
    // same length in paragraphs each just over a line. A text code sets as it is (`asIs`, a
    // document's) is the text itself.
    const asIs = whole && planning ? asIsOf(alone.block.index, part) : undefined;
    const given = whole && planning && !asIs ? countOf(alone.block.index, part) : undefined;
    const told = whole && crowded && planning ? (asIs ?? (given !== undefined ? lengthsOf(alone.block.index, part) : askedParagraphs(tensDown(DENSE_TOP * chars), DENSE_PARAGRAPHS))) : undefined;
    const split = whole && planning !== undefined && asIs === undefined && (given !== undefined || told !== undefined);
    // Text that is written, and a person's text code sets as it is, gives way in its spacing
    // before its size (`GIVE`), but only to set larger a text that fits at the template's own
    // spacing: what is cut, reported or set whole is decided at that spacing, as 0.6.8 decided it
    // (the one judge of 0.6.3 is never loosened). What a writer is told before it writes is
    // counted at the template's own spacing too: a text inside it is set at the size planned for
    // it, or larger.
    const give = (planning !== undefined && asIs === undefined) || context.ownSpacing ? NO_GIVE : GIVE;
    /**
     * The texts to fit in those rooms: as written or planned, or, for a text alone, with the
     * paragraphs given (`lengths`). `wide` is how wide a character is counted. A text as written
     * or set as it is counts its headings as they are set (`Fitted.runIn`): run in at a size where
     * the renderer sets them so (`runsIn`), in a line a character shorter than it counts, so that
     * a heading counted run in is set run in. Over a paragraph that draws a bullet none is.
     */
    const fitOf = (list: ReturnType<typeof bodiesIn>, count: number, wide = bodyWide, lengths?: readonly number[]): Fitted[] =>
      list.map((body): Fitted => {
        const each = lengths ?? lengthsOf(body.room.block.index, part);
        const tall = whole ? { height: body.height, columns: count, ...(split ? { split: true } : {}) } : { height: body.height * count };
        // Lengths given are a writer's (no headings known), or, while planning, the text code sets as it is (`asIs`), whose headings are known.
        const heads = (lengths && lengths !== asIs) || bulleted ? [] : headsOf(body.room.block.index, part).filter((head) => head.at + 1 < each.length);
        const runIn = heads.length === 0 ? undefined : (size: number) => new Set(heads.filter((head) => runsIn(head.heading, head.next, context.direction, headingRoom(body.perColumn, size, kit.characterWidth) - 1)).map((head) => head.at));
        return { chars: each.reduce((sum, length) => sum + length, 0), paragraphs: each.length, width: body.perColumn, ...tall, style: bodyStyle, each, bullets: bulleted, wide, ...(runIn ? { runIn } : {}) };
      });
    let bodies = bodiesIn(columns);
    let toFit = fitOf(bodies, columns);
    // A text alone is set down to four fifths of the template's size here, and a text a person asked
    // to have whole no smaller than the least size the template states (`leastTextSize`), which is
    // never over the template's own size: what needs smaller goes to the larger room.
    const stated = whole && kit.leastTextSize !== undefined ? Math.min(kit.leastTextSize, halfDown(Math.min(most, bodyN))) : 0;
    /** The least size such a text is set at in the larger room: the template's own, never over its size, or the size that reads. */
    const denseFloor = kit.leastTextSize !== undefined ? Math.min(kit.leastTextSize, halfDown(Math.min(most, bodyN))) : leastSizeOf(kit);
    /** The least size a text alone is set at here, in the first room. */
    const firstFloor = Math.max(HARD, 0.8 * Math.min(most, bodyN), stated);
    const firstAt = (list: Fitted[], levels: readonly number[] = NO_GIVE): number => (alone && chars > 0 ? Math.max(HARD, fitSize(list, most, firstFloor, hang, levels)) : sized(list, most, blocksOnSlide <= 2 ? 1.25 : 1, mostThen, levels));
    // Which room a text a person asked to have whole takes is chosen at the template's spacing:
    // chosen with the spacing given way, texts that two columns set large stayed in one, smaller.
    const atOwn = firstAt(toFit);
    let fitted = atOwn;
    /** Where spacing that gives way lets the text be set: the size, in so many columns. Settled below. */
    let lift: { size: number; columns: number } | undefined = give === NO_GIVE ? undefined : { size: firstAt(toFit, give), columns };
    if (told) toFit = fitOf(bodies, columns, undefined, told);
    /**
     * Whether a text a person asked to have whole needed the larger room as 0.6.6 counted it: its
     * headings each on a line of their own, a character at `WIDE`. Its headings set run in, or a
     * template's narrower width (`characterWidth`), let it fit here more often; kept out of the
     * larger room, where it is set in columns, it came out smaller than 0.6.6 set it there (a text
     * of nine short parts under headings, 22 pt in two columns, at 14.5 in one). So the room is
     * chosen as then, and the size in it is found with the count of now, which is never larger. A
     * template whose width is wider than the estimate is counted at its own, as its text is.
     */
    const neededMore = (): boolean => {
      if (!whole || !((kit.characterWidth !== undefined && kit.characterWidth < WIDE) || toFit.some((text) => text.runIn) || fitOf(bodies, columns).some((text) => text.runIn))) return false;
      const asThen = (list: Fitted[]): Fitted[] => list.map(({ runIn: _runIn, ...text }): Fitted => ({ ...text, wide: Math.max(WIDE, kit.characterWidth ?? WIDE) }));
      const first = asThen(fitOf(bodies, columns));
      const at = Math.max(HARD, fitSize(first, most, Math.max(HARD, 0.8 * Math.min(most, bodyN), stated), hang));
      return !(told ? asThen(fitOf(bodies, columns, undefined, told)) : first).every((text) => fitsAt(text, at, hang));
    };
    const fitsFirst = whole && toFit.every((text) => fitsAt(text, atOwn, hang));
    if (whole && (!fitsFirst || neededMore())) {
      if (!crowded) cramped = true;
      else {
        const floor = denseFloor;
        const columnsAt = (size: number) => clamp(Math.ceil(alone.w / (DENSE_LINE * SET_WIDE * size)), 2, 3);
        /** The columns it may take at a size: as long as its lines are (`columnsAt`), or three where two do not hold it and lines in three are long enough (`DENSE_SHORT`). */
        const columnsFor = (size: number): number[] => {
          const usual = columnsAt(size);
          return usual === 2 && (alone.w - 2 * kit.gap * Wpt) / 3 / (SET_WIDE * size) >= DENSE_SHORT ? [2, 3] : [usual];
        };
        /** The largest size, and its columns (no more than `fewest`), at which the text fits at one of the spacings `levels`; where it fits at none, the least size in the columns its lines take there, as before. */
        const search = (levels: readonly number[], fewest = Number.POSITIVE_INFINITY): { chosen: { size: number; columns: number }; found: boolean } => {
          for (let size = halfDown(most); size >= floor - 1e-9; size -= 0.5) {
            for (const count of columnsFor(size)) {
              if (count > fewest) continue;
              if (spacingAt(fitOf(bodiesIn(count), count, denseWide, told), size, hang, levels) !== undefined) return { chosen: { size, columns: count }, found: true };
            }
          }
          return { chosen: { size: floor, columns: columnsAt(floor) }, found: false };
        };
        const { chosen, found } = search(NO_GIVE);
        // A text that came here only by the count of then (`neededMore`), and that the larger room
        // holds at no size, stays as it is set before the columns: there it fits, and in columns it
        // would be cut. The least size here can be over four fifths of a small template's size.
        if (found || !fitsFirst || kit.characterWidth === undefined) {
          wide = denseWide;
          columns = chosen.columns;
          bodies = bodiesIn(columns);
          toFit = fitOf(bodies, columns, wide, told);
          fitted = chosen.size;
          // Spacing that gives way may set it larger, in as many columns or fewer, where it fits here at all.
          lift = lift && found ? search(give, chosen.columns).chosen : undefined;
        }
      }
    }
    /** The size the texts are set at where they could be set at `at`: no smaller than the floor of the slide, no larger than a heading over them. */
    const settled = (at: number): number => Math.max(floorOf(blocks[0]?.index, kind === 'text' ? 'text' : 'list'), Math.min(at, Math.max(under, Math.min(most, bodyN))));
    let size = settled(fitted);
    // The spacing: the template's own, unless the texts fit at it at that size (they are set whole,
    // as 0.6.8 set them) and spacing that gives way sets them larger. Then they are set at that
    // size, in its columns, at the loosest spacing that fits.
    let spacing = 0;
    if (lift && toFit.every((text) => fitsAt(text, size, hang))) {
      const at = settled(lift.size);
      const list = lift.columns === columns ? toFit : fitOf(bodiesIn(lift.columns), lift.columns, wide, told);
      const level = at > size ? spacingAt(list, at, hang, give) : undefined;
      if (level !== undefined) {
        if (lift.columns !== columns) {
          columns = lift.columns;
          bodies = bodiesIn(columns);
        }
        toFit = list;
        size = at;
        spacing = level;
      }
    }
    if (spacing > 0) toFit = toFit.map((text) => giving(text, spacing));
    /** A text counted at the template's own spacing: what its budgets say. */
    const own = (text: Fitted): Fitted => (text.give ? { ...text, give: 0 } : text);
    proseSize = Math.min(proseSize, size);
    bodies.forEach((body, position) => {
      const { room } = body;
      const fit = toFit[position] as Fitted;
      const each = fit.each ?? [];
      if (room.onPanel) panel(room.rect, room.block.index);
      if (alone && chars > 0) {
        // Short text does not start at the very top of an empty slide. A text a person asked to
        // have whole (`arrange.dense`) is counted paragraph by paragraph, each ending on a line of
        // its own (`paragraphsHigh`): counted by its length, a list of twelve short items came to
        // five lines, and was set so far down that its last two items ran under its box.
        // A short text not set dense is counted no narrower than the estimate here: counted at a
        // template's own width, a text of many short paragraphs came to fewer lines than its
        // paragraphs take, and was set so far down that it ran under its box.
        const needed = dense ? paragraphsHigh(fit, size) : linesHigh(bodyStyle, Math.ceil(chars / (perLine(body.perColumn, size, Math.max(WIDE, bodyWide)) * FULL)), size) + Math.max(0, each.length - 1) * 0.5 * size;
        // A text whose spacing gave way fills its room: it is never short, and is not set down. Nor
        // is any text set further down than its paragraphs, each in whole lines, leave free: a text
        // of many short paragraphs, counted by its length, was set down until it ran under its box.
        const lower = columns === 1 && spacing === 0 && needed < 0.6 * body.height ? 0.3 * (body.height - needed) : 0;
        // Its last line may hang under the box as far as the check lets it (`hang`): a text it does not report stays where 0.6.8 set it.
        down = planning || context.ownSpacing ? lower : Math.max(0, Math.min(lower, body.height + hang(size) - paragraphsHigh(fit, size)));
      }
      if (room.heading) text(room.block.index, 'heading', kit.styles.heading, headSize, inBlock(room.rect, room.x, room.y + down, body.width, headHeight), body.width, headHeight, { slot: freeSlotName(room.block.index, 'Heading'), most: perLine(body.width, headSize), describe: 'A heading of a few words. Give one to every text of the slide, or to none' });
      // Text in two columns fills the first before it starts the second: its box is only as
      // high as half the text needs, so that the two columns come out alike. The box does not
      // grow or shrink with its text, so one that is a little low lets the second column run on
      // in the room the block has anyway.
      // A text a person asked to have whole (`arrange.dense`) in columns has a box as high as its
      // columns are where they come out alike (`judged`) and a line more, the line its columns
      // keep free (`columnRoom`). Columns that come out unlike, the last the shorter, are as
      // PowerPoint fills them.
      // The line as it is drawn, and as the template has it: budgets are counted at that.
      const line = lineOf(bodyStyle, size, spacing);
      const ownLine = lineOf(bodyStyle, size);
      const set = whole && columns > 1 ? judged(fit, size, hang).high + line : undefined;
      const halfLines = Math.ceil(Math.ceil((fit.chars * Math.max(0.47, kit.characterWidth ?? 0) * size) / body.perColumn) / columns);
      // Text not set dense in columns keeps the box its lines take at the template's spacing: counted
      // closer, the box of a Hebrew text was a few points lower than PowerPoint filled its first column.
      const halved = set ?? linesHigh(bodyStyle, halfLines, size) + Math.max(0, each.length - 1) * 0.5 * size;
      const boxHeight = columns > 1 ? Math.min(body.height, Math.max(line, halved)) : body.height - down;
      // What the slot states. Before there is content, what the room holds at this size: a text
      // that was set a little down the slide because the usual text is short has the whole room.
      // A text a person asked to have whole has the box it is drawn in, as it will be drawn.
      const high = planning && !dense ? body.height : body.height - down;
      const none = !holdsLine((body.left - (planning && !dense ? 0 : down)) * columns, size, line, hang);
      const perSlot = kind === 'list' ? 1.5 : 3;
      const paragraphsHeld = clamp(Math.floor((high * columns) / ownLine / perSlot), 1, 12);
      /** The text in its room as high as `height`: of a text a person asked to have whole, each column. */
      const inRoom = (text: Fitted, height: number): Fitted => (text.columns !== undefined ? { ...text, height } : { ...text, height: height * columns });
      const here = inRoom(fit, high);
      const usual = kind === 'list' ? listCapacity(body.perColumn, high * columns, size, planning ? Math.min(Math.max(1, each.length), paragraphsHeld) : Math.max(1, each.length), bodyStyle, hang, bodyWide) : capacity(body.perColumn, high * columns, size, hang, 2, ownLine, wide);
      let most = usual;
      let fitted = 0;
      /** Of a text a person asked to have whole: the paragraphs its slot takes, exactly. */
      let told: number | undefined;
      if (whole && planning && (split || (asIs !== undefined && !fitsAt(here, size, hang)))) {
        // A text a person asked to have whole, before there is content, that a writer writes: its
        // budget is what it holds at its size in round tens, in any paragraphs of any length up to
        // the most it is told (`heldIn`, `toldParagraphs`), which its slot takes. A writer inside
        // both is not sent back. A document's text that code would set as it is (`asIs`) but that
        // the slide does not hold even at the least size is written so too: code does not set it
        // (`filledText`), and its writer is told what the slide holds.
        told = fitsAt(here, size, hang) ? fit.paragraphs : toldParagraphs(own(here), size, hang, fit.paragraphs);
        most = tensDown(heldIn(own(here), size, hang, told));
        fitted = told;
      } else if (whole && asIs) {
        // A document's text that code sets as it is, and that fits: the slide holds it.
        most = Math.max(usual, fit.chars, planning?.lengths?.[freeSlotName(room.block.index, part)] ?? 0);
        told = fitted = each.length;
      } else if (whole && !planning && fit.chars > 0) {
        // Once written, the text itself is judged. A text that does not fit is always told so. It
        // is told the characters its paragraphs hold however they fall (`heldIn`), in as many
        // paragraphs as it has: as many hold no more than it has, or it would fit. A text of more
        // paragraphs than leave half its room for its lines (`roomyParagraphs`) is told that it
        // has too many, and what so many hold. It is never told a writer's twelve: a writer told
        // 32 was sent back to join them into twelve.
        const count = Math.max(1, each.length);
        if (judged(here, size, hang).fits) {
          most = Math.max(usual, fit.chars);
          told = fitted = each.length;
        } else {
          told = fitted = Math.min(count, roomyParagraphs(own(here), size));
          most = Math.min(usual, heldIn(own(here), size, hang, told), told === count ? fit.chars - 1 : Infinity);
        }
      } else if (whole && fit.chars > 0) {
        // Before there is content, a text of an asked length that its slide holds at four fifths of
        // the template's size or more: it is set as without the flag, smaller only where what is
        // written needs it.
        if (fitsAt(here, size, hang)) most = Math.max(usual, fit.chars);
        // Where the template's least size (`leastTextSize`) is over the size that reads, the larger
        // room sets what is written no smaller than this room does, and holds no more of it at its
        // least size than this room holds at its own: a writer told what the box holds in a few
        // long paragraphs wrote it in many and was cut. Its writer is told what this box holds at
        // the least size it sets the text at, in the paragraphs it is told (`heldIn`): what it
        // writes inside both is set here. The most paragraphs, two or more, that hold the asked
        // length, or else, as in the larger room, what the slide holds.
        if (denseFloor > readableSize(kit.slideWidth) + 1e-9) {
          const low = Math.min(size, firstFloor);
          let count = paragraphsHeld;
          while (count > 2 && tensDown(heldIn(own(here), low, hang, count)) < fit.chars) count--;
          if (tensDown(heldIn(own(here), low, hang, count)) < fit.chars) count = toldParagraphs(own(here), low, hang, fit.paragraphs);
          told = fitted = count;
          most = Math.min(most, tensDown(heldIn(own(here), low, hang, count)));
        }
      } else if (!whole && (!planning || countOf(room.block.index, part) !== undefined) && fit.chars > 0) {
        // Before there is content, a text whose number of paragraphs the planner gave is counted
        // by the lines its paragraphs take, as written text is.
        if (fitsAt(here, size, hang)) {
          // Text that was fitted is within what its slot holds, however its paragraphs share the room.
          most = Math.max(usual, fit.chars);
          fitted = each.length;
        } else {
          // Text that does not fit is over its budget: by its length, or by the lines its paragraphs take one by one.
          const needed = paragraphsHigh(own(here), size);
          const room = here.height + hang(size);
          most = Math.min(usual, charsHeld(own(here), size, hang), needed > room ? Math.floor((fit.chars * room) / needed) : Infinity);
        }
      }
      const aim = target(most);
      const describe = kind === 'list' ? say(paragraphsHeld > 1 ? 'One short point in each item' : 'One short point', aim !== undefined && paragraphsHeld > 1 ? `${aim} in all` : aim) : aim;
      text(room.block.index, kind === 'text' ? 'text' : 'list', bodyStyle, size, inBlock(room.rect, room.x, body.topOf + down, body.width, boxHeight), body.width, high, { slot: freeSlotName(room.block.index, part), required: true, paragraphs: true, columns, lines: perSlot, most, none, fitted, ...(spacing > 0 ? { give: spacing } : {}), ...(told !== undefined ? { told } : {}), ...(describe !== undefined ? { describe } : {}) });
    });
  }

  // ---- cards and steps
  for (const kind of ['card', 'step'] as const) {
    const blocks = blocksOf(kind);
    if (blocks.length === 0) continue;
    const onPanel = !!kit.panel;
    const withIcon = kind === 'card' && blocks.every((block) => has(block.index, kind, 'Icon'));
    const geometry = (rect: Rect, textSize: number, headSize: number) => {
      const p = onPanel ? pad(rect.w, rect.h) : 0;
      const lying = lies;
      const icon = withIcon ? clamp(0.175 * rect.w, 0.018 * Wpt, 0.0375 * Wpt) : 0;
      const number = kind === 'step' ? half(Math.min(figureN, 1.6 * headSize)) : 0;
      const lead = kind === 'step' ? number * 1.3 : icon;
      const x = rect.x + p + (lying && lead > 0 ? lead + 0.75 * Math.max(p, 0.008 * Wpt) : 0);
      const w = rect.x + rect.w - p - x;
      const y = rect.y + p + (!onPanel && kit.rule ? 0.012 * Wpt : 0);
      const headTop = lying || lead === 0 ? y : y + lead + (kind === 'step' ? 0.1 * number : Math.max(p, 0.008 * Wpt) * 0.75);
      return { p, lying, icon, number, x, w, y, headTop, bottom: rect.y + rect.h - p, lead };
    };
    // A card that took room that was left over is drawn as high as it grew, and holds what it
    // held: its heading and text are set out in the box it asked for, at the top of the panel,
    // as the template's own cards have theirs. So a card says the same before there is content
    // and after, whatever room its slide has to spare.
    const asked = (block: LaidBlock): Rect => {
      const rect = rectOf(block);
      return { ...rect, h: rect.h - (block.grown ?? 0) * Hpt };
    };
    const rects = blocks.map(asked);
    // Whether the icon or the number stands beside the heading, in a card that lies, or over it is
    // decided once for the slide, by its narrowest card: a short last line that is spread wider
    // keeps the form of the cards above it.
    const lies = Math.min(...rects.map((rect) => rect.w / rect.h)) >= 1.6;
    // What the width allows is worked out with the template's own sizes, whatever the deck's factor.
    const first = rects.map((rect) => geometry(rect, body1, head1));
    const wideText = Math.min(...first.map((entry) => halfDown(entry.w / (WIDE * 20))));
    const wideHead = Math.min(...first.map((entry) => halfDown(entry.w / (WIDE * 12))));
    const headFor = (size: number) => Math.max(floorOf(blocks[0]?.index, 'heading'), HARD, Math.min(headN, Math.max(LINE * size, (head1 * size) / body1), wideHead));
    /** The blocks with their text at a size: the heading's size (`headFor`, or `given`), where everything goes, and how many lines the longest heading of the slide takes there. */
    const set = (size: number, given?: number) => {
      let head = given ?? headFor(size);
      let all = rects.map((rect) => geometry(rect, size, head));
      // The lines of a heading are counted against the width its budget is stated for. Every
      // card keeps as many as the longest heading of the slide takes, so that the texts start on one line.
      const linesAt = (at: number, placed: typeof all): number => Math.max(1, ...placed.map((entry, i) => headingLines((blocks[i] as LaidBlock).index, perLine(entry.w, at))));
      let lines = linesAt(head, all);
      // A heading of three lines is set smaller, by up to 15%, when it then takes two; and one of
      // two lines when it then takes one, so that no card keeps a line of its heading empty
      // above its text for one heading that is a few characters too long. Never smaller than the
      // text under it.
      const full = head;
      for (const fewer of [2, 1]) {
        for (let smaller = halfDown(head - 0.5); lines > fewer && smaller >= Math.max(HARD, 0.85 * full, size) - 1e-9; smaller -= 0.5) {
          const placed = rects.map((rect) => geometry(rect, size, smaller));
          if (linesAt(smaller, placed) > fewer) continue;
          head = smaller;
          all = placed;
          lines = fewer;
        }
      }
      return { head, all, lines };
    };
    // Sizes follow the height too: the heading and one line of text have to end inside the card,
    // within its padding or, in a card too low for that, above its own edge. A line is as high
    // as the template sets it: with a line spacing of 130% the text is smaller in the same card.
    const headStyle = kit.styles.heading;
    const textStyle = kit.styles.body;
    /** From the top of a heading to the top of the text under it: its lines, and half its size of space. */
    const underHead = (head: number, lines: number): number => (headStyle.leading ? linesHigh(headStyle, lines, head) + 0.5 * head : (lines * LINE + 0.5) * head);
    const fitsIn = (size: number): boolean => {
      const { head, all, lines } = set(size);
      return all.every((entry, i) => entry.headTop + underHead(head, lines) + lineOf(textStyle, size) <= Math.max(entry.bottom + hang(size), (rects[i] as Rect).y + (rects[i] as Rect).h));
    };
    let textMost = Math.max(HARD, Math.min(bodyN, wideText));
    while (textMost > HARD && !fitsIn(textMost)) textMost = Math.max(HARD, halfDown(textMost - 0.25));
    /** The cards with their headings set for text of a size: the heading's size and lines, and where each text starts and how high it may be. */
    const roomsFor = (size: number, given?: number) => {
      const { head, all, lines } = set(size, given);
      const headHeight = linesHigh(headStyle, lines, head) + 0.15 * head;
      const rooms = blocks.map((block, i) => {
        const g = all[i] as (typeof all)[number];
        const rect = rects[i] as Rect;
        const textTop = g.headTop + headHeight + 0.35 * head;
        // The text has the room down to the card's padding. A card too low for a line of it there
        // still has that line, as far as its own edge.
        return { block, g, textTop, height: Math.max(LINE * HARD, g.bottom - textTop, Math.min(lineOf(textStyle, size), rect.y + rect.h - textTop)) };
      });
      return { head, lines, headHeight, rooms };
    };
    // The heading keeps the size the card's width and height give it, whatever size its text then
    // gets: a heading of a slide is not set smaller because its text is long.
    let laid = roomsFor(textMost);
    // Short text is set larger, but never as large as the heading over it, never larger than the
    // running text or the list of its slide, which it would then outweigh, and not at all in a
    // card that only just holds a line of it.
    const low = laid.rooms.some((room) => room.g.bottom - room.textTop < lineOf(textStyle, textMost));
    const grow = low ? 1 : Math.max(1, Math.min(blocksOnSlide <= 2 ? 1.6 : 1.25, laid.head / (1.05 * textMost), proseSize / textMost));
    const textFloor = floorOf(blocks[0]?.index, 'text');
    // The size the text is fitted at, and its budget stated at: what text written to its budget
    // is set at, down to four fifths of the size the card gives it.
    const statedSize = Math.max(textFloor, sized(laid.rooms.map((room) => ({ chars: said(room.block.index, 'Text'), paragraphs: 1, width: room.g.w, height: room.height, style: textStyle, card: true })), textMost, grow));
    // What is written is drawn at a size at which PowerPoint most likely sets it inside its box
    // (`setLinesOf`): half a point smaller where the count of characters says it fits and its
    // words do not, and smaller still where it is longer than its budget, down to three fifths of
    // the card's size and never under the size a projected slide is read at (9 pt on a slide
    // 13.33 inches wide). Such text is still reported at the size its budget is stated at.
    const likelyIn = (size: number, cards = laid): boolean =>
      cards.rooms.every((room) => linesHigh(textStyle, paragraphsOf(at(room.block.index, 'Text')).reduce((sum, paragraph) => sum + setLinesOf(paragraph, room.g.w, size), 0), size) <= room.height + hang(size));
    const textLeast = Math.min(statedSize, Math.max(textFloor, HARD, halfDown(0.6 * textMost), half(9 * k)));
    // Text longer than its budget: the count of its characters says it does not fit at the size it was fitted at.
    const overBudget = !planning && laid.rooms.some((room) => !fitsAt({ chars: said(room.block.index, 'Text'), paragraphs: 1, width: room.g.w, height: room.height, style: textStyle, card: true }, statedSize, hang));
    let textSize = statedSize;
    while (!planning && textSize > textLeast && !likelyIn(textSize)) textSize = Math.max(textLeast, textSize - 0.5);
    // Text that runs out of its card even at that size would cross the card's edge. Only then
    // does the heading give way: it is set smaller, half a point at a time and never under a
    // size and a sixth of the text, until the text stays inside, or to the largest size at
    // which it takes a line fewer, which gives the text the most room: of the two, the one that
    // leaves the text larger. The text takes what that leaves it.
    if (!planning && !likelyIn(textSize)) {
      const textIn = (cards: typeof laid): number => {
        let at = statedSize;
        while (at > textLeast && !likelyIn(at, cards)) at = Math.max(textLeast, at - 0.5);
        return at;
      };
      let fits: typeof laid | undefined;
      let fewer: typeof laid | undefined;
      for (let head = halfDown(laid.head - 0.5); head >= Math.max(HARD, (7 / 6) * textLeast) - 1e-9 && !(fits && fewer); head -= 0.5) {
        const tried = roomsFor(textMost, head);
        if (!fits && likelyIn(textLeast, tried)) fits = tried;
        if (!fewer && tried.lines < laid.lines) fewer = tried;
      }
      const better = fits && fewer ? (textIn(fewer) > textIn(fits) ? fewer : fits) : (fits ?? fewer);
      if (better) {
        laid = better;
        textSize = textIn(better);
      }
    }
    // Text over its budget that had to be drawn smaller makes the slide crowded (see `freePlan`); a
    // half point lost to how PowerPoint breaks the lines of text within its budget does not.
    if (overBudget && textSize < statedSize - 1e-9) squeezed = true;
    const { head: headMost, lines: headLines, headHeight, rooms } = laid;
    const twoLines = headLines > 1;
    for (const { block, g, textTop, height } of rooms) {
      const rect = asked(block);
      if (onPanel) panel(rectOf(block), block.index);
      else if (kit.rule) shapes.push({ id: String(nextId++), name: `Block ${block.index} rule`, kind: 'rule', from: { slide: kit.rule.slide, shapeId: kit.rule.shapeId }, box: toBox(rect.x, rect.y, rect.w, 0) });
      if (withIcon) {
        const id = String(nextId++);
        const box = inBlock(rect, rect.x + g.p, rect.y + g.p, g.icon, g.icon);
        shapes.push({ id, name: `Block ${block.index} icon`, kind: 'picture', box });
        const slot: ImageSlotDef = { name: freeSlotName(block.index, 'Icon'), kind: 'image', imageKind: 'icon', shapeId: id, fit: 'contain', box, empty: 'remove', description: 'An icon that suits the card. Give one to every card of the slide, or to none' };
        made.push(slot);
      }
      if (kind === 'step') {
        const ordinal = blocks.indexOf(block) + 1;
        text(block.index, 'number', kit.styles.figure, g.number, inBlock(rect, rect.x + g.p, rect.y + g.p, g.lying ? g.lead : g.w, LINE * g.number), g.w, LINE * g.number, { fixed: String(ordinal), wrap: false });
      }
      // A card's padding follows its height, and its height follows what is written: a card of
      // other text is a point or two narrower inside. Before there is content the heading's
      // budget is stated for the narrowest the card gets, so that a heading as long as it says
      // is a heading of one line whatever the card's text.
      const narrowest = planning && onPanel ? g.w - 2 * Math.max(0, clamp(0.09 * rect.w, 0.006 * Wpt, 0.03 * Wpt) - g.p) : g.w;
      text(block.index, 'heading', kit.styles.heading, headMost, inBlock(rect, g.x, g.headTop, g.w, headHeight), g.w, headHeight, { slot: freeSlotName(block.index, 'Heading'), required: true, most: perLine(narrowest, headMost) * (twoLines ? 2 : 1), describe: 'A heading of a few words', several: twoLines });
      // A card too low for one line of its text, even at the smallest size, holds none of it:
      // what is written for it is reported, not drawn across the card's edge in silence.
      const outside = textTop + lineOf(textStyle, textSize) > rect.y + rect.h + 0.05;
      // What the slot states: with content, what the fitting counted with, so that text that was
      // fitted is not reported; before there is content, what the box surely holds, with its
      // lines as high as the template sets them.
      const holds = planning ? capacity(g.w, height, textSize, hang, 1, lineOf(textStyle, textSize)) : cardCapacity(g.w, height, statedSize, textStyle, hang);
      // Its last line may hang into the card's padding (`hang`): the box is drawn that much higher,
      // where the card has the room, so that what PowerPoint sets stays inside its own box too.
      const drawnHigh = Math.max(height, Math.min(height + hang(textSize), rect.y + rect.h - textTop));
      text(block.index, 'text', textStyle, textSize, inBlock(rect, g.x, textTop, g.w, drawnHigh), g.w, height, { slot: freeSlotName(block.index, 'Text'), required: true, most: holds, none: outside || !holdsLine(height, textSize, lineOf(textStyle, textSize), hang) });
    }
  }

  // ---- figures
  {
    const blocks = blocksOf('figure');
    if (blocks.length > 0) {
      // A figure that is the slide's only block is the slide: it stands on the ground itself, large, where a sentence alone would stand.
      const lone = single;
      const onPanel = !!kit.panel && (kit.figureOnPanel || tiles) && !lone;
      const inner = blocks.map((block) => {
        const rect = rectOf(block);
        const p = onPanel ? pad(rect.w, rect.h) : 0;
        return { block, rect, p, w: rect.w - 2 * p, h: rect.h - 2 * p };
      });
      const labelStyle = kit.styles.body;
      const narrowest = Math.min(...inner.map((entry) => entry.w));
      // How high the value and its label may be together: the lowest card without its padding, counted
      // as the padding of a card as wide as it is high (a seventh of its height at each end), which
      // no card's padding is more than. So the size follows the height of the cards, and not how
      // wide they are: fewer figures, in cards no lower, are never set smaller than more.
      const lowest = Math.min(...inner.map((entry) => (onPanel ? Math.min(entry.h, entry.rect.h * (1 - 2 * 0.14)) : entry.h)));
      // The label of a figure alone keeps a line a person reads in one go.
      const labelWide = lone ? Math.min(narrowest, 0.5 * Wpt) : narrowest;
      // The fewer figures a slide has, the larger each is set: six or more have the template's
      // own size, one has up to 2.4 times that, and none is over two and a half times the title.
      const most = Math.min(figureN * clamp(Math.sqrt(6 / blocks.length), 1, 2.4), 2.5 * T);
      // A label is set at a size a person reads: about half the size of its value, between the
      // size of small running text and a quarter over that of running text.
      const labelAt = (value: number) => Math.max(floorOf(blocks[0]?.index, 'label'), HARD, halfDown(Math.min(1.25 * bodyN, Math.max(Math.min(0.9 * bodyN, value / 2), value / 2.6), labelWide / (WIDE * 14))));
      const labels = blocks.map((block) => said(block.index, 'Label'));
      /** The lines the longest label of the slide takes at a size: one to three, and two before there is content. */
      const labelLines = (size: number) => (planning ? 2 : clamp(Math.max(1, ...labels.map((chars) => Math.ceil(chars / (perLine(labelWide, size) * FULL)))), 1, 3));
      /** The lines PowerPoint most likely sets the longest label in: it sets a character narrower than the estimate has it. */
      const likelyLines = (size: number) => (planning ? 2 : clamp(Math.max(1, ...labels.map((chars) => Math.ceil(chars / (setAcross(labelWide, size) * 0.9)))), 1, 3));
      /** A value with its label under it: how high the two are together, the label in the lines it most likely takes. */
      const pairHigh = (value: number) => LINE * value + linesHigh(labelStyle, likelyLines(labelAt(value)), labelAt(value));
      /** How wide the widest value is, in sizes: a figure's face is heavy, and a % or an M is wider than a digit. */
      const wideness = Math.max(0.6, ...blocks.map((block) => [...paragraphsOf(at(block.index, 'Value')).join('')].reduce((sum, char) => sum + (/[%MWmw@]/.test(char) ? 0.9 : /[A-Z]/.test(char) ? 0.7 : /[.,:'\s]/.test(char) ? 0.32 : 0.64), 0)));
      // A value does not break: it is the largest at which the longest one stays in its box and
      // the value, a little air and the lines of the longest label fit the lowest card.
      const valueAt = (chars: number, wide = 0.6 * chars) => {
        let size = Math.max(HARD, halfDown(Math.min(most, narrowest / wide)));
        while (size > HARD && pairHigh(size) > lowest + hang(labelAt(size))) size = Math.max(HARD, size - 0.5);
        return size;
      };
      // Never smaller than a value of that length was set in a box of that size before figures
      // were sized by their room (the template's own size, at most 0.42 of the box's height),
      // and never wider than its box.
      const before = (chars: number) => Math.max(HARD, halfDown(Math.min(figureN, ...inner.map((entry) => Math.min(entry.w / (0.6 * chars), 0.42 * entry.h)))));
      const longest = Math.max(1, ...blocks.map((block) => said(block.index, 'Value')));
      const fitsAcross = planning ? Infinity : halfDown(narrowest / wideness);
      // The values written are counted by the kind of their letters (`wideness`); a value of a length, by its length.
      const valueSize = Math.max(floorOf(blocks[0]?.index, 'value'), valueAt(longest, planning ? 0.6 * longest : Math.max(0.6 * longest, wideness)), Math.min(before(longest), fitsAcross));
      // What a value may be is another matter: four fifths of the size a figure of ordinary length
      // has here. The budget is stated at that size, so that a longer value is reported; before
      // there is content it is stated at that size outright, so that a writer is not held to
      // the five characters the slide was set out for.
      const valueFloor = planning ? 0.8 * valueAt(5) : Math.max(valueSize, 0.8 * valueAt(5));
      // The labels keep to the size of a value that is no longer than a value may be, so that one long value
      // does not take every label down with it. But a label is never larger than the value over it,
      // and it is set smaller where the lines it most likely takes do not fit under the value.
      // The value's box is as high as PowerPoint sets its one line, so that nothing of it hangs under its box.
      const valueHeight = LINE * valueSize;
      const under = lowest - valueHeight;
      let labelSize = Math.max(floorOf(blocks[0]?.index, 'label'), HARD, Math.min(labelAt(Math.max(valueSize, 0.8 * valueAt(5))), valueSize));
      while (labelSize > Math.max(HARD, floorOf(blocks[0]?.index, 'label')) && linesHigh(labelStyle, likelyLines(labelSize), labelSize) > under + hang(labelSize)) labelSize = Math.max(HARD, labelSize - 0.5);
      // The label's box keeps the lines the estimate gives it where the card has room for them,
      // and the lines PowerPoint most likely sets it in where it has not.
      const generous = labelLines(labelSize);
      const lines = linesHigh(labelStyle, generous, labelSize) <= under + hang(labelSize) ? generous : Math.min(generous, likelyLines(labelSize));
      const labelHeight = linesHigh(labelStyle, lines, labelSize);
      // Between the two there is a fifth of the value's size of air, where the lowest card has it.
      const air = Math.max(0, Math.min(0.2 * valueSize, lowest - valueHeight - labelHeight));
      const pair = valueHeight + air + labelHeight;
      // The lines of the label that PowerPoint most likely leaves empty.
      const likely = planning ? lines : Math.min(lines, likelyLines(labelSize));
      const spared = (lines - likely) * lineOf(labelStyle, labelSize);
      for (const { block, rect, p, w, h } of inner) {
        if (onPanel) panel(rect, block.index);
        // The number and its label sit together in the middle of their card, with as much
        // space over the number as under the label; a figure alone stands a little above the
        // middle of the slide. Cards of one height have them on one line.
        // Where they sit is worked out from the lines PowerPoint is likely to set the label in,
        // which can be one fewer than the lines its box keeps for it: the pair is then in the
        // middle as it is drawn, and a label that takes every line of its box still ends
        // inside the card, at most half the padding lower.
        const down = Math.max(0, Math.min((h - (pair - spared)) * (lone ? 0.3 : 0.5), h - pair + 0.5 * p));
        text(block.index, 'value', kit.styles.figure, valueSize, inBlock(rect, rect.x + p, rect.y + p + down, w, valueHeight), w, valueHeight, { slot: freeSlotName(block.index, 'Value'), required: true, wrap: false, most: Math.max(2, Math.floor(w / (0.6 * valueFloor))), describe: 'The figure itself, such as 87% or 2.4M. When the brief gives no number for it, leave it empty' });
        const wide = Math.min(w, labelWide);
        // What the slot holds: what its box holds by the estimate, and a label that was fitted by
        // the lines PowerPoint most likely sets it in is within it.
        const held = capacity(wide, labelHeight, labelSize, hang, 1, lineOf(labelStyle, labelSize));
        const own = said(block.index, 'Label');
        const fitted = !planning && Math.ceil(own / (setAcross(labelWide, labelSize) * 0.9)) <= lines ? own : 0;
        text(block.index, 'label', labelStyle, labelSize, inBlock(rect, rect.x + p, rect.y + p + down + valueHeight + air, wide, labelHeight), wide, labelHeight, { slot: freeSlotName(block.index, 'Label'), required: true, most: Math.max(held, fitted), describe: say('What the figure counts', target(held)) });
      }
    }
  }

  // ---- quotations
  {
    const blocks = blocksOf('quote');
    if (blocks.length > 0) {
      const rooms = blocks.map((block) => {
        const rect = rectOf(block);
        const onPanel = tiles && !!kit.panel;
        const p = onPanel ? pad(rect.w, rect.h) : 0;
        return { block, rect, onPanel, p, x: rect.x + p, y: rect.y + p, w: rect.w - 2 * p, h: rect.h - 2 * p, source: has(block.index, 'quote', 'Source') };
      });
      const quoteStyle = kit.styles.quote;
      const most = Math.max(HARD, Math.min(single ? Math.max(quoteN, half(0.85 * T)) : quoteN, ...rooms.map((room) => halfDown(room.w / (WIDE * 18)))));
      const widths = rooms.map((room) => Math.min(room.w, 50 * WIDE * most));
      // The source is the small line under a quotation: half the quotation's size, never under
      // a caption's and never over running text's. It is set in the template's accent, the
      // colour of its big numbers, where the quotation has another colour: a source in the size
      // and colour of a heading reads as the heading of whatever stands under it.
      const sourceAt = (quote: number) => Math.max(floorOf(rooms.find((room) => room.source)?.block.index, 'source'), HARD, Math.min(halfDown(clamp(0.5 * quote, Math.min(captionN, bodyN), bodyN)), ...rooms.map((room) => halfDown(room.w / (WIDE * 20)))));
      const sourceHigh = (at: number) => lineOf(kit.styles.heading, at) + 0.2 * at;
      const size = Math.max(floorOf(blocks[0]?.index, 'quote'), HARD, fitSize(rooms.map((room, i) => ({ chars: said(room.block.index, 'Quote'), paragraphs: 1, width: widths[i] as number, height: room.h - (room.source ? sourceHigh(sourceAt(most)) + 0.5 * most : 0), style: quoteStyle })), most, Math.max(HARD, Math.min(bodyN, most)), hang));
      const sourceSize = sourceAt(size);
      const sourceHeight = sourceHigh(sourceSize);
      const accent = kit.styles.figure.set?.color ?? kit.styles.figure.color;
      const sourceColor = accent !== (quoteStyle.set?.color ?? quoteStyle.color) ? accent : (kit.styles.body.set?.color ?? kit.styles.body.color);
      const set = rooms.map((room, i) => {
        const width = widths[i] as number;
        const needed = linesHigh(quoteStyle, Math.max(1, Math.ceil(said(room.block.index, 'Quote') / (perLine(width, size) * FULL))), size);
        const room4 = room.h - (room.source ? sourceHeight + 0.5 * size : 0);
        return { room, width, room4, height: Math.min(room4, needed + 0.3 * size) };
      });
      for (const { room, width, room4, height: own } of set) {
        if (room.onPanel) panel(room.rect, room.block.index);
        // Quotations that stand side by side have boxes of one height, so that their sources
        // stand on one line. Every quotation has its text at the foot of its box: its source then
        // follows its last line at half its size, however many lines PowerPoint sets it in, and
        // quotations side by side have their sources at one distance from their last lines.
        const beside = set.filter((other) => Math.abs(other.room.rect.y - room.rect.y) < 0.5);
        const height = Math.min(room4, Math.max(...beside.map((other) => other.height)));
        const down = single ? Math.max(0, 0.4 * (room4 - own)) : 0;
        const held = capacity(width, room4, size, hang, 1, lineOf(quoteStyle, size));
        const none = !holdsLine(room4, size, lineOf(quoteStyle, size), hang);
        text(room.block.index, 'quote', quoteStyle, size, inBlock(room.rect, room.x, room.y + down, width, height), width, planning ? room4 : height, { slot: freeSlotName(room.block.index, 'Quote'), required: true, most: held, none, anchor: 'b', describe: say('The quotation, word for word, without quotation marks', none ? undefined : target(held)) });
        if (room.source) text(room.block.index, 'source', kit.styles.heading, sourceSize, inBlock(room.rect, room.x, room.y + down + height + 0.5 * size, width, sourceHeight), width, sourceHeight, { slot: freeSlotName(room.block.index, 'Source'), most: perLine(width, sourceSize), describe: 'Who said it, when the brief says who. Never make a name up', ...(contrast(sourceColor, groundOf(room.onPanel)) >= 3 ? { color: sourceColor } : {}) });
      }
    }
  }

  // ---- notes
  // On a slide whose notes ask for more than their share of the room, every note was set out
  // for its text at one smaller size (`layout.lower`). The notes of such a slide share that
  // size, like parts of like blocks do. A band that is lower still than its text needs there
  // (more notes than the room holds at the smallest size) has the largest size that fits it,
  // and then all of them have that one.
  const lower = layout.lower ?? 1;
  const notes = blocksOf('note').map((block) => {
    const rect = rectOf(block);
    let size = noteSize(rect.w, lower);
    const needs = (at: number) => linesHigh(noteStyle, noteLines(block.index, rect.w, at), at) + noteAround(at, lower);
    if (lower < 1) while (size > HARD && needs(size) > rect.h + 0.05) size = Math.max(HARD, halfDown(size - 0.25));
    return { block, rect, size };
  });
  const lowered = lower < 1 ? Math.min(...notes.map((note) => note.size)) : undefined;
  for (const { block, rect, size: own } of notes) {
    const size = lowered ?? own;
    const width = rect.w - 2 * noteInset;
    const around = (at: number) => noteAround(at, lower);
    const lines = Math.max(1, Math.round((rect.h - around(size)) / lineOf(noteStyle, size)));
    const high = linesHigh(noteStyle, lines, size);
    // Many notes are set smaller, and a line of small text holds more. Before there is content a note is
    // told no more than a note alone on the slide: it is the one line to take away.
    const alone = noteSize(rect.w);
    // One line holds what its width holds; several lose a little where they break at words.
    const held = (at: number, count: number) => (count === 1 ? perLine(width, at) : Math.floor(perLine(width, at) * count * FULL));
    // A note that was set on one line by the count of its face holds what it says.
    const holds = Math.max(Math.min(held(size, lines), planning ? held(alone, noteLines(block.index, rect.w, alone)) : Infinity), lines === 1 && oneLine(block.index, rect.w, size) ? said(block.index, 'Text') : 0);
    // A band squeezed lower than a line of its text holds none of it.
    const none = !holdsLine(rect.h, size, lineOf(noteStyle, size), hang);
    const options = { slot: freeSlotName(block.index, 'Text'), required: true, most: holds, none, anchor: 'ctr' as const, inset: noteInset, describe: say('The one line to take away', none ? undefined : target(holds)) as string };
    if (kit.styles.note) {
      text(block.index, 'note', kit.styles.note, size, toBox(rect.x, rect.y, rect.w, rect.h), rect.w, high, options);
    } else if (kit.notePanel) {
      shapes.push({ id: String(nextId++), name: `Block ${block.index} panel`, kind: 'panel', from: { slide: kit.notePanel.slide, shapeId: kit.notePanel.shapeId }, box: toBox(rect.x, rect.y, rect.w, rect.h) });
      text(block.index, 'note', kit.notePanel.text, size, toBox(rect.x, rect.y, rect.w, rect.h), rect.w, high, options);
    } else {
      text(block.index, 'note', kit.styles.heading, size, toBox(rect.x, rect.y, rect.w, rect.h), rect.w, high, { ...options, inset: 0 });
    }
  }

  // ---- charts, tables and pictures: a caption strip over or under, and the thing itself
  const visuals = layout.blocks.filter((block) => block.kind === 'chart' || block.kind === 'table' || block.kind === 'photo' || block.kind === 'graphic');
  for (const kind of ['chart', 'table', 'photo', 'graphic'] as const) {
    const blocks = visuals.filter((block) => block.kind === kind);
    if (blocks.length === 0) continue;
    // The strip is kept for every block of a kind when any of them has a caption, so that they line up.
    const captioned = blocks.some((block) => has(block.index, kind, 'Caption'));
    const over = kind === 'chart' || kind === 'table';
    const rooms = blocks.map((block) => {
      const rect = rectOf(block);
      // Among tiles a chart stands on a panel like its neighbours, when the slide has blocks that
      // are drawn on one. Charts, tables and pictures alone need none. A table never has one: it
      // brings its own ground, and a panel would cost it the width its columns need.
      const onPanel = tiles && !!kit.panel && kind === 'chart' && layout.blocks.some((other) => other.kind !== 'chart' && other.kind !== 'table' && other.kind !== 'photo' && other.kind !== 'graphic' && other.kind !== 'note');
      const p = onPanel ? pad(rect.w, rect.h) : 0;
      return { block, rect, onPanel, x: rect.x + p, y: rect.y + p, w: rect.w - 2 * p, h: rect.h - 2 * p };
    });
    const charts: ChartPlace[] = [];
    const captionMost = Math.max(floorOf(blocks.find((block) => has(block.index, kind, 'Caption'))?.index, 'caption'), HARD, Math.min(captionN, ...rooms.map((room) => halfDown(room.w / (WIDE * 20)))));
    /**
     * The lines the longest caption of the kind takes at a size, as PowerPoint most likely sets
     * them (`setLinesOf`). Before there is content a caption is one line.
     */
    const captionLinesAt = (size: number): number =>
      planning ? 1 : rooms.reduce((most, room) => Math.max(most, has(room.block.index, kind, 'Caption') ? setLinesOf(paragraphsOf(at(room.block.index, 'Caption')).join(' '), room.w, size) : 1), 1);
    // A caption that takes a second line is set smaller, by up to 15%, where it then takes one.
    // One that still takes more keeps its size, and its strip is as high as its lines, at most
    // three: what it names starts under it, as it is really set, and not under the line it was
    // planned to take.
    let captionSize = captionMost;
    let captionLines = Math.min(3, captionLinesAt(captionMost));
    for (let smaller = halfDown(captionMost - 0.5); captionLines > 1 && smaller >= Math.max(HARD, 0.85 * captionMost) - 1e-9; smaller -= 0.5) {
      if (captionLinesAt(smaller) > 1) continue;
      captionSize = smaller;
      captionLines = 1;
    }
    // A caption has one line, in a strip one and a half times its size high, or as high as the template sets that line.
    // In a block too low for both, the strip takes half of the height and no more: the thing itself keeps the other half.
    // Over a chart the strip is a quarter of its size high and a fifth apart, no more: the plot
    // keeps space over its bars of its own, and a low chart needs every point of its height.
    const lineBox = kind === 'chart' ? 1.25 : 1.5;
    const apart = kind === 'chart' ? 0.2 : 0.3;
    const captionLine = Math.max(lineBox * captionSize, lineOf(kit.styles.caption, captionSize)) + (captionLines - 1) * lineOf(kit.styles.caption, captionSize);
    const share = Math.min(1, ...rooms.map((room) => (0.5 * room.h) / (captionLine + apart * captionSize)));
    const captionHigh = captionLine * share;
    const strip = captioned ? captionHigh + apart * captionSize * share : 0;
    for (const room of rooms) {
      if (room.onPanel) panel(room.rect, room.block.index);
      const index = room.block.index;
      let area: Rect = { x: room.x, y: room.y + (over ? strip : 0), w: room.w, h: room.h - strip };
      // A caption stands at the edge of what it names: as wide as a chart, a table or a photo
      // and starting where it starts, and under the middle of a graphic, which is shown whole in
      // the middle of its box.
      const captionAt = (y: number, x = room.x, w = room.w) => {
        // What the slot holds: a line's worth by the estimate, and a caption that PowerPoint most
        // likely sets on one line is within it.
        const written = planning ? '' : paragraphsOf(at(index, 'Caption')).join(' ');
        const held = Math.max(perLine(w, captionSize), written !== '' && setLinesOf(written, w, captionSize) <= 1 ? written.length : 0);
        if (captioned && has(index, kind, 'Caption')) text(index, 'caption', kit.styles.caption, captionSize, inBlock(room.rect, x, y, w, captionHigh), w, captionHigh, { slot: freeSlotName(index, 'Caption'), most: held, describe: `What it shows, in a few words. Give one to every ${kind} of the slide, or to none`, several: captionLines > 1, ...(kind === 'graphic' ? { align: 'ctr' as const } : {}) });
      };
      if (over) captionAt(room.y);

      if (kind === 'chart') {
        const id = String(nextId++);
        const box = toBox(area.x, area.y, area.w, area.h);
        shapes.push({ id, name: `Block ${index} chart`, kind: 'anchor', box });
        // The charts of a slide are worked out together, once all have their place: like charts share an axis.
        charts.push({ index, id, box, area, ground: groundOf(room.onPanel), tile: room.onPanel });
      } else if (kind === 'table') {
        const value = at(index, 'Table');
        const given = !planning && isTableValue(value) ? value : undefined;
        // What each cell holds as it will be drawn, and how long it is, row by row, the heading row first.
        const written = given ? [given.header, ...given.rows].map((row) => row.map((cell) => plainText(cell).trim())) : [];
        const cells = written.map((row) => row.map((cell) => cell.length));
        const rows = given ? cells.length : 5;
        const columns = Math.max(1, given ? given.header.length : 1);
        const sample = kit.table;
        const base = sample ? sample.size * deckScale * (sample.size <= 12 ? kit.textScale : 1) : bodyN * 0.85;
        const least = half(Math.min(6 * k, base));
        const sum = (values: number[]) => values.reduce((total, next) => total + next, 0);
        /** The width of the text in a cell: a column without what the renderer keeps free at its sides. */
        const cellWidth = (count: number) => area.w / count - 2 * Math.min(7.2, (0.06 * area.w) / count);
        // What a cell keeps free at each side: the renderer narrows it in narrow columns, and leaves
        // a sample table's own margin where it does not, which may be wider (fitTable).
        const narrowed = Math.min(7.2, (0.06 * area.w) / columns);
        const side = narrowed < 7.2 - 1e-9 ? narrowed : Math.max(narrowed, sample?.margin ?? 7.2);
        const column = (pick: (cell: string, row: number) => number) => Array.from({ length: columns }, (_, j) => written.reduce((most, row, i) => Math.max(most, pick(row[j] ?? '', i)), 0) + 2 * side);
        /**
         * How wide each column has to be at a size for its longest word (a heading's a little wider, for its
         * heavier face) to stand on one line: its letters counted by their kind, since a word of capitals and
         * wide letters (Medium) is wider than its length says.
         */
        const wordsAt = (at: number): number[] => column((cell, row) => unbroken(cell).reduce((most, word) => Math.max(most, wordWide(word)), 0) * at * (row === 0 ? 1.05 : 1));
        /**
         * How wide each column is at a size: as wide as its longest cell asks (a heading a little
         * more, for its heavier face), all made wider by one factor when the table has room to
         * spare. In a table that has not, every column is first as wide as its longest word, so
         * that no word is broken where another column has width to give, and the width that is
         * left is shared by what each column asks for beyond that. Only where the longest words
         * alone are wider than the table do the columns share the width by them. A table of
         * words of one length has columns of one width.
         */
        const widthsAt = (at: number): number[] => {
          const natural = column((cell, row) => cell.length * WIDE * at * (row === 0 ? 1.05 : 1));
          const total = sum(natural);
          if (total <= area.w) return natural.map((width) => (width * area.w) / total);
          const words = wordsAt(at);
          const least = sum(words);
          if (least >= area.w) return words.map((width) => (width * area.w) / least);
          const beyond = natural.map((width, j) => Math.max(0, width - (words[j] as number)));
          const asking = sum(beyond);
          return words.map((width, j) => width + (asking > 0 ? ((area.w - least) * (beyond[j] as number)) / asking : (area.w - least) / columns));
        };
        /** The lines each row takes at a size: those of its fullest cell, each cell counted in the width its column has. */
        const linesAt = (at: number): number[] => {
          if (!given) return Array<number>(rows).fill(1);
          const across = widthsAt(at).map((width) => perLine(width - 2 * side, at));
          return cells.map((row) => row.reduce((most, chars, j) => Math.max(most, chars <= (across[j] as number) ? 1 : Math.ceil(chars / ((across[j] as number) * 0.85))), 1));
        };
        /** Whether so many lines of text at a size fit the room, each with the space a row keeps above and below it. */
        const high = area.h;
        const holds = (count: number, at: number) => tableTextSize(high * 12700, count, at) >= at;
        // The text is set smaller, in half points, until every row is as high as its fullest cell needs.
        // It starts no higher than the room is: no size above that holds a line, and a text size
        // that is no size (a deck's factor of a million) must not keep this counting for minutes.
        // It is also set smaller, by up to a fifth, where that lets the longest word of every
        // column stand on one line.
        const ceiling = Math.max(4, high);
        let size = Math.max(least, half(Math.min(base, ceiling)));
        const wordLeast = Math.max(least, half(0.8 * size));
        const wordsFit = (at: number) => !given || sum(wordsAt(at)) <= area.w + 0.01;
        while (size > least && (!holds(sum(linesAt(size)), size) || (size > wordLeast && !wordsFit(size)))) size = Math.max(least, size - 0.5);
        const fits = holds(sum(linesAt(size)), size);
        // Rows are made as high as the lines PowerPoint most likely sets them in (`setLinesOf`, the
        // heading row in its heavier face): rows that are set in as many lines are as high, whatever
        // the estimate above counted for each, and a row that PowerPoint sets on one line is not
        // kept as high as one of two.
        const linesOf = (widths: number[]): number[] => written.map((row, i) => row.reduce((most, cell, j) => Math.max(most, setLinesOf(cell, (widths[j] as number) - 2 * side, size, i === 0 ? 1 : SET_SHARE)), 1));
        // A cell a few characters too long for its column takes a whole line more, and its row is
        // as high as two for nothing. A column gives up to a fifth of its width to another where
        // that sets some row in fewer lines and no row in more, and no word of its own then breaks.
        // A table of a reader's size is looked at so; a larger one keeps its widths as they are.
        const setWidths = widthsAt(size);
        if (given && columns <= 10 && written.length * columns <= 300) {
          const words = wordsAt(size);
          /** The lines each cell of a column takes at a width, row by row. */
          const columnLines = (j: number, width: number): number[] => written.map((row, i) => setLinesOf(row[j] ?? '', width - 2 * side, size, i === 0 ? 1 : SET_SHARE));
          const perColumn = setWidths.map((width, j) => columnLines(j, width));
          const rowsOf = (columnsNow: number[][]): number[] => written.map((_, i) => columnsNow.reduce((most, column) => Math.max(most, column[i] as number), 1));
          for (let round = 0; round < columns; round++) {
            const was = rowsOf(perColumn);
            const now = sum(was);
            let best: { to: number; from: number; moved: number; toLines: number[]; fromLines: number[]; total: number } | undefined;
            for (let to = 0; to < columns; to++) {
              for (let from = 0; from < columns; from++) {
                if (from === to) continue;
                for (let share = 0.02; share <= 0.2 + 1e-9; share += 0.02) {
                  const moved = share * (setWidths[to] as number);
                  if ((setWidths[from] as number) - moved < (words[from] as number)) break;
                  const toLines = columnLines(to, (setWidths[to] as number) + moved);
                  const fromLines = columnLines(from, (setWidths[from] as number) - moved);
                  const lines = rowsOf(perColumn.map((column, j) => (j === to ? toLines : j === from ? fromLines : column)));
                  const total = sum(lines);
                  if (total < now && lines.every((count, row) => count <= (was[row] as number)) && (!best || total < best.total)) best = { to, from, moved, toLines, fromLines, total };
                  if (total < now) break;
                }
              }
            }
            if (!best) break;
            setWidths[best.to] = (setWidths[best.to] as number) + best.moved;
            setWidths[best.from] = (setWidths[best.from] as number) - best.moved;
            perColumn[best.to] = best.toLines;
            perColumn[best.from] = best.fromLines;
          }
        }
        const lines = given ? linesOf(setWidths) : Array<number>(rows).fill(1);
        /**
         * The size a table of these rows and columns has in this room when every cell holds one
         * word (eight characters): what the room alone asks of it, whatever is written in it.
         */
        const plainFrom = (from: number): number => {
          const wordLines = (at: number): number => {
            const across = perLine(cellWidth(columns), at);
            return rows * (across >= 8 ? 1 : Math.ceil(8 / (across * 0.85)));
          };
          let at = Math.max(least, half(Math.min(from, ceiling)));
          while (at > least && !holds(wordLines(at), at)) at = Math.max(least, at - 0.5);
          return at;
        };
        // What a table may hold is another matter: what the room takes with the text at four
        // fifths of that size, counted without a deck's wish for larger text. Cells that take
        // the text under it are reported on the slot, though the table is still drawn so that it
        // fits: shorter cells make its text larger.
        const reads = Math.max(least, half(0.8 * Math.min(plainFrom(base), plainFrom(sample ? sample.size * (sample.size <= 12 ? kit.textScale : 1) : body1 * 0.85))));
        const drawn = sample && sample.rowHeight > 0 ? (sample.rowHeight / 12700) * (size / sample.size) : 2.2 * size;
        // A row of one line is as high as the template draws its rows.
        const one = Math.max(drawn, 1.9 * size);
        let height: number;
        let rowHeights: number[] | undefined;
        // Rows that leave their frame free height are made higher, by up to two fifths, so that a
        // table of a few rows does not hang at the top of a block whose neighbour fills it. A table
        // alone on its slide hangs from the top line at its own rows' height, as a text alone does.
        const ROOMY = single ? 1 : 1.4;
        if (lines.every((count) => count === 1)) {
          height = Math.min(area.h / rows, ROOMY * one) * rows;
        } else {
          // A row of several lines needs them and the space kept above and below; it is given
          // the comfort of a row of one line on top while the room allows.
          const needs = lines.map((count) => Math.min(count * 1.25 * size + 7.2, (count * 1.25 * size) / 0.8));
          const wishes = lines.map((count, row) => Math.max(needs[row] as number, one + (count - 1) * 1.25 * size));
          const spare = area.h - sum(needs);
          const roomy = sum(wishes) <= area.h ? Math.min(ROOMY, area.h / sum(wishes)) : 1;
          const heights = sum(wishes) <= area.h ? wishes.map((wish) => wish * roomy) : spare <= 0 ? needs.map((need) => (need * area.h) / sum(needs)) : needs.map((need, row) => need + ((wishes[row] as number) - need) * (spare / (sum(wishes) - sum(needs))));
          height = sum(heights);
          rowHeights = heights.map((value) => Math.round(value * 100) / 100);
        }
        area = { ...area, h: height };
        const id = String(nextId++);
        const box = toBox(area.x, area.y, area.w, height);
        const shape: MadeShape = { id, name: `Block ${index} table`, kind: 'table', box };
        if (sample) shape.from = { slide: sample.slide, shapeId: sample.shapeId };
        smallest = Math.min(smallest, size);
        const header = sample?.slot.headerText ?? { bold: true, color: 'FFFFFF' };
        // A table the code draws has nothing behind its rows: its text has to read on the ground itself.
        const ground = groundOf(room.onPanel);
        const bodyColor = sample ? undefined : legible(undefined, ground);
        const bodyText = sample?.slot.bodyText ?? (bodyColor ? { color: bodyColor } : {});
        if (bodyColor) shape.look = { band: '', line: faint(bodyColor, ground) };
        shapes.push(shape);
        const columnsHeld = Math.max(2, Math.floor(room.w / (0.07 * Wpt)));
        const slot: FreeTableSlot = { name: freeSlotName(index, 'Table'), kind: 'table', shapeId: id, box, headerText: { ...header, size }, bodyText: { ...bodyText, size } };
        if (rowHeights) slot.rowHeights = rowHeights;
        if (given) {
          const widths = setWidths.map((width) => Math.round(width * 10) / 10);
          if (widths.some((width) => Math.abs(width - (widths[0] as number)) > 0.05)) slot.columnWidths = widths;
          // A column of numbers is read by its last digits: in a deck read from the left its cells
          // end on one edge, the heading with them. Read from the right every column does already.
          const numeric = Array.from({ length: columns }, (_, j) => given.rows.length > 0 && written.slice(1).every((row) => NUMBER_CELL.test(row[j] ?? '')));
          if (!rtl && numeric.some(Boolean)) slot.columnAligns = numeric.map((yes) => (yes ? 'r' : undefined));
        }
        const comfortable = { columns: Math.min(columnsHeld, 8), rows: Math.max(2, Math.floor(room.h / (1.9 * Math.max(6 * k, size * 0.7))) - 1) };
        if (given) {
          // What the room holds at that size: so many lines of text, shared by the rows there are.
          let held = Math.ceil(high / reads) + 1;
          while (held > 0 && !holds(held, reads)) held--;
          const each = rows <= held ? Math.floor(held / rows) : 0;
          const across = perLine(cellWidth(columns), reads);
          slot.maxRows = Math.max(0, held - 1);
          slot.maxColumns = Math.max(1, Math.floor((0.88 * area.w) / (4 * WIDE * reads)));
          slot.maxCellChars = each <= 1 ? across : Math.floor(across * 0.85 * each);
          // A table that fits at that size or a larger one is within what its box holds, however
          // its rows share the room.
          if (fits && size >= reads) {
            slot.maxRows = Math.max(slot.maxRows, rows - 1);
            slot.maxColumns = Math.max(slot.maxColumns, columns);
            for (const row of cells) for (const chars of row) slot.maxCellChars = Math.max(slot.maxCellChars, chars);
          }
        } else {
          // Before there is content, what a table is written to: the numbers that read well.
          slot.maxColumns = comfortable.columns;
          slot.maxRows = comfortable.rows;
          slot.maxCellChars = Math.max(4, perLine(cellWidth(comfortable.columns), Math.max(least, size * 0.7)));
        }
        slot.description = 'A table; a word or a short phrase in each cell. When the brief gives nothing for it, leave it empty';
        made.push(slot);
      } else {
        // A photo keeps a shape it reads in; a graphic is shown whole, whatever its shape. A
        // photo fills its box by being cut, so a wide strip is a shape too: only a box more than
        // 3.2 times as wide as high is cut narrower, and the photo then stands at the edge
        // reading starts at, where the title and its caption start.
        let frame = area;
        if (kind === 'photo') {
          const shapeOf = area.w / area.h;
          if (shapeOf > 3.2) frame = { ...area, w: 3.2 * area.h };
          else if (shapeOf < 0.6) frame = { ...area, y: area.y + (area.h - area.w / 0.6) / 2, h: area.w / 0.6 };
        }
        const id = String(nextId++);
        const box = inBlock(room.rect, frame.x, frame.y, frame.w, frame.h);
        const shape: MadeShape = { id, name: `Block ${index} ${kind}`, kind: 'picture', box };
        if (kind === 'photo' && kit.picture) shape.outline = kit.picture;
        shapes.push(shape);
        const slot: ImageSlotDef = { name: freeSlotName(index, kind === 'photo' ? 'Photo' : 'Graphic'), kind: 'image', imageKind: kind, shapeId: id, fit: kind === 'photo' ? 'cover' : 'contain', box };
        made.push(slot);
        area = frame;
      }
      // Under a picture cut to a sliver, on a slide of far too many blocks, the caption keeps the width of a short line where its block has it.
      if (!over) captionAt(area.y + area.h + apart * captionSize * share, area.x, Math.min(room.w, Math.max(area.w, 12 * WIDE * captionSize)));
    }
    if (charts.length > 0) made.push(...chartSlots(charts));
  }

  // ---- the legend of a slide of charts: a key and a name for each series, from the edge reading starts at
  if (legend) {
    if (bare) return undefined;
    const L = legend.size;
    const colors = kit.chart.colors ?? [];
    const ground = groundOf(false);
    const stripTop = (layout.dense ? denseTop : top) * Hpt;
    const side = 0.7 * L;
    const ink = legible(kit.chart.text?.color, ground) ?? kit.chart.text?.color;
    const start = kit.area.x * Wpt;
    const width = kit.area.w * Wpt;
    /** A box counted from the edge reading starts at, mirrored in the room for a deck read from the right. */
    const inRoom = (x: number, y: number, w: number, h: number): SlideBox => toBox(rtl ? start + width - (x - start) - w : x, y, w, h);
    let x = start;
    legend.names.forEach((name, at) => {
      const wide = name.length * WIDE * L + 0.5 * L;
      shapes.push({ id: String(nextId++), name: `Legend key ${at + 1}`, kind: 'key', box: inRoom(x, stripTop + (1.4 * L - side) / 2, side, side), color: seriesColor(colors.length > 0 ? colors : ['808080'], at, ground) });
      const style = kit.styles.caption;
      const label: MadeShape = { id: String(nextId++), name: `Legend ${at + 1}`, kind: 'text', from: { slide: style.slide, shapeId: style.shapeId }, box: inRoom(x + side + 0.4 * L, stripTop, wide, 1.4 * L), size: L, bold: false, italic: false, align: 'start', anchor: 'ctr', wrap: false, inset: 0, text: name };
      if (ink) label.color = ink;
      shapes.push(label);
      smallest = Math.min(smallest, L);
      x += side + 0.4 * L + wide + 1.2 * L;
    });
  }

  // The slots in the order of the blocks, each block's in the order they are read.
  const orderOf = new Map<string, number>();
  kinds.forEach((kind, position) => FREE_PARTS[kind].forEach((entry, place) => orderOf.set(freeSlotName(position + 1, entry.name), position * 10 + place)));
  made.sort((a, b) => (orderOf.get(a.name) ?? 0) - (orderOf.get(b.name) ?? 0));
  // Shapes are drawn block by block, a panel first: their ids say so, and a block's shapes stay together.
  const blockOf = (shape: MadeShape) => Number(/^Block (\d+) /.exec(shape.name)?.[1] ?? 0);
  const rank = (shape: MadeShape) => (shape.kind === 'panel' || shape.kind === 'rule' ? 0 : 1);
  shapes.sort((a, b) => blockOf(a) - blockOf(b) || rank(a) - rank(b) || Number(a.id) - Number(b.id));

  // Everything stays inside its block. On a slide that holds too many blocks one of them can be
  // lower than a line of its text, or than its heading: a shape that reached over the edge of
  // its block would lie on the neighbour. It is cut to its block, and one that lies wholly
  // outside gets the last sliver of it. Such a slide is crowded, and its texts are reported where
  // their boxes hold no line; but no two blocks share a place. A slide whose blocks have their
  // sizes is not touched: nothing on it reaches over an edge by more than half a thousandth of the slide.
  const placeOf = new Map(layout.blocks.map((block) => [block.index, block.box]));
  const round4 = (value: number) => Math.round(value * 10000) / 10000;
  const within = (start: number, extentOf: number, low: number, high: number): [number, number] => {
    if (start >= low - 0.0005 && start + extentOf <= high + 0.0005) return [start, extentOf];
    const from = round4(clamp(start, low, Math.max(low, high - 0.0001)));
    const to = round4(clamp(start + extentOf, from, high));
    return [from, extentOf > 0 ? Math.max(0.0001, round4(to - from)) : 0];
  };
  for (const shape of shapes) {
    const block = placeOf.get(blockOf(shape));
    if (!block || shape.kind === 'rule' || !(block.w > 0 && block.h > 0)) continue;
    const [x, w] = within(shape.box.x, shape.box.w, block.x, block.x + block.w);
    const [y, h] = within(shape.box.y, shape.box.h, block.y, block.y + block.h);
    // The slot of a shape shares its box: both are changed in place.
    if (x !== shape.box.x || w !== shape.box.w || y !== shape.box.y || h !== shape.box.h) Object.assign(shape.box, { x, y, w, h });
  }

  const frame: SlotDef[] = [];
  if (kit.title) {
    const title: TextSlotDef = { ...structuredClone(kit.title.slot), name: 'title', empty: 'remove', maxChars: kit.title.maxChars };
    delete title.deckField;
    if (kit.title.place) title.box = kit.title.box;
    else delete title.box;
    frame.push(title);
  }
  frame.push(...kit.frameSlots.map((slot) => structuredClone(slot)));

  const pattern: PatternDef = {
    id: FREE_PATTERN,
    sourceSlide: kit.baseSlide,
    description: `A slide the software lays out itself: ${describeBlocks(kinds)}.`,
    slots: [...frame, ...made],
  };
  // Under 9 pt on a slide 13.33 inches wide nobody reads a projected slide: a slide whose smallest text is under that is told.
  return { pattern, shapes, layout, smallest: Number.isFinite(smallest) ? smallest : bodyN, comfortable: Math.min(readableSize(kit.slideWidth), leastSizeOf(kit)), ...(squeezed ? { squeezed } : {}), ...(cramped ? { cramped } : {}) };
}
