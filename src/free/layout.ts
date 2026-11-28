// Where the blocks of a free slide go. Geometry only: kinds in, a box for each block out. Any
// number of blocks is set out, from one to hundreds; nothing here refuses a count.
//
// Lengths are counted in slide widths, across and down alike, so that a gap is as wide as it is
// high; boxes are given back as fractions of the slide.
import type { FreeArrange, FreeBlockKind } from '../deck/spec.ts';
import type { SlideBox } from '../template/manifest.ts';

/** What the arrangement needs to know of the template. */
export interface LayoutFrame {
  /** Left edge and width of the room for content, as fractions of the slide's width. */
  x: number;
  w: number;
  /** Top and bottom of the room the template's own layouts use, as fractions of the slide's height. */
  top: number;
  bottom: number;
  /** Top and bottom of the most room a crowded slide may take. */
  denseTop: number;
  denseBottom: number;
  /** The space between neighbours, as a fraction of the slide's width. */
  gap: number;
  /** The slide's width over its height. */
  aspect: number;
  /**
   * The shape of a card as the template draws one: its height over its width, both in slide
   * widths. A card that has room to grow is never made higher than this times its width. Left
   * out, a card is at most as high as it is wide.
   */
  card?: number;
}

export interface LaidBlock {
  /** The block's number in the slide, from 1. */
  index: number;
  kind: FreeBlockKind;
  box: SlideBox;
  /** It is the main block, beside or above the others. */
  lead?: boolean;
  /**
   * A card or a step that took room that was left over: by how much its box is higher than its
   * content asked, as a fraction of the slide's height. What is written on it is set out as in
   * the box without that, at its top, so that a card that grew holds what it held.
   */
  grown?: number;
}

export interface FreeLayout {
  blocks: LaidBlock[];
  /** Every block is drawn as a tile on a panel: the slide mixes kinds that do not line up in bands. */
  tiles: boolean;
  /** The slide took the larger room: some block would have been too small in the usual one. */
  dense: boolean;
  /** How many blocks the fullest line holds. */
  columns: number;
  /** The notes of the slide are set smaller to share their room: the factor on the size of their text. Left out when every note is as it asked to be. */
  lower?: number;
  /**
   * The slide asked for a number of blocks side by side that would leave some block under its
   * smallest size, and was set out with the nearest number that does not: what was asked for,
   * and what the slide holds. Left out when the wish was followed, or when no number helps.
   */
  relaxed?: { asked: number; columns: number };
}

interface Shape {
  /** The cell shapes (width over height) a block of the kind sits well in. */
  lo: number;
  hi: number;
  /** The least comfortable width and height, in slide widths. */
  minW: number;
  minH: number;
  /** The most a block is high, as a factor of its width. */
  maxH?: number;
  /** Its share of the height when kinds are stacked in bands. */
  share: number;
}

const PICTURE: Shape = { lo: 0.75, hi: 2.2, minW: 0.09, minH: 0.09, maxH: 1 / 0.75, share: 1 };
const KINDS: Record<FreeBlockKind, Shape> = {
  text: { lo: 0.9, hi: 4, minW: 0.2, minH: 0.075, share: 0.6 },
  list: { lo: 0.7, hi: 3, minW: 0.18, minH: 0.09, share: 0.9 },
  quote: { lo: 1.5, hi: 6, minW: 0.3, minH: 0.075, share: 0.6 },
  card: { lo: 0.6, hi: 4.5, minW: 0.15, minH: 0.09, maxH: 1.6, share: 0.8 },
  step: { lo: 0.6, hi: 4.5, minW: 0.15, minH: 0.09, maxH: 1.6, share: 0.8 },
  figure: { lo: 0.9, hi: 4, minW: 0.11, minH: 0.06, maxH: 0.55, share: 0.45 },
  // A phase is an arrow with a few lines under it, a milestone a date over a few lines: both are read along their line, so they are low and may be narrow.
  phase: { lo: 0.9, hi: 5, minW: 0.12, minH: 0.07, maxH: 1.2, share: 0.7 },
  milestone: { lo: 0.6, hi: 4, minW: 0.12, minH: 0.1, maxH: 1.4, share: 0.8 },
  // An option is a column of points: higher than wide where it has several.
  option: { lo: 0.4, hi: 2.4, minW: 0.17, minH: 0.14, maxH: 2.4, share: 1 },
  // A chart sits best in a cell wider than high (`lo`), and may be a little higher than wide where a line of charts has the room: it then ends where the room ends.
  chart: { lo: 1.1, hi: 2.6, minW: 0.105, minH: 0.105, maxH: 1 / 0.85, share: 1 },
  table: { lo: 1.2, hi: 5, minW: 0.22, minH: 0.09, share: 1 },
  photo: PICTURE,
  graphic: PICTURE,
  note: { lo: 4, hi: 40, minW: 0.3, minH: 0.03, share: 0.2 },
};

const PROSE = new Set<FreeBlockKind>(['text', 'list', 'quote']);
const VISUAL = new Set<FreeBlockKind>(['chart', 'table', 'photo', 'graphic']);
/** The most of the room that notes take together when the slide holds anything else. */
const NOTES = 0.4;
/** The least share of its size that the text of a note is asked to be set at. Notes that do not fit even then are made lower still, as they are. */
const NOTE_LEAST = 0.25;
/** Kinds whose blocks are stretched to the room there is; the others are only as high as they need to be. */
const FILLS = new Set<FreeBlockKind>(['chart', 'photo', 'graphic']);
/**
 * Blocks that are lower than their room sit down it by this share of what they leave free, when
 * they leave SETTLE_FULL of the room or more. Blocks that leave less sit down by less, evenly, and
 * blocks that leave SETTLE_NONE of the room or less start at its top: there is no height at which a
 * slide jumps from one to the other.
 */
const SETTLE = 0.4;
const SETTLE_FULL = 0.3;
const SETTLE_NONE = 0.1;
/**
 * Cards and steps that are lower than their room grow before anything is moved: a line of them
 * by up to this share of its height, or by less when the cards lie (they are 1.6 times as wide as
 * high, or more), and never higher than the template draws a card of that width. A row of cards
 * then starts where the template's own rows start, and what is left under it is little.
 */
const GROW = 0.35;
const GROW_LYING = 0.15;
/** Up to this many cards or steps never stand more than four on a line when nothing was asked for: a fifth makes every text small. */
const FEW = 8;
const FEW_ON_A_LINE = 4;
/** Kinds that are drawn in a box of their own. A short last line of them is spread over the width when it holds at least half as many as the lines above, and stands in the middle otherwise. */
const BOXED = new Set<FreeBlockKind>(['card', 'figure', 'step', 'phase', 'option']);
/**
 * Phases and milestones are read along one line, and options side by side: up to so many of one
 * of these kinds stand on one line when nothing was asked for, whatever shape that gives each,
 * as long as each still has the least width its kind sits well in. More of them, or narrower
 * ones, are set out in lines as other blocks are.
 */
const ONE_LINE: Partial<Record<FreeBlockKind, number>> = { phase: 6, milestone: 6, option: 4 };
/**
 * A card or a step grows into room that is left over only as far as what is written on it, from
 * the card's top to the foot of its text, still fills this share of it: a card twice as high as
 * its text reads as unfinished.
 */
const FULL_ENOUGH = 0.6;
/**
 * In the larger room, which only a crowded slide takes, the space between two lines of blocks is
 * no more than this share of a line's height, and a note stands this share of the usual gap from
 * its neighbour: on such a slide the height belongs to the blocks.
 */
const SNUG_ROWS = 0.22;
const SNUG_NOTE = 0.6;
/**
 * The least cell a chart is drawn in with names and numbers, in slide widths: narrower or lower, the
 * text it carries (4.2% of its width, a sixth and a half of its height, under a caption) would be
 * under 7 pt on a slide 13.33 inches wide, and the chart is drawn bare (plan.ts). Of arrangements
 * that suit alike, one that keeps every chart readable wins.
 */
const CHART_READS = { w: 0.174, h: 0.075 };
/**
 * A chart lower than this, in slide widths, is drawn small, with less of its text and a plot of a
 * few steps: a slide of such charts takes the larger room where that makes them higher, so that
 * eight charts over a bottom line are set out as ten are.
 */
const CHART_LOW = 0.13;
/** A card on a tile that needs this many times its height for what is written on it makes a slide crowded. */
const SHORT_TILE = 1.25;
/** More charts than this on a slide are drawn bare whatever their size (plan.ts). */
const MOST_READ = 24;
/** Kinds whose short last line shares the whole width, however few they are: text has no box whose edge would have to line up. */
const FLOWS = new Set<FreeBlockKind>(['text', 'list', 'quote', 'table']);

/** How far down its room a group of blocks sits that takes `used` of a room `height` high. */
export function settle(used: number, height: number): number {
  const free = height - used;
  if (!(free > 0) || !(height > 0)) return 0;
  const share = Math.min(1, Math.max(0, (free / height - SETTLE_NONE) / (SETTLE_FULL - SETTLE_NONE)));
  return SETTLE * free * share;
}

interface Box {
  x: number;
  y: number;
  w: number;
  h: number;
}
interface Block {
  index: number;
  kind: FreeBlockKind;
}
interface Placed extends Block {
  box: Box;
  lead?: boolean;
  /** How much higher the block was made than it asked to be, in slide widths. */
  grown?: number;
}

/**
 * The smallest size of a block: the least comfortable width of its kind, and this share of its
 * least comfortable height. A wish for blocks side by side that leaves a block under it is not
 * followed.
 *
 * The two are not alike because text is not. The size of a text follows the width of its block,
 * so a block narrower than its kind sits well in has text nobody reads: twelve cards on one line
 * have theirs at about 5 pt. A block that is low keeps the width of its lines, and its text is set
 * smaller only as far as its height asks: seven cards one under the other, at 52% of the least
 * comfortable height, hold their heading and text at 12 pt. Under about half there is no room for
 * a block's own parts, a heading and a line of text or a caption and a chart: measured on two
 * templates, shapes of neighbouring blocks began to lie over each other from 47% down.
 */
const SMALLEST_HEIGHT = 0.5;

/** `count` blocks on `rows` lines, spread evenly, earlier lines one longer: five make three and two. */
export function spread(count: number, rows: number): number[] {
  const out: number[] = [];
  for (let row = 0, left = count; row < rows; row++) {
    const here = Math.ceil(left / (rows - row));
    out.push(here);
    left -= here;
  }
  return out;
}

/** The space between cells: the template's own up to four on a line and three lines, less beyond, never under 0.3% of the slide's width. */
const gapFor = (gap: number, columns: number, rows: number): number => Math.min(gap, Math.max(0.003, gap * Math.min(1, 4 / columns, 3 / rows)));

/**
 * The most a block of a kind is high, as a factor of its width. Beside a main block a figure may be
 * as high as it is wide: the figures there share the height of the main block. Figures that stand
 * by themselves may be three quarters as high as wide: three of them on a line are not lower than
 * four in two lines, so that fewer figures are never set smaller than more in the same room.
 */
const highest = (kind: FreeBlockKind, beside?: boolean, figures?: boolean): number => (beside && kind === 'figure' ? 1 : figures && kind === 'figure' ? FIGURES_HIGH : (KINDS[kind].maxH ?? Infinity));
const FIGURES_HIGH = 0.75;

/**
 * How many lines suit some blocks in a box: the count whose cells are nearest the shapes the
 * kinds sit well in. `beside` is for the blocks beside a main block: an arrangement whose last
 * line is full, with every cell at least as wide and high as its kind sits well in, is taken
 * before any other, so that three small blocks stand as a column and not as two and one.
 */
function choose(kinds: FreeBlockKind[], box: Box, gap: number, beside?: boolean): { rows: number; columns: number; cw: number; ch: number } {
  const count = kinds.length;
  let best = { rows: 1, columns: count, cw: 0, ch: 0 };
  let least = Infinity;
  let even: typeof best | undefined;
  let leastEven = Infinity;
  const few = count <= FEW && kinds.every((kind) => kind === 'card' || kind === 'step');
  const figures = kinds.every((kind) => kind === 'figure');
  const only = kinds[0] !== undefined && kinds.every((kind) => kind === kinds[0]) ? kinds[0] : undefined;
  const oneLine = only !== undefined && count <= (ONE_LINE[only] ?? 0) && (box.w - (count - 1) * gapFor(gap, count, 1)) / count >= KINDS[only].minW - 1e-9;
  for (let rows = 1; rows <= count; rows++) {
    const split = spread(count, rows);
    const columns = split[0] as number;
    if (rows > 1 && Math.ceil(count / (rows - 1)) === columns) continue;
    if (few && columns > FEW_ON_A_LINE) continue;
    if (oneLine && rows > 1) continue;
    const between = gapFor(gap, columns, rows);
    const cw = (box.w - (columns - 1) * between) / columns;
    const ch = (box.h - (rows - 1) * between) / rows;
    if (cw <= 0 || ch <= 0) continue;
    const shape = cw / ch;
    // A short last line costs; so does a cell that is the wrong shape or too small for what it holds.
    let cost = 0.5 * (1 - (split[split.length - 1] as number) / columns);
    for (const kind of kinds) {
      const wanted = KINDS[kind];
      cost += (Math.max(0, Math.log(wanted.lo / shape), Math.log(shape / wanted.hi)) + 0.5 * (Math.max(0, Math.log(wanted.minW / cw)) + Math.max(0, Math.log(wanted.minH / ch)))) / count;
    }
    // So does a chart in a cell too small for its names and numbers: it would be drawn bare.
    if (kinds.includes('chart') && (cw < CHART_READS.w - 1e-9 || ch < CHART_READS.h - 1e-9)) cost += 0.5;
    // So does room that stays empty: lines that cannot be as high as their cells leave the
    // box open below. Of two arrangements that suit the blocks alike, the one that uses the
    // height wins; up to 15% left over costs nothing.
    const high = Math.min(ch, Math.max(...kinds.map((kind) => highest(kind, beside, figures) * cw)));
    cost += 0.1 * Math.max(0, 0.85 - (rows * high + (rows - 1) * between) / box.h);
    if (cost < least - 1e-9) {
      least = cost;
      best = { rows, columns, cw, ch };
    }
    const whole = split[split.length - 1] === columns && kinds.every((kind) => cw >= KINDS[kind].minW - 1e-9 && ch >= KINDS[kind].minH - 1e-9);
    if (beside && whole && cost < leastEven - 1e-9) {
      leastEven = cost;
      even = { rows, columns, cw, ch };
    }
  }
  return even ?? best;
}

/** How high a block needs to be at a width, in slide widths; `undefined` when it takes the room there is. */
type Needed = ((index: number, kind: FreeBlockKind, width: number) => number | undefined) | undefined;

/** How a grid is made: so many on a line when that was asked for, how high its blocks need to be, whether it stands beside a main block, whether its blocks are tiles. */
interface GridOptions {
  columns?: number | undefined;
  needed?: Needed;
  beside?: boolean | undefined;
  tiles?: boolean;
  /** The slide is crowded: the lines of blocks stand closer together (`SNUG_ROWS`). */
  snug?: boolean | undefined;
  /** How high a text, a list or a quotation needs to be at a width: asked where kinds stand in bands, so that a short text does not keep a share of the height it does not use. */
  prose?: Needed;
}

/**
 * Where a last line with fewer blocks than the lines above stands. Tiles, running text, lists,
 * quotations and tables share the whole width. Cards, figures and steps do when the line holds at
 * least half as many as a full one, and stand in the middle when it holds fewer. Charts and
 * pictures, which are read as a raster, keep the width of the lines above and start where they
 * start, so that the fifth stands under the first.
 */
function shortLine(kinds: FreeBlockKind[], columns: number, tiles: boolean): 'spread' | 'middle' | 'start' {
  if (tiles || kinds.every((kind) => FLOWS.has(kind))) return 'spread';
  if (kinds.every((kind) => BOXED.has(kind))) return kinds.length * 2 >= columns ? 'spread' : 'middle';
  return 'start';
}

function grid(blocks: Block[], box: Box, gap: number, options: GridOptions = {}): Placed[] {
  const { columns, needed, beside } = options;
  const count = blocks.length;
  if (count === 0) return [];
  const picked = columns ? { rows: Math.ceil(count / Math.min(columns, count)), columns: Math.min(columns, count) } : choose(blocks.map((block) => block.kind), box, gap, beside);
  const split = columns ? Array.from({ length: picked.rows }, (_, row) => Math.min(picked.columns, count - row * picked.columns)) : spread(count, picked.rows);
  const between = gapFor(gap, picked.columns, picked.rows);
  // The gaps never take more than half the room, across or down: however many blocks are asked
  // for on one line, each keeps a width and a height.
  const across = picked.columns > 1 ? Math.min(between, (0.5 * box.w) / (picked.columns - 1)) : between;
  let down = picked.rows > 1 ? Math.min(between, (0.5 * box.h) / (picked.rows - 1)) : between;
  // On a crowded slide the lines stand no further apart than a share of their own height.
  if (options.snug && picked.rows > 1) down = Math.min(down, (SNUG_ROWS * box.h) / (picked.rows + SNUG_ROWS * (picked.rows - 1)));
  const cw = (box.w - (picked.columns - 1) * across) / picked.columns;
  const room = (box.h - (picked.rows - 1) * down) / picked.rows;
  // A line is no higher than its highest block wants to be: three figures do not become three
  // towers, and a card is as high as what is written on it.
  const figures = blocks.every((block) => block.kind === 'figure');
  const wants = (block: Block) => Math.min(highest(block.kind, beside, figures) * cw, Math.max(KINDS[block.kind].minH, needed?.(block.index, block.kind, cw) ?? Infinity));
  const ch = Math.min(room, Math.max(...blocks.map(wants)));
  const out: Placed[] = [];
  let next = 0;
  split.forEach((inRow, row) => {
    const how = inRow < picked.columns ? shortLine(blocks.slice(next, next + inRow).map((block) => block.kind), picked.columns, options.tiles === true) : 'start';
    const w = how === 'spread' ? (box.w - (inRow - 1) * across) / inRow : cw;
    const rowWidth = inRow * w + (inRow - 1) * across;
    const x = how === 'middle' ? box.x + (box.w - rowWidth) / 2 : box.x;
    for (let column = 0; column < inRow; column++, next++) {
      const block = blocks[next] as Block;
      // Tiles of one line are equally high, whatever they hold: a chart beside a text ends where the text ends.
      out.push({ ...block, box: { x: x + column * (w + across), y: box.y + row * (ch + down), w, h: needed || options.tiles ? ch : Math.min(ch, highest(block.kind, beside, figures) * w) } });
    }
  });
  return out;
}

/** Runs of one kind in reading order; running text, lists and quotations count as one kind. */
function runsOf(blocks: Block[]): Block[][] {
  const runs: Block[][] = [];
  for (const block of blocks) {
    const last = runs[runs.length - 1];
    const head = last?.[0];
    if (last && head && (head.kind === block.kind || (PROSE.has(head.kind) && PROSE.has(block.kind)))) last.push(block);
    else runs.push([block]);
  }
  return runs;
}

/** Blocks of several kinds: each kind a band of its own across the slide, or, where that does not fit, one grid of tiles. */
function group(blocks: Block[], box: Box, gap: number, options: GridOptions = {}): { placed: Placed[]; tiles: boolean } {
  const { columns, needed, beside, snug, prose } = options;
  const runs = runsOf(blocks);
  if (columns || runs.length <= 1) return { placed: grid(blocks, box, gap, options), tiles: false };
  if (runs.length <= 3) {
    const total = runs.reduce((sum, run) => sum + KINDS[(run[0] as Block).kind].share, 0);
    const free = box.h - (runs.length - 1) * gap;
    /** A text, a list or a quotation is asked how high it needs to be; the other kinds as on any slide. */
    const ask: Needed = (index, kind, width) => (PROSE.has(kind) ? prose?.(index, kind, width) : needed?.(index, kind, width));
    /** The bands with these heights: what is placed, whether every band suits its blocks, and what each band really takes. */
    const lay = (heights: number[]): { placed: Placed[]; fits: boolean; used: number[] } => {
      let y = box.y;
      let fits = true;
      const placed: Placed[] = [];
      const used: number[] = [];
      runs.forEach((run, at) => {
        const kind = (run[0] as Block).kind;
        const height = heights[at] as number;
        const band = { x: box.x, y, w: box.w, h: height };
        const pick = choose(run.map((block) => block.kind), band, gap, beside);
        if (pick.ch < KINDS[kind].minH || pick.cw < KINDS[kind].minW || (FILLS.has(kind) && pick.cw / pick.ch > KINDS[kind].hi * 1.3)) fits = false;
        const cells = grid(run, band, gap, { needed: PROSE.has(kind) && prose ? ask : needed, beside, snug });
        placed.push(...cells);
        // The next band starts under what this one really takes, so that low figures do not leave a hole.
        const taken = cells.length > 0 ? Math.max(...cells.map((cell) => cell.box.y + cell.box.h)) - y : height;
        used.push(FILLS.has(kind) ? height : Math.min(height, taken));
        y += (used[at] as number) + gap;
      });
      return { placed, fits, used };
    };
    const heights = runs.map((run) => (free * KINDS[(run[0] as Block).kind].share) / total);
    let laid = lay(heights);
    // Whether the kinds suit bands is told by their shares: a band that then gives up what it has
    // no use for still holds what it holds, though it is lower than its kind sits well in.
    const suits = laid.fits;
    // What a band does not use of its share goes to the bands that take all of theirs: a text of
    // two lines over a row of cards leaves the cards the height it has no use for. Twice, because
    // a band that was given more may leave some of that again.
    for (let round = 0; round < 2 && suits; round++) {
      const spare = laid.used.reduce((sum, taken, at) => sum + Math.max(0, (heights[at] as number) - taken), 0);
      const hungry = laid.used.map((taken, at) => taken >= (heights[at] as number) - 1e-9);
      const wanting = heights.reduce((sum, height, at) => sum + (hungry[at] ? height : 0), 0);
      if (!(spare > 1e-6) || !(wanting > 0)) break;
      for (let at = 0; at < heights.length; at++) heights[at] = hungry[at] ? (heights[at] as number) + (spare * (heights[at] as number)) / wanting : (laid.used[at] as number);
      laid = lay(heights);
    }
    if (suits) return { placed: laid.placed, tiles: false };
  }
  return { placed: grid(blocks, box, gap, { tiles: true, snug }), tiles: true };
}

export interface ArrangeOptions {
  frame: LayoutFrame;
  arrange?: FreeArrange | undefined;
  /** The deck reads right to left: the first block of a line is at the right. */
  rtl: boolean;
  /**
   * How high a note is, in slide widths, for its number in the slide and the width it gets.
   * `lower` is 1, or less when the notes of the slide have to be set smaller to share their
   * room: the factor on the size of the note's text. Smaller text takes fewer lines, so a note
   * is asked again for its height at every size that is tried.
   */
  noteHeight?: (index: number, width: number, lower: number) => number;
  /** How high a block needs to be at a width, in slide widths: a card as high as its text. Not asked for blocks beside a lead, which share its height. */
  needed?: (index: number, kind: FreeBlockKind, width: number) => number | undefined;
  /**
   * How far down a card or a step what is written on it most likely reaches once it is set, from
   * the card's top, in slide widths: less than it needs, which is counted with room to spare. A
   * card grows into room that is left over only while this fills enough of it (`FULL_ENOUGH`).
   * Left out, a card grows by what it needs.
   */
  likely?: (index: number, kind: FreeBlockKind, width: number) => number | undefined;
  /** How high a text, a list or a quotation needs to be at a width, in slide widths, where it stands in a band over or under blocks of other kinds. */
  prose?: (index: number, kind: FreeBlockKind, width: number) => number | undefined;
  /**
   * What is written on the slide does not fit the usual room, though its blocks have their
   * sizes there: card text longer than its budget had to be drawn smaller to stay inside its
   * card. The slide counts as crowded, and takes the larger room where that gives some block
   * more room.
   */
  crowded?: boolean;
}

/** Sets out the blocks of a free slide. */
export function arrangeBlocks(kinds: readonly FreeBlockKind[], options: ArrangeOptions): FreeLayout {
  const { frame, rtl } = options;
  const wish = options.arrange ?? {};
  const asked = wish.columns !== undefined && Number.isInteger(wish.columns) && wish.columns > 0 ? wish.columns : undefined;
  // Left and right are sides of the slide; the arrangement is worked out as if read from the left.
  const leadSide: 'start' | 'end' | 'top' | 'none' | undefined = wish.lead === 'left' ? (rtl ? 'end' : 'start') : wish.lead === 'right' ? (rtl ? 'start' : 'end') : wish.lead;
  const gap = frame.gap;

  /**
   * The blocks in a room, `columns` side by side when that is given, with every note as high as
   * `heights` says and `noteGap` between a note and its neighbour. `grow` lets cards and steps
   * take room that is left over. `spare` is what the blocks leave of their room, in slide widths,
   * before anything grows. `snug` is for the larger room of a crowded slide: its lines of blocks
   * stand closer together.
   */
  const set = (room: Box, heights: number[], noteGap: number, columns: number | undefined, grow: boolean, snug: boolean): { notes: Placed[]; placed: Placed[]; tiles: boolean; tight: boolean; small: boolean; spare: number } => {
    let box = room;
    const notes: Placed[] = [];
    const rest: Block[] = [];
    const trailing: Array<{ block: Block; height: number }> = [];
    kinds.forEach((kind, position) => {
      const block = { index: position + 1, kind };
      if (kind !== 'note') return void rest.push(block);
      const height = heights[position] as number;
      if (rest.length === 0 && trailing.length === 0) {
        notes.push({ ...block, box: { x: box.x, y: box.y, w: box.w, h: height } });
        box = { ...box, y: box.y + height + noteGap, h: box.h - height - noteGap };
      } else {
        trailing.push({ block, height });
      }
    });
    // Notes after the first block stand under the rest; their room is kept free below.
    const kept = trailing.reduce((sum, entry) => sum + entry.height + noteGap, 0);
    const body: Box = { ...box, h: Math.max(0.02, box.h - kept) };

    let placed: Placed[] = [];
    let tiles = false;
    const visual = rest.filter((block) => VISUAL.has(block.kind));
    const others = rest.length - 1;
    const automatic = leadSide === undefined && !columns && visual.length === 1 && others >= 1 && others <= 4;
    if (rest.length > 1 && leadSide === 'top') {
      const lead = rest[0] as Block;
      const height = (body.h - gap) * (VISUAL.has(lead.kind) ? 0.58 : 0.35);
      const under = group(rest.slice(1), { x: body.x, y: body.y + height + gap, w: body.w, h: body.h - height - gap }, gap, { columns, needed: options.needed, snug });
      // A chart or a picture across the whole slide would be a ribbon: it keeps a shape it reads in.
      const widest = FILLS.has(lead.kind) ? Math.min(body.w, height * KINDS[lead.kind].hi * 1.3) : body.w;
      placed = [{ ...lead, lead: true, box: { x: body.x + (body.w - widest) / 2, y: body.y, w: widest, h: height } }, ...under.placed];
      tiles = under.tiles;
    } else if (rest.length > 1 && (leadSide === 'start' || leadSide === 'end' || automatic)) {
      const lead = automatic ? (visual[0] as Block) : (rest[0] as Block);
      const atStart = automatic ? rest.indexOf(lead) === 0 : leadSide === 'start';
      const share = lead.kind === 'photo' || lead.kind === 'graphic' ? 0.43 : VISUAL.has(lead.kind) ? 0.57 : 0.5;
      const width = (body.w - gap) * share;
      const beside = group(rest.filter((block) => block !== lead), { x: atStart ? body.x + width + gap : body.x, y: body.y, w: body.w - width - gap, h: body.h }, gap, { columns, beside: true, snug });
      placed = [{ ...lead, lead: true, box: { x: atStart ? body.x : body.x + body.w - width, y: body.y, w: width, h: body.h } }, ...beside.placed];
      tiles = beside.tiles;
    } else {
      const grouped = group(rest, body, gap, { columns, needed: options.needed, snug, prose: options.prose });
      placed = grouped.placed;
      tiles = grouped.tiles;
    }
    const bottomOf = () => (placed.length > 0 ? Math.max(...placed.map((entry) => entry.box.y + entry.box.h)) - body.y : 0);
    const spare = placed.length > 0 ? Math.max(0, body.h - bottomOf()) : 0;
    // Cards and steps that are lower than their room grow first. Every line of them may grow by
    // the same share of what it is allowed, and what stands under a line that grew moves down
    // with it. Not beside or under a main block, where the blocks share its height already, and
    // not among tiles.
    if (grow && spare > 1e-9 && !tiles && !placed.some((entry) => entry.lead)) {
      const lines = new Map<number, Placed[]>();
      for (const entry of placed) lines.set(Math.round(entry.box.y * 1e6), [...(lines.get(Math.round(entry.box.y * 1e6)) ?? []), entry]);
      const sorted = [...lines.entries()].sort((a, b) => a[0] - b[0]).map(([, line]) => line);
      const mayGrow = (line: Placed[]): number => {
        if (!line.every((entry) => entry.kind === 'card' || entry.kind === 'step')) return 0;
        const high = Math.max(...line.map((entry) => entry.box.h));
        const narrowest = Math.min(...line.map((entry) => entry.box.w));
        const lying = narrowest >= 1.6 * high;
        // No further than what is written on the fullest card of the line still fills enough of it.
        const written = Math.max(...line.map((entry) => options.likely?.(entry.index, entry.kind, entry.box.w) ?? high));
        return Math.max(0, Math.min(high * (1 + (lying ? GROW_LYING : GROW)), (frame.card ?? 1) * narrowest, Math.max(high, written / FULL_ENOUGH)) - high);
      };
      const wishes = sorted.map(mayGrow);
      const wished = wishes.reduce((sum, wish) => sum + wish, 0);
      const share = wished > 0 ? Math.min(1, spare / wished) : 0;
      let moved = 0;
      sorted.forEach((line, at) => {
        const more = (wishes[at] as number) * share;
        for (const entry of line) {
          entry.box = { ...entry.box, y: entry.box.y + moved, h: entry.box.h + more };
          if (more > 0) entry.grown = more;
        }
        moved += more;
      });
    }
    // Notes that are not first close the blocks over them: they stand one gap under the lowest
    // block, never under a hole that the blocks leave. What is still low, the blocks and the
    // notes that close them as one group, sits a little down the room, not pressed to its top.
    // A slide of nothing but running text, lists, quotations and tables hangs from the top line,
    // as the template's own text layouts do. Blocks that fill their room have their notes at its foot.
    const used = bottomOf();
    const down = placed.length > 0 && !placed.every((entry) => PROSE.has(entry.kind) || entry.kind === 'table') ? settle(used + kept, box.h) : 0;
    if (down > 0) for (const entry of placed) entry.box = { ...entry.box, y: entry.box.y + down };
    let y = body.y + down + (placed.length > 0 ? used : body.h) + noteGap;
    for (const { block, height } of trailing) {
      notes.push({ ...block, box: { x: box.x, y, w: box.w, h: height } });
      y += height + noteGap;
    }

    /** Whether some block is under the least comfortable width of its kind, or under a share of its least comfortable height. */
    const under = (share: number): boolean =>
      placed.some((entry) => {
        const shape = KINDS[entry.kind];
        const least = shape.maxH === undefined ? shape.minH : Math.min(shape.minH, shape.maxH * entry.box.w);
        return entry.box.w < shape.minW - 1e-6 || entry.box.h < share * least - 1e-6;
      });
    // Charts in cells too small for their names and numbers make a slide crowded too, while they
    // are few enough to be read at all, and so do charts that are drawn small (`CHART_LOW`): the
    // larger room is tried for them.
    const charts = placed.filter((entry) => entry.kind === 'chart');
    const unread = charts.length <= MOST_READ && charts.some((entry) => entry.box.w < CHART_READS.w - 1e-9 || entry.box.h < CHART_READS.h - 1e-9 || (!tiles && !entry.lead && entry.box.h < CHART_LOW - 1e-9));
    // So do tiles of cards or steps that what is written on them needs a quarter more height than
    // they have: a slide of mixed tiles shares one height among all its blocks, and gives its
    // cards no more than that.
    const short = tiles && placed.some((entry) => (entry.kind === 'card' || entry.kind === 'step') && (options.needed?.(entry.index, entry.kind, entry.box.w) ?? 0) > SHORT_TILE * entry.box.h + 1e-6);
    return { notes, placed, tiles, tight: under(1) || unread || short, small: under(SMALLEST_HEIGHT), spare };
  };

  const attempt = (top: number, bottom: number, columns: number | undefined, snug = false): { placed: Placed[]; tiles: boolean; tight: boolean; small: boolean; lower: number } => {
    // Heights in slide widths.
    const room: Box = { x: frame.x, y: top / frame.aspect, w: frame.w, h: (bottom - top) / frame.aspect };
    // Notes are for the bottom line, not for the slide. When they ask for more than two fifths of
    // the room (NOTES), counting the gap that sets each apart from its neighbour, their text is
    // set smaller, by one factor for every note of the slide, and each is asked again how high it
    // is at that size: smaller text takes fewer lines. Space is given up before size: the gaps
    // come down by the square of that factor. So the size follows from how much the notes say,
    // and short notes are never set smaller than long ones would be in their place. A slide of
    // notes alone has the whole room for them.
    const noteCount = kinds.filter((kind) => kind === 'note').length;
    const alone = noteCount === kinds.length;
    const gaps = alone ? noteCount - 1 : noteCount;
    const gapAt = (lower: number): number => gap * (snug ? SNUG_NOTE : 1) * lower * lower;
    const heightsAt = (lower: number): number[] => kinds.map((kind, position) => (kind === 'note' ? (options.noteHeight?.(position + 1, room.w, lower) ?? 0.045 * lower) : 0));
    const taken = (heights: number[], lower: number): number => heights.reduce((sum, height) => sum + height, 0) + gaps * gapAt(lower);
    const full = heightsAt(1);
    const asked = taken(full, 1);
    /** The notes at the largest size at which they and their gaps take no more than `limit`. */
    const fitted = (limit: number): { heights: number[]; noteGap: number; lower: number } => {
      if (!(asked > limit + 1e-9)) return { heights: full, noteGap: gapAt(1), lower: 1 };
      let low = NOTE_LEAST;
      if (taken(heightsAt(low), low) <= limit) {
        let high = 1;
        for (let step = 0; step < 30; step++) {
          const middle = (low + high) / 2;
          if (taken(heightsAt(middle), middle) <= limit) low = middle;
          else high = middle;
        }
      }
      const heights = heightsAt(low);
      // Notes that ask for more than their share at the smallest size too are made lower still,
      // all by one factor, and their gaps with them: nothing leaves the room.
      const total = taken(heights, low);
      const squeeze = total > limit ? limit / total : 1;
      return { heights: heights.map((height) => height * squeeze), noteGap: gapAt(low) * squeeze, lower: low };
    };
    const most = alone ? room.h : NOTES * room.h;
    let fit = fitted(most);
    // Cards grow only on a slide whose notes are as they asked to be: room that is left over
    // goes to notes that were set smaller first.
    let made = set(room, fit.heights, fit.noteGap, columns, fit.lower >= 1, snug);
    // What the other blocks leave of their share is the notes' again: three low cards do not
    // keep three fifths of a slide empty. No block is made smaller for it. It is asked for again
    // until nothing more comes of it.
    for (let round = 0; round < 8 && fit.lower < 1 && !alone; round++) {
      const spare = made.spare;
      if (!(spare > 1e-6)) break;
      const more = fitted(fit.heights.reduce((sum, height) => sum + height, 0) + gaps * fit.noteGap + spare);
      const again = set(room, more.heights, more.noteGap, columns, false, snug);
      const before = new Map(made.placed.map((entry) => [entry.index, entry.box]));
      const kept = again.placed.every((entry) => {
        const was = before.get(entry.index);
        return was !== undefined && entry.box.w >= was.w - 1e-9 && entry.box.h >= was.h - 1e-9;
      });
      if (!kept || again.tiles !== made.tiles || again.tight !== made.tight || !(more.lower > fit.lower)) break;
      made = again;
      fit = more;
    }
    // A slide whose notes had to be set smaller is a crowded one: the larger room is tried.
    return { placed: [...made.notes, ...made.placed], tiles: made.tiles, tight: made.tight || fit.lower < 1, small: made.small, lower: fit.lower };
  };

  /**
   * The slide with so many blocks side by side: in the usual room, or in the larger one when a
   * block would be too small in the usual one and the larger one gives some block more room.
   * Blocks that are too narrow are no wider there: such a slide stays where it is, and is not
   * moved up under the title over an empty foot.
   */
  const lay = (columns: number | undefined): { result: ReturnType<typeof attempt>; dense: boolean } => {
    const usual = attempt(frame.top, frame.bottom, columns);
    if (!(usual.tight || options.crowded) || !(frame.denseTop < frame.top - 1e-6 || frame.denseBottom > frame.bottom + 1e-6)) return { result: usual, dense: false };
    const larger = attempt(frame.denseTop, frame.denseBottom, columns, true);
    const was = new Map(usual.placed.map((entry) => [entry.index, entry.box]));
    const gains = larger.tiles !== usual.tiles || larger.lower > usual.lower + 1e-9 || larger.placed.some((entry) => {
      const before = was.get(entry.index);
      return before === undefined || entry.box.w > before.w + 1e-6 || entry.box.h > before.h + 1e-6;
    });
    return gains ? { result: larger, dense: true } : { result: usual, dense: false };
  };
  let { result, dense } = lay(asked);
  let relaxed: FreeLayout['relaxed'];
  if (asked !== undefined && result.small) {
    // A wish the room cannot hold: some block is under its smallest size, also in the larger
    // room. The slide takes the nearest number of blocks side by side at which every block has
    // its smallest size, fewer before more. No block is narrower than the narrowest kind may
    // be, so no more than fit at that width are tried. When no number helps, the slide holds too
    // many blocks: it is made as asked.
    const blocks = kinds.filter((kind) => kind !== 'note').length;
    const most = Math.min(blocks, Math.floor((frame.w + gap) / Math.min(...Object.values(KINDS).map((shape) => shape.minW))));
    const now = Math.min(asked, blocks);
    const others = Array.from({ length: most }, (_, index) => index + 1)
      .filter((columns) => columns !== now)
      .sort((a, b) => Math.abs(a - asked) - Math.abs(b - asked) || a - b);
    for (const columns of others) {
      const other = lay(columns);
      if (other.result.small) continue;
      ({ result, dense } = other);
      relaxed = { asked, columns };
      break;
    }
  }
  const round = (value: number) => Math.round(value * 10000) / 10000;
  const blocks = result.placed
    .slice()
    .sort((a, b) => a.index - b.index)
    .map((entry): LaidBlock => {
      const x = rtl ? 2 * frame.x + frame.w - entry.box.x - entry.box.w : entry.box.x;
      const laid: LaidBlock = { index: entry.index, kind: entry.kind, box: { x: round(x), y: round(entry.box.y * frame.aspect), w: round(entry.box.w), h: round(entry.box.h * frame.aspect) } };
      if (entry.lead) laid.lead = true;
      if (entry.grown) laid.grown = round(entry.grown * frame.aspect);
      return laid;
    });
  const lines = new Map<number, number>();
  for (const block of blocks) if (block.kind !== 'note' && !block.lead) lines.set(block.box.y, (lines.get(block.box.y) ?? 0) + 1);
  const layout: FreeLayout = { blocks, tiles: result.tiles, dense, columns: Math.max(1, ...lines.values()) };
  if (result.lower < 1) layout.lower = result.lower;
  if (relaxed) layout.relaxed = relaxed;
  return layout;
}
