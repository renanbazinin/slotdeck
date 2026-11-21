// What a writer is told a template's slot of paragraphs holds, counted without fonts from the
// template's own spacing: the lines its box holds at the paragraph count and text size told, less
// the space PowerPoint draws around its paragraphs, each paragraph counted in whole lines. The
// deck's check is not this count: it stays at the stored budget (`textBudget`), so a writer inside
// what it is told is never sent back.
import { FREE_PATTERN } from '../deck/spec.ts';
import type { Template } from '../render/render.ts';
import { slotFrame } from '../text/measure.ts';
import type { ParagraphStyle, Spacing, TextFrame } from '../text/styles.ts';
import { scalesWithText, templateTextSetting, textBudget } from './manifest.ts';
import type { PatternDef, TextSlotDef } from './manifest.ts';

/** PowerPoint sets lines 1.2 times the font size apart, before the line spacing. */
const LINE = 1.2;
/** The width of a character as a share of its size where the template states none: the free slides' estimate. */
const WIDE = 0.55;
/** A line is counted this full: what is lost where lines break at words. The free slides' share. */
const FULL = 0.85;
/** A line that holds fewer characters than this at the template's stated width is counted with the estimate where that is narrower: a short line loses more of itself at a word. */
const SHORT_LINE = 20;

/** What the count says of one template slot at one text scale. */
export interface SlotCount {
  /** The slot's stored budget at that scale (`textBudget`): what the deck is checked against. */
  stored: { maxChars: number; maxParagraphs: number };
  /** For p = 1 to the stored paragraphs: the most characters any text of p paragraphs fits in, never over the stored budget. Never rises with p. */
  curve: readonly number[];
  /** The pair the writer is told (see {@link toldPair}). */
  maxChars: number;
  maxParagraphs: number;
  /** Whether a text of paragraphs of these lengths fits, each counted in its own whole lines, within the stored budget. */
  holds(each: readonly number[]): boolean;
}

/** The count of a slot of a pattern at a text scale (`budgetScale`), and for a row of so many items. `undefined` where the count does not speak for the slot: it is told its stored budget. */
export type ToldCounter = (pattern: PatternDef, slot: TextSlotDef, scale: number, items?: number) => SlotCount | undefined;

const spacing = (value: Spacing, size: number): number => ('points' in value ? value.points : value.percent * LINE * size);

interface Line {
  size: number;
  pitch: number;
  before: number;
  after: number;
  /** Characters a line of the paragraph is counted to hold. */
  across: number;
}

/** A paragraph's size as the renderer writes it (as the measurer's `sizeAt`): a deck's text size rescales only a run that states its size, or a slot that sets one. */
function sizeOf(frame: TextFrame, style: ParagraphStyle, slot: TextSlotDef, scale: number): number {
  const body = scalesWithText(slot) ? scale : 1;
  const stated = slot.fontSize ?? style.run.size;
  const factor = slot.fontSize !== undefined || style.run.ownSize ? body : 1;
  const size = factor !== 1 || slot.fontSize !== undefined ? Math.round((stated * 100 * factor) / 50) / 2 : stated;
  return size * frame.fontScale;
}

function lineOf(frame: TextFrame, slot: TextSlotDef, scale: number, index: number, width: number, characterWidth: number | undefined): Line {
  // New paragraphs take the sample's styles in order, the last one for the rest (`setShapeText`).
  const style = frame.paragraphs[Math.min(index, frame.paragraphs.length - 1)] as ParagraphStyle;
  const size = sizeOf(frame, style, slot, scale);
  const pitch = 'points' in style.lineSpacing ? style.lineSpacing.points : LINE * size * Math.max(0.1, style.lineSpacing.percent - frame.lineReduction);
  const room = width - style.marginStart - style.marginEnd - Math.max(0, style.indent);
  const set = Math.floor(room / ((characterWidth ?? WIDE) * size));
  // Never more than the stated width holds: a template wider than the estimate is counted at its own width.
  const usable = characterWidth !== undefined && set < SHORT_LINE ? Math.min(set, Math.floor(room / (WIDE * size))) : set;
  // The renderer writes a slot's gap as the space after each paragraph, and keeps the space before.
  return { size, pitch, before: spacing(style.before, size), after: slot.paragraphGap ?? spacing(style.after, size), across: FULL * Math.max(1, usable) };
}

/** What a text of `count` paragraphs is counted in: the lines the box holds, and the characters a line holds in the narrowest of them. */
function linesFor(frame: TextFrame, lines: readonly Line[], count: number): { lines: number; across: number; space: number; pitch: number } {
  let pitch = 0;
  let across = Number.POSITIVE_INFINITY;
  let gap = 0;
  let space = 0;
  for (let index = 0; index < count; index++) {
    const line = lines[index] as Line;
    pitch = Math.max(pitch, line.pitch);
    across = Math.min(across, line.across);
    gap = Math.max(gap, line.before + line.after);
    // PowerPoint draws the first paragraph's space before and the last one's space after only with `spcFirstLastPara`.
    if (index > 0 || frame.firstLast) space += line.before;
    if (index < count - 1 || frame.firstLast) space += line.after;
  }
  const height = frame.height - frame.insets.top - frame.insets.bottom;
  // In columns, each column after the first may lose a line and a paragraph's spacing where the text breaks into it.
  const room = frame.columns * height - (frame.columns - 1) * (pitch + gap) - space;
  return { lines: Math.max(0, Math.floor(room / pitch + 1e-9)), across, space, pitch };
}

/** The count of a slot, or `undefined` where it does not speak for it. */
function countSlot(frame: TextFrame, slot: TextSlotDef, scale: number, characterWidth: number | undefined): SlotCount | undefined {
  if ((slot.maxParagraphs ?? 1) < 2 || slot.maxChars === undefined || frame.vertical || !frame.wrap || frame.paragraphs.length === 0) return undefined;
  const budget = textBudget(slot, scale);
  const stored = { maxChars: budget.maxChars as number, maxParagraphs: budget.maxParagraphs ?? 1 };
  const width = (frame.width - frame.insets.left - frame.insets.right - (frame.columns - 1) * frame.columnGap) / frame.columns;
  const lines = Array.from({ length: stored.maxParagraphs }, (_, index) => lineOf(frame, slot, scale, index, width, characterWidth));
  // Only where the template's spacing matters at the paragraphs told: space drawn around them, or lines set further apart than 1.2 times their size.
  const told = linesFor(frame, lines, stored.maxParagraphs);
  if (told.space <= 0 && !lines.some((line) => line.pitch > LINE * line.size + 1e-9)) return undefined;
  const curve: number[] = [];
  for (let count = 1; count <= stored.maxParagraphs; count++) {
    const held = linesFor(frame, lines, count);
    // Each paragraph ends on a line of its own: c characters take fewer than c/across + 1 lines,
    // so p paragraphs of at most across x (lines - p + 1) characters take at most `lines`.
    curve.push(held.lines < count ? 0 : Math.min(stored.maxChars, Math.floor(held.across * (held.lines - count + 1))));
  }
  if (curve.every((chars) => chars <= 0)) return undefined;
  const holds = (each: readonly number[]): boolean => {
    const chars = each.reduce((sum, one) => sum + one, 0);
    if (each.length === 0 || each.length > stored.maxParagraphs || chars > stored.maxChars) return false;
    const needed = each.reduce((sum, one, index) => sum + Math.max(1, Math.ceil(one / (lines[index] as Line).across - 1e-9)), 0);
    return needed <= linesFor(frame, lines, each.length).lines;
  };
  const count: SlotCount = { stored, curve, maxChars: stored.maxChars, maxParagraphs: stored.maxParagraphs, holds };
  return { ...count, ...toldPair(count) };
}

/**
 * The pair a writer is told from a slot's count: of the paragraphs it may take (`paragraphs`, a
 * text's own count, or else its stored paragraphs), the number whose larger loss, of the stored
 * characters or of those paragraphs, is least, and what any text of that many holds. Of equal
 * losses, the one of more characters. Where the count holds the stored budget, the stored budget.
 */
export function toldPair(count: SlotCount, paragraphs?: number): { maxChars: number; maxParagraphs: number } {
  const { stored, curve } = count;
  const top = Math.max(1, Math.min(stored.maxParagraphs, paragraphs ?? stored.maxParagraphs));
  if ((curve[top - 1] as number) >= stored.maxChars) {
    let most = top;
    while (most < stored.maxParagraphs && (curve[most] as number) >= stored.maxChars) most++;
    return { maxChars: stored.maxChars, maxParagraphs: most };
  }
  const loss = (p: number): number => Math.max(1 - (curve[p - 1] as number) / stored.maxChars, 1 - p / top);
  let best = top;
  for (let p = 1; p <= top; p++) {
    const better = loss(p) < loss(best) - 1e-9 || (Math.abs(loss(p) - loss(best)) <= 1e-9 && (curve[p - 1] as number) > (curve[best - 1] as number));
    if (better) best = p;
  }
  const chars = Math.max(1, curve[best - 1] as number);
  // As many paragraphs as hold as much: a text of fewer paragraphs than the slot takes may be written in more, where that costs nothing.
  let most = best;
  while (most < stored.maxParagraphs && (curve[most] as number) >= chars) most++;
  return { maxChars: chars, maxParagraphs: most };
}

/** The fewest paragraphs the writer of an asked length is told: half those it is told otherwise, so that a list stays a list. */
const fewest = (pair: { maxParagraphs: number }): number => Math.max(1, Math.ceil(pair.maxParagraphs / 2));

/**
 * The pair the writer of an asked length is told, where its slot's share of the length is `need`
 * characters: the pair it is told otherwise ({@link toldPair}) where that holds the need, else the
 * most paragraphs that hold it, down to half the pair's, and what any text of that many holds. The
 * writer of a length chooses its paragraphs, so a text that holds in fewer, longer paragraphs is
 * told that and is not cut. Where none of those holds it, the pair it is told otherwise.
 */
export function toldFor(count: SlotCount, need: number): { maxChars: number; maxParagraphs: number } {
  const pair = toldPair(count);
  if (pair.maxChars >= need) return pair;
  for (let p = pair.maxParagraphs - 1; p >= fewest(pair); p--) {
    const chars = count.curve[p - 1] as number;
    if (chars >= need) return { maxChars: chars, maxParagraphs: p };
  }
  return pair;
}

/** The most characters a slot's count holds of an asked length ({@link toldFor}): in half the paragraphs it is told otherwise, or the pair. */
export function mostHeld(count: SlotCount): number {
  const pair = toldPair(count);
  return Math.max(pair.maxChars, count.curve[fewest(pair) - 1] ?? 0);
}

const counters = new WeakMap<Template, { manifest: Template['manifest']; counter: ToldCounter }>();

/**
 * The count of a template's slots of paragraphs, read from its package without fonts, with the
 * width its manifest states for a character (`characterWidth`) or the estimate. Kept per template
 * and manifest. A free slide's slots, and any slot it does not speak for, have none.
 */
export function toldCounter(template: Template): ToldCounter {
  const kept = counters.get(template);
  if (kept && kept.manifest === template.manifest) return kept.counter;
  const characterWidth = templateTextSetting(template.manifest, 'characterWidth');
  const known = new Map<string, SlotCount | undefined>();
  const counter: ToldCounter = (pattern, slot, scale, items) => {
    if (pattern.id === FREE_PATTERN || (slot as { styleOf?: unknown }).styleOf !== undefined || slot.deckField) return undefined;
    const key = JSON.stringify([pattern.id, slot.name, slot.shapeId, slot.maxChars, slot.maxParagraphs, slot.fontSize, slot.paragraphGap, slot.box, scale, items ?? null]);
    if (!known.has(key)) {
      let count: SlotCount | undefined;
      try {
        const frame = slotFrame(template.pkg, pattern, slot, items);
        count = frame ? countSlot(frame, slot, scale, characterWidth) : undefined;
      } catch {
        count = undefined; // a slot the package cannot frame is told its stored budget
      }
      known.set(key, count);
    }
    return known.get(key);
  };
  counters.set(template, { manifest: template.manifest, counter });
  return counter;
}
