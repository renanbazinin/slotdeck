// How wide text is set, counted by the kinds of its letters, without the fonts: for the checks a
// deck gets where no fonts measure it (production), and for the free slides' own counts.
import type { PptxPackage } from '../pptx/package.ts';
import { inspectTemplate } from '../template/inspect.ts';
import type { PatternDef, TextSlotDef } from '../template/manifest.ts';

/**
 * How wide a word is set, in sizes of its text: about half a size a letter, more for capitals and for
 * the widest letters, less for the narrowest. Generous, so that a word counted to fit a column does.
 */
export const wordWide = (word: string): number =>
  [...word].reduce((sum, char) => sum + (/[MWmw@%]/.test(char) ? 0.88 : /[A-Z]/.test(char) ? 0.68 : /[ijlt.,:;'!|I]/.test(char) ? 0.32 : /[fr()]/.test(char) ? 0.4 : 0.56), 0);

/**
 * How wide PowerPoint sets text of ordinary weight, as a share of `wordWide`: 0.77 to 0.81 of it
 * on average in its own rendering of two templates, one in Latin and one in Hebrew letters, and
 * no more than this on nineteen lines of twenty. A heavy face (a heading, the title of a table)
 * is set at about the whole of it.
 */
export const SET_SHARE = 0.85;

/** Where the lines of a text slot run: how wide a line of its box is and the size its text is set at, in points. */
export type LineRoom = (pattern: PatternDef, slot: TextSlotDef) => { width: number; size: number } | undefined;

/** The space PowerPoint keeps free at each side of a text box that states none: a tenth of an inch. */
const INSET = 7.2;
const rooms = new WeakMap<PptxPackage, LineRoom>();

/**
 * The line room of every text slot of a template, read from its slides once: the width of the
 * slot's box (its own `box` where the manifest moves it, else its shape's) less the space kept
 * free at its sides, and its size (`fontSize`, else `size`, else the largest size its shape
 * states). A box the software drew itself (a free slide's, with `drawn`) keeps no space free.
 * `undefined` for a slot whose box or size is not known.
 */
export function lineRoomOf(pkg: PptxPackage): LineRoom {
  const known = rooms.get(pkg);
  if (known) return known;
  let info: ReturnType<typeof inspectTemplate> | undefined;
  const room: LineRoom = (pattern, slot) => {
    info ??= inspectTemplate(pkg);
    const shape = info.slides[pattern.sourceSlide - 1]?.shapes.find((entry) => entry.id === slot.shapeId);
    const box = slot.box ?? shape?.box;
    const size = slot.fontSize ?? slot.size ?? shape?.fontSizes?.[0];
    if (!box || size === undefined || !(size > 0)) return undefined;
    const drawn = (slot as { drawn?: unknown }).drawn !== undefined;
    return { width: (box.w * info.slideWidth) / 12700 - (drawn ? 0 : 2 * INSET), size };
  };
  rooms.set(pkg, room);
  return room;
}

/**
 * The words of a text that no line of its box holds, as PowerPoint sets them: an address, a long
 * number, a word with no space to break at. PowerPoint breaks such a word inside it.
 */
export function wordsWiderThanLine(paragraphs: readonly string[], width: number, size: number): string[] {
  if (!(width > 0) || !(size > 0)) return [];
  const room = width / (SET_SHARE * size);
  const words = paragraphs.flatMap((paragraph) => paragraph.split(/\s+/)).filter((word) => word.length > 3 && wordWide(word) > room);
  return [...new Set(words)];
}
