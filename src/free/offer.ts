// What a planner needs to know of free slides before any content exists: whether a template can
// carry them, which kinds of block a deck may be offered, how many read well, and the pattern a
// planned slide is written against.
import type { DeckSpec, FreeArrange, FreeBlockKind, SlideSpec } from '../deck/spec.ts';
import { FREE_BLOCK_KINDS } from '../deck/spec.ts';
import type { Template } from '../render/render.ts';
import type { PatternDef, TemplateManifest } from '../template/manifest.ts';
import { freeSlotName, isFreeBlockKind } from './blocks.ts';
import { freeKitOrReason } from './kit.ts';
import { freeBodySize, readableSize } from './plan.ts';
import type { FreeContext } from './plan.ts';
import { freeFor } from './resolve.ts';

/** What a deck has to draw on, as far as free slides care. */
export interface FreeSupplies {
  /** The brief holds numbers: figures, charts and tables may be planned. */
  figures: boolean;
  /** Photos can be had: generated, or supplied. */
  photos: boolean;
  /** Pictures to show whole (diagrams, screenshots) were supplied. */
  graphics: boolean;
  /** An icon set was supplied: a card may carry one. */
  icons: boolean;
  /** Every slide is to be plain text. */
  textOnly?: boolean;
}

export interface FreeOffer {
  /** The kinds of block a slide may be planned with, in the order they are listed to a model. */
  kinds: FreeBlockKind[];
  /** About how many blocks a slide holds at the template's own text size. */
  comfortable: number;
  /** The id of the layout whose frame free slides stand on. */
  base: string;
}

/** One line on each kind of block, for a model that chooses them. */
export const FREE_KIND_NOTES: Record<FreeBlockKind, string> = {
  text: 'running text, with a heading when there are several',
  list: 'a list of short points',
  card: 'a heading and a short text in a box',
  figure: 'one big number and what it counts',
  step: 'a numbered step: a heading and a short text',
  quote: 'a quotation',
  note: 'one highlighted line across the slide, for the bottom line',
  chart: 'a chart of numbers the brief states',
  table: 'a table',
  photo: 'a photo',
  graphic: 'one of the supplied graphics, shown whole',
};

const PLAIN: readonly FreeBlockKind[] = ['text', 'list', 'quote', 'note'];

/**
 * What free slides a deck on this template may be offered, or `undefined` when it may be offered
 * none: the manifest turns them off, or the template has no layout to stand them on.
 */
export function freeOffer(template: Template, supplies: FreeSupplies): FreeOffer | undefined {
  if (template.manifest.free === false) return undefined;
  const kit = freeKitOrReason(template);
  if (typeof kit === 'string') return undefined;
  const kinds = FREE_BLOCK_KINDS.filter((kind) => {
    if (supplies.textOnly) return PLAIN.includes(kind);
    if (kind === 'figure' || kind === 'chart' || kind === 'table') return supplies.figures;
    if (kind === 'photo') return supplies.photos;
    if (kind === 'graphic') return supplies.graphics;
    return true;
  });
  return { kinds: [...kinds], comfortable: comfortableBlocks(template, supplies.icons), base: kit.base };
}

// Remembered for the manifest object it was worked out under, as a template's kit is.
const comfort = new WeakMap<Template, { manifest: TemplateManifest; known: Map<boolean, number> }>();

/**
 * How many cards a slide holds before their text is set smaller than the template's own: the
 * number a planner is told reads well. Cards are the measure because they are the block a slide
 * most often has several of. Each card is taken to say what a card usually says, as a slide
 * is planned, so the number is that of a slide as it will be written.
 */
function comfortableBlocks(template: Template, icons: boolean): number {
  let cache = comfort.get(template);
  if (!cache || cache.manifest !== template.manifest) comfort.set(template, (cache = { manifest: template.manifest, known: new Map() }));
  const { known } = cache;
  const before = known.get(icons);
  if (before !== undefined) return before;
  let most = 1;
  for (let count = 1; count <= 12; count++) {
    const resolved = freeFor(template, {}, { blocks: Array<FreeBlockKind>(count).fill('card'), slots: {} }, { icons });
    if (typeof resolved === 'string') break;
    const text = resolved.plan.pattern.slots.find((slot) => slot.name === freeSlotName(1, 'Text'));
    const size = text && text.kind === 'text' ? (text.fontSize ?? text.size ?? 0) : 0;
    // The text of one card or two may be set larger than the template's own; the measure is the template's own size.
    if (resolved.plan.layout.dense || size < 0.85 * freeBodySize(resolved.kit)) break;
    most = count;
  }
  known.set(icons, most);
  return most;
}

/** A free slide as a plan holds it, before its content is written. */
export interface PlannedFree {
  blocks: FreeBlockKind[];
  arrange?: FreeArrange | undefined;
  base?: string | undefined;
  /** Texts a person asked a length of, in characters, by slot name (`b1Text`). */
  lengths?: Readonly<Record<string, number>> | undefined;
  /** For some of those texts: how many paragraphs they will have, by slot name. Left out, a text is taken to have paragraphs of about 300 characters. */
  paragraphs?: Readonly<Record<string, number>> | undefined;
  /** For some of those texts on a slide set dense: the lengths of the paragraphs code sets as they are, by slot name (see `FreeContext.planning`). */
  asIs?: Readonly<Record<string, readonly number[]>> | undefined;
}

/**
 * The pattern a planned free slide is written against: its slots with the room each has, worked
 * out with no content. The slide is set out as it will be when every text is about as long as
 * texts of its kind usually are, and each slot says what its box holds then. Optional parts are
 * counted in by rule: an icon for every card when `icons` says an icon set is at hand, a heading
 * over each text of a slide of several blocks, a caption for each chart, table, photo or graphic
 * of a slide that has two or more of the kind, the source of a quotation. `withParts` names
 * slots to count in besides, such as `b2Caption`: a change to a deck passes the optional slots
 * that hold a value on the slide, so that a caption or a heading the slide has can be written
 * again. A text whose length a person asked for (`lengths`) is set out for that length instead
 * of the usual one, in the number of paragraphs `paragraphs` gives it where it gives one. A
 * sentence instead when the template cannot carry the slide.
 */
export function freePattern(template: Template, deck: Pick<DeckSpec, 'direction' | 'textScale'>, planned: PlannedFree, icons: boolean, withParts?: readonly string[]): { pattern: PatternDef; base: string } | string {
  const slide: Pick<SlideSpec, 'blocks' | 'arrange' | 'slots' | 'base'> = { blocks: planned.blocks, slots: {} };
  if (planned.arrange) slide.arrange = planned.arrange;
  if (planned.base !== undefined) slide.base = planned.base;
  const planning: NonNullable<FreeContext['planning']> = { icons };
  if (withParts && withParts.length > 0) planning.withParts = withParts;
  if (planned.lengths && Object.keys(planned.lengths).length > 0) planning.lengths = planned.lengths;
  if (planned.lengths && planned.paragraphs && Object.keys(planned.paragraphs).length > 0) planning.paragraphs = planned.paragraphs;
  if (planned.lengths && planned.asIs && Object.keys(planned.asIs).length > 0) planning.asIs = planned.asIs;
  const resolved = freeFor(template, deck, slide, planning);
  return typeof resolved === 'string' ? resolved : { pattern: resolved.plan.pattern, base: resolved.kit.base };
}

/** Like blocks that follow one another, as a plan names them: 6 charts, then 1 note. */
export interface BlockRun {
  kind: FreeBlockKind;
  number: number;
  /** The number of the run's first block in the slide, from 1. */
  first: number;
}

/** The blocks of a slide as runs of like blocks, in order. */
export function blockRuns(blocks: readonly FreeBlockKind[]): BlockRun[] {
  const runs: BlockRun[] = [];
  blocks.forEach((kind, index) => {
    const last = runs[runs.length - 1];
    if (last && last.kind === kind) last.number++;
    else runs.push({ kind, number: 1, first: index + 1 });
  });
  return runs;
}

/** The most blocks that runs are ever written out as, whatever their numbers say: more than any slide shows, and few enough that an answer with a number gone wrong cannot fill the memory. */
export const MOST_BLOCKS = 2000;

/**
 * Runs of like blocks written out again as the list of blocks: what a deck holds. How many
 * blocks a run has is its `count`, or, failing that, its `number`. Anything that is no kind, or
 * no count, is passed over, and the list is never longer than `MOST_BLOCKS`: what the runs name
 * beyond that is left out.
 */
export function blocksOfRuns(runs: ReadonlyArray<{ kind: unknown; count?: unknown; number?: unknown }>): FreeBlockKind[] {
  const blocks: FreeBlockKind[] = [];
  for (const run of runs) {
    const stated = Math.floor(Number(run.count));
    const count = stated >= 1 ? stated : Math.floor(Number(run.number));
    if (!isFreeBlockKind(run.kind) || !(count >= 1)) continue;
    for (let n = 0; n < count && blocks.length < MOST_BLOCKS; n++) blocks.push(run.kind);
    if (blocks.length >= MOST_BLOCKS) break;
  }
  return blocks;
}

/** What one free slide of running text holds, in characters. */
export interface FreeTextRoom {
  /** At the template's own text size (with the template's and the deck's factor). */
  comfortable: number;
  /** Set smaller where it must, down to four fifths of that size: what any free slide of one text holds. */
  dense: number;
  /** Set dense (`FreeArrange.dense`): down to the least size that reads, in two or three columns, in the most room the frame has. */
  densest: number;
  /** The sizes those are set at, in points, at their most. */
  sizes: { comfortable: number; dense: number; densest: number };
}

// Remembered for the manifest object it was worked out under, as the comfortable number of blocks is.
const textRooms = new WeakMap<Template, { manifest: TemplateManifest; known: Map<string, FreeTextRoom | undefined> }>();
/** The lengths `freeTextRoom` tries, in characters: from the shortest long text (a text of 600 characters or fewer is set larger, on a narrower line) up, in steps of 50. */
const ROOM_LEAST = 650;
const ROOM_MOST = 8000;
const ROOM_STEP = 50;

/**
 * How many characters a free slide of one running text holds on this template, for a deck in this
 * direction and at this text size: at the template's own size, at four fifths of it, and set dense.
 * `paragraphs` is the number of paragraphs the text has. Left out, it is taken to have paragraphs
 * of about 300 characters, and set dense, to be written as a writer told a length writes it: in
 * any number of paragraphs up to the 12 a text set dense may have. `densest` is then the most a
 * slide can be asked for and hold. More paragraphs hold less, and a few can take more lines than
 * many: each ends on a line of its own, with space under it. `dense` and `densest` are the longest
 * text for which that slide, and every shorter one, holds it; `comfortable` is the longest of
 * those that is also set at the template's own size. `undefined` when the template cannot carry a
 * free slide. Work of the engine alone, with no model call, remembered for the template, the
 * deck's direction and text size, and the number of paragraphs.
 */
export function freeTextRoom(template: Template, deck: Pick<DeckSpec, 'direction' | 'textScale'>, paragraphs?: number): FreeTextRoom | undefined {
  let cache = textRooms.get(template);
  if (!cache || cache.manifest !== template.manifest) textRooms.set(template, (cache = { manifest: template.manifest, known: new Map() }));
  const key = `${deck.direction ?? ''}|${deck.textScale ?? 1}|${paragraphs ?? ''}`;
  if (cache.known.has(key)) return cache.known.get(key);
  const kit = template.manifest.free === false ? undefined : freeKitOrReason(template);
  let found: FreeTextRoom | undefined;
  if (kit !== undefined && typeof kit !== 'string') {
    const body = freeBodySize(kit, deck.textScale ?? 1);
    const least = readableSize(kit.slideWidth);
    /** The budget of the text of a slide of so many characters, and the size it is set at. */
    const at = (chars: number, dense: boolean): { held: number; size: number } => {
      const count = paragraphs ?? Math.min(12, Math.max(1, Math.round(chars / 300)));
      const resolved = freeFor(template, deck, { blocks: ['text'], slots: {}, ...(dense ? { arrange: { dense: true } } : {}) }, { icons: false, lengths: { b1Text: chars }, ...(dense && paragraphs === undefined ? {} : { paragraphs: { b1Text: count } }) });
      const slot = typeof resolved === 'string' ? undefined : resolved.plan.pattern.slots.find((entry) => entry.name === freeSlotName(1, 'Text'));
      return slot?.kind === 'text' ? { held: slot.maxChars ?? 0, size: slot.fontSize ?? slot.size ?? 0 } : { held: 0, size: 0 };
    };
    const room = { comfortable: 0, dense: 0, densest: 0 };
    const sizes = { comfortable: body, dense: body, densest: body };
    const open = { dense: true, densest: true };
    for (let chars = ROOM_LEAST; chars <= ROOM_MOST && (open.dense || open.densest); chars += ROOM_STEP) {
      if (open.dense) {
        const plain = at(chars, false);
        if (plain.held >= chars) {
          [room.dense, sizes.dense] = [chars, plain.size];
          if (plain.size >= Math.floor(body * 2) / 2 - 0.01) [room.comfortable, sizes.comfortable] = [chars, plain.size];
        } else open.dense = false;
      }
      if (open.densest) {
        const dense = at(chars, true);
        if (dense.held >= chars && dense.size >= least - 1e-9) [room.densest, sizes.densest] = [chars, dense.size];
        else open.densest = false;
      }
    }
    found = { ...room, densest: Math.max(room.densest, room.dense), sizes };
  }
  cache.known.set(key, found);
  return found;
}
