// The blocks of a free slide: what each kind is made of, and what its content is called in the
// slide's `slots`. A free slide keeps its content as a row of items does: block 3 of kind `card`
// has `b3Heading` and `b3Text`. No template is needed to know that, so a deck can be checked
// and built from this file alone; where the blocks go is worked out elsewhere.
import { FREE_BLOCK_KINDS, FREE_PATTERN, isImageRef } from '../deck/spec.ts';
import type { ChartValue, DeckColors, FreeArrange, FreeBlockKind, ImageRef, SlideSpec, SlotValue, TableValue } from '../deck/spec.ts';
import type { SlotDef } from '../template/manifest.ts';

/** One part of a block. `name` is what follows `b3` in its slot's name. */
export interface FreePart {
  name: string;
  /** What the part holds: one text, several paragraphs, a pictogram, a picture that fills its frame, a picture shown whole, a chart, a table. */
  holds: 'text' | 'paragraphs' | 'icon' | 'photo' | 'graphic' | 'chart' | 'table';
  /** A block always has this part. An optional part is drawn when it has a value and is simply not there without one. */
  required: boolean;
}

const part = (name: string, holds: FreePart['holds'], required = true): FreePart => ({ name, holds, required });

/** The parts of each kind of block, in the order they are read. */
export const FREE_PARTS: Record<FreeBlockKind, readonly FreePart[]> = {
  text: [part('Heading', 'text', false), part('Text', 'paragraphs')],
  list: [part('Heading', 'text', false), part('Items', 'paragraphs')],
  card: [part('Icon', 'icon', false), part('Heading', 'text'), part('Text', 'text')],
  figure: [part('Value', 'text'), part('Label', 'text')],
  step: [part('Heading', 'text'), part('Text', 'text')],
  quote: [part('Quote', 'text'), part('Source', 'text', false)],
  note: [part('Text', 'text')],
  chart: [part('Caption', 'text', false), part('Chart', 'chart')],
  table: [part('Caption', 'text', false), part('Table', 'table')],
  photo: [part('Photo', 'photo'), part('Caption', 'text', false)],
  graphic: [part('Graphic', 'graphic'), part('Caption', 'text', false)],
};

export const isFreeBlockKind = (value: unknown): value is FreeBlockKind => typeof value === 'string' && (FREE_BLOCK_KINDS as readonly string[]).includes(value);

/** The name of the slot that holds a part of block `index` (from 1): `b3Heading`. */
export const freeSlotName = (index: number, partName: string): string => `b${index}${partName}`;

/** Every slot of a list of blocks: its name, the block it belongs to (from 1) and the part it is. */
export function freeSlots(blocks: readonly FreeBlockKind[]): Array<{ name: string; block: number; kind: FreeBlockKind; part: FreePart }> {
  return blocks.flatMap((kind, index) => (FREE_PARTS[kind] ?? []).map((entry) => ({ name: freeSlotName(index + 1, entry.name), block: index + 1, kind, part: entry })));
}

/**
 * The slots of a free slide as slot definitions without a shape: enough to check that every
 * value is of the right kind and that nothing required is missing, with no template at hand.
 * The title is always called `title` and may be left out.
 */
export function freeSlotDefs(blocks: readonly FreeBlockKind[]): SlotDef[] {
  const slots: SlotDef[] = [{ name: 'title', kind: 'text', shapeId: '', empty: 'remove' }];
  for (const { name, part: entry } of freeSlots(blocks)) {
    const base = { name, shapeId: '', ...(entry.required ? {} : { empty: 'remove' as const }) };
    if (entry.holds === 'chart') slots.push({ ...base, kind: 'chart' });
    else if (entry.holds === 'table') slots.push({ ...base, kind: 'table' });
    else if (entry.holds === 'icon' || entry.holds === 'photo' || entry.holds === 'graphic') slots.push({ ...base, kind: 'image', imageKind: entry.holds });
    else slots.push({ ...base, kind: 'text', ...(entry.holds === 'paragraphs' ? { maxParagraphs: 99 } : {}) });
  }
  return slots;
}

/** A sentence that says what a free slide holds: "3 charts and a note". For a plan shown to a person. */
export function describeBlocks(blocks: readonly FreeBlockKind[]): string {
  const counts = new Map<FreeBlockKind, number>();
  for (const kind of blocks) counts.set(kind, (counts.get(kind) ?? 0) + 1);
  const names = [...counts].map(([kind, count]) => (count === 1 ? `${/^[aeiou]/.test(kind) ? 'an' : 'a'} ${kind}` : `${count} ${kind}s`));
  return names.length <= 1 ? (names[0] ?? 'nothing') : `${names.slice(0, -1).join(', ')} and ${names[names.length - 1]}`;
}

/** A picture by its asset id, or as a full reference. */
type Picture = ImageRef | string;
const picture = (value: Picture): ImageRef => (typeof value === 'string' ? { asset: value } : value);

/** One block of a free slide with its content, as {@link freeSlide} takes it. */
export type FreeBlockContent =
  | { text: string | string[]; heading?: string }
  | { list: string[]; heading?: string }
  | { card: { heading: string; text: string; icon?: Picture } }
  | { figure: { value: string; label: string } }
  | { step: { heading: string; text: string } }
  | { quote: string; source?: string }
  | { note: string }
  | { chart: ChartValue; caption?: string }
  | { table: TableValue; caption?: string }
  | { photo: Picture; caption?: string }
  | { graphic: Picture; caption?: string };

/** A free slide with its content written block by block. */
export interface FreeSlideContent {
  title?: string;
  blocks: FreeBlockContent[];
  /** How many blocks stand side by side on one line. */
  columns?: number;
  lead?: FreeArrange['lead'];
  /** The layout whose frame the slide stands on. */
  base?: string;
  /** Colours for this slide alone. */
  colors?: DeckColors;
}

/**
 * Writes a free slide: a slide the software lays out itself from the blocks given, any number of
 * them, in reading order. It needs no template, and gives the slide as a deck holds it: the
 * kinds in `blocks`, the content in `slots` under `b1Text`, `b2Chart`, ...
 *
 * ```ts
 * freeSlide({ title: 'Sales by region', blocks: [{ chart, caption: 'North' }, { chart: other, caption: 'South' }, { note: 'Both grew.' }] })
 * ```
 */
export function freeSlide(content: FreeSlideContent): SlideSpec {
  const blocks: FreeBlockKind[] = [];
  const slots: Record<string, SlotValue> = {};
  if (content.title !== undefined && content.title.trim() !== '') slots.title = content.title;
  content.blocks.forEach((given, index) => {
    const n = index + 1;
    const set = (name: string, value: SlotValue | undefined) => {
      if (value !== undefined && value !== null && value !== '') slots[freeSlotName(n, name)] = value;
    };
    const block = given as Record<string, unknown>;
    if ('text' in given) {
      blocks.push('text');
      set('Heading', given.heading);
      set('Text', given.text);
    } else if ('list' in given) {
      blocks.push('list');
      set('Heading', given.heading);
      set('Items', given.list);
    } else if ('card' in given) {
      blocks.push('card');
      set('Icon', given.card.icon === undefined ? undefined : picture(given.card.icon));
      set('Heading', given.card.heading);
      set('Text', given.card.text);
    } else if ('figure' in given) {
      blocks.push('figure');
      set('Value', given.figure.value);
      set('Label', given.figure.label);
    } else if ('step' in given) {
      blocks.push('step');
      set('Heading', given.step.heading);
      set('Text', given.step.text);
    } else if ('quote' in given) {
      blocks.push('quote');
      set('Quote', given.quote);
      set('Source', given.source);
    } else if ('note' in given) {
      blocks.push('note');
      set('Text', given.note);
    } else if ('chart' in given) {
      blocks.push('chart');
      set('Caption', given.caption);
      set('Chart', given.chart);
    } else if ('table' in given) {
      blocks.push('table');
      set('Caption', given.caption);
      set('Table', given.table);
    } else if ('photo' in given) {
      blocks.push('photo');
      set('Photo', picture(given.photo));
      set('Caption', given.caption);
    } else if ('graphic' in given) {
      blocks.push('graphic');
      set('Graphic', picture(given.graphic));
      set('Caption', given.caption);
    } else {
      throw new Error(`block ${n} is none of the kinds a free slide has (${FREE_BLOCK_KINDS.join(', ')}): ${JSON.stringify(Object.keys(block))}`);
    }
  });
  const slide: SlideSpec = { pattern: FREE_PATTERN, blocks, slots };
  if (content.columns !== undefined || content.lead !== undefined) {
    slide.arrange = {};
    if (content.columns !== undefined) slide.arrange.columns = content.columns;
    if (content.lead !== undefined) slide.arrange.lead = content.lead;
  }
  if (content.base !== undefined) slide.base = content.base;
  if (content.colors !== undefined) slide.colors = content.colors;
  return slide;
}

/**
 * The content of a free slide block by block: what {@link freeSlide} would take to write it. For
 * code that adds, removes or reorders blocks: `freeSlide(change(freeContent(slide)))`, so that
 * nobody renumbers slot names by hand. A block whose required content is missing comes back
 * with empty values.
 */
export function freeContent(slide: SlideSpec): FreeSlideContent {
  const slots = slide.slots ?? {};
  const text = (value: SlotValue | undefined): string => (typeof value === 'string' ? value : Array.isArray(value) ? value.join('\n') : '');
  const lines = (value: SlotValue | undefined): string[] => (Array.isArray(value) ? value : typeof value === 'string' && value !== '' ? [value] : []);
  const image = (value: SlotValue | undefined): ImageRef => (isImageRef(value) ? value : { asset: '' });
  const optional = <T extends object>(block: T, name: string, value: SlotValue | undefined): T => (typeof value === 'string' && value !== '' ? { ...block, [name]: value } : block);
  const blocks = (slide.blocks ?? []).map((kind, index): FreeBlockContent => {
    const at = (name: string) => slots[freeSlotName(index + 1, name)];
    switch (kind) {
      case 'text': {
        const body = at('Text');
        return optional({ text: Array.isArray(body) ? body : text(body) }, 'heading', at('Heading'));
      }
      case 'list':
        return optional({ list: lines(at('Items')) }, 'heading', at('Heading'));
      case 'card': {
        const icon = at('Icon');
        const card: { heading: string; text: string; icon?: ImageRef } = { heading: text(at('Heading')), text: text(at('Text')) };
        if (isImageRef(icon)) card.icon = icon;
        return { card };
      }
      case 'figure':
        return { figure: { value: text(at('Value')), label: text(at('Label')) } };
      case 'step':
        return { step: { heading: text(at('Heading')), text: text(at('Text')) } };
      case 'quote':
        return optional({ quote: text(at('Quote')) }, 'source', at('Source'));
      case 'note':
        return { note: text(at('Text')) };
      case 'chart':
        return optional({ chart: at('Chart') as ChartValue }, 'caption', at('Caption'));
      case 'table':
        return optional({ table: at('Table') as TableValue }, 'caption', at('Caption'));
      case 'photo':
        return optional({ photo: image(at('Photo')) }, 'caption', at('Caption'));
      default:
        return optional({ graphic: image(at('Graphic')) }, 'caption', at('Caption'));
    }
  });
  const content: FreeSlideContent = { blocks };
  if (typeof slots.title === 'string' && slots.title !== '') content.title = slots.title;
  if (slide.arrange?.columns !== undefined) content.columns = slide.arrange.columns;
  if (slide.arrange?.lead !== undefined) content.lead = slide.arrange.lead;
  if (slide.base !== undefined) content.base = slide.base;
  if (slide.colors !== undefined) content.colors = slide.colors;
  return content;
}
