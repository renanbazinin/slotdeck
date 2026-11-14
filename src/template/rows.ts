// Rows and grids of like items: finding one from the names of a pattern's slots, and working
// out how many items a row has room for.
import type { ShapeInfo, SlideInfo } from './inspect.ts';
import type { RepeatDef, SlotDef } from './manifest.ts';
import { rowRoom } from './measure.ts';

/** Works out a row or grid of like items from numbered slot names such as card1Heading, card2Heading. */
export function detectRepeat(slots: SlotDef[], slide: SlideInfo): RepeatDef | undefined {
  const byId = new Map(slide.shapes.map((shape) => [shape.id, shape]));
  const items = new Map<string, Map<number, SlotDef[]>>();
  for (const slot of slots) {
    const match = /^([a-z][A-Za-z]*?)(\d+)([A-Z][A-Za-z0-9]*)?$/.exec(slot.name);
    // Decorations may be numbered too (sticker1, sticker2), but they are not items of a row.
    if (!match || (slot.kind === 'text' && slot.deckField) || slot.empty === 'drop' || slot.empty === 'keep') continue;
    const prefix = match[1] as string;
    const byNumber = items.get(prefix) ?? new Map<number, SlotDef[]>();
    byNumber.set(Number(match[2]), [...(byNumber.get(Number(match[2])) ?? []), slot]);
    items.set(prefix, byNumber);
  }

  for (const [prefix, byNumber] of items) {
    const count = byNumber.size;
    if (count < 2 || !Array.from({ length: count }, (_, index) => byNumber.has(index + 1)).every(Boolean)) continue;
    const suffixes = (n: number) => (byNumber.get(n) ?? []).map((slot) => slot.name.slice(prefix.length + String(n).length));
    const first = suffixes(1);
    if (!Array.from({ length: count }, (_, index) => suffixes(index + 1).slice().sort().join()).every((set) => set === first.slice().sort().join())) continue;

    // Each item moves as its outermost shapes: the group around a card, or the shapes themselves.
    const chain = (id: string): string[] => {
      const parent = byId.get(id)?.groupId;
      return parent ? [...chain(parent), id] : [id];
    };
    const chains = Array.from({ length: count }, (_, index) => (byNumber.get(index + 1) ?? []).map((slot) => chain(slot.shapeId)));
    // When every item sits inside one shared group, the items are that group's children.
    const shared = chains.flat().every((path) => path[0] === chains[0]?.[0]?.[0] && path.length > 1) ? 1 : 0;
    const shapes = chains.map((paths) => [...new Set(paths.map((path) => path[Math.min(shared, path.length - 1)] as string))]);
    if (new Set(shapes.flat()).size !== shapes.flat().length) continue; // items share a shape: not separable

    const boxOf = (ids: string[]) => {
      const boxes = ids.map((id) => (byId.get(id) as ShapeInfo).box);
      const x = Math.min(...boxes.map((box) => box.x));
      const y = Math.min(...boxes.map((box) => box.y));
      return { x, y, right: Math.max(...boxes.map((box) => box.x + box.w)), bottom: Math.max(...boxes.map((box) => box.y + box.h)) };
    };
    if (shared === 0) {
      // A card's background is often a separate shape behind its group: it belongs to the item too.
      const taken = new Set([...shapes.flat(), ...slots.map((slot) => slot.shapeId)]);
      for (const candidate of slide.shapes) {
        if (candidate.groupId || candidate.kind !== 'shape' || taken.has(candidate.id) || candidate.box.w > 0.6) continue;
        const holds = shapes.filter((ids) => {
          const box = boxOf(ids);
          const slack = 0.01;
          return candidate.box.x <= box.x + slack && candidate.box.y <= box.y + slack && candidate.box.x + candidate.box.w >= box.right - slack && candidate.box.y + candidate.box.h >= box.bottom - slack;
        });
        if (holds.length === 1) (holds[0] as string[]).unshift(candidate.id);
      }
    }

    const tops = shapes.map((ids) => boxOf(ids).y);
    const firstRow = tops.filter((top) => Math.abs(top - (tops[0] as number)) < 0.03).length;
    // Things stacked one under another (two paragraphs, say) are numbered alike but are not a row of items.
    if (firstRow < 2) continue;
    const repeat: RepeatDef = { item: prefix.toLowerCase(), slots: first.map((suffix) => `${prefix}{n}${suffix}`), shapes, min: Math.min(2, count), max: count };
    if (firstRow < count) repeat.columns = firstRow;
    return repeat;
  }
  return undefined;
}

/**
 * How many items of a width fit side by side between two limits, all as fractions of the slide
 * width. Items keep their width; what tightens is the space between them, down to 3% of an item.
 */
export const rowFit = (itemWidth: number, left: number, right: number): number => Math.max(1, Math.floor((right - left - itemWidth) / (itemWidth * 1.03) + 1 + 1e-6));

/** Whether a row's items sit directly on the slide. Only such a row can spread beyond where it is drawn. */
export const rowOnSlide = (slide: SlideInfo, repeat: RepeatDef): boolean => repeat.shapes.flat().every((id) => slide.shapes.find((shape) => shape.id === id)?.groupId === undefined);

/**
 * The most items a row can take on its slide: as many as fit side by side at their drawn width,
 * spreading out to the slide's margin or the next shape where the row is free to. Never fewer
 * than it draws. A grid takes what it draws, because one more would need another line.
 */
export function rowCapacity(slide: SlideInfo, repeat: RepeatDef): number {
  const drawn = repeat.shapes.length;
  if (repeat.columns !== undefined) return drawn;
  const room = rowRoom(slide, repeat.shapes);
  const span = repeat.span ?? { left: room.area.x, right: room.area.x + room.area.w };
  return Math.max(drawn, rowFit(room.itemWidth, span.left, span.right), rowOnSlide(slide, repeat) ? rowFit(room.itemWidth, room.left, room.right) : 0);
}

/** How far the space between a row's items may tighten before the row stops looking as designed: to a third. */
const TIGHTEST = 1 / 3;

/**
 * The most items a row takes while it still looks spaced as the template draws it: the space
 * between items no less than a third of the drawn one. This is what a draft settles on by itself.
 * Never more than {@link rowCapacity}, which a person may still ask for, and never fewer than drawn.
 */
export function rowComfort(slide: SlideInfo, repeat: RepeatDef): { most: number; span?: { left: number; right: number } } {
  const drawn = repeat.shapes.length;
  if (repeat.columns !== undefined || drawn < 2) return { most: drawn };
  const room = rowRoom(slide, repeat.shapes);
  const current = repeat.span ?? { left: room.area.x, right: room.area.x + room.area.w };
  const drawnGap = (room.area.w - drawn * room.itemWidth) / (drawn - 1);
  const gap = Math.max(drawnGap * TIGHTEST, room.itemWidth * 0.03);
  const spaced = (left: number, right: number) => Math.floor((right - left + gap) / (room.itemWidth + gap) + 1e-6);
  const within = Math.max(drawn, spaced(current.left, current.right));
  if (!rowOnSlide(slide, repeat)) return { most: within };
  // The row may spread into the free room beside it. The count holds only with that span, so the span is part of the answer.
  const left = Math.max(0, Math.floor(Math.min(current.left, room.left) * 1000) / 1000);
  const right = Math.min(1, Math.ceil(Math.max(current.right, room.right) * 1000) / 1000);
  const spread = Math.min(rowCapacity(slide, repeat), spaced(left, right));
  return spread > within ? { most: spread, span: { left, right } } : { most: within };
}
