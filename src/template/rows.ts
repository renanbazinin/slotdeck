// Rows and grids of like items: finding one from the names of a pattern's slots, and working
// out how many items a row has room for.
import type { Box, ShapeInfo, SlideInfo } from './inspect.ts';
import type { PatternDef, RepeatDef, SlotDef } from './manifest.ts';
import { rowRoom } from './measure.ts';

/** A numbered slot name: what the item is called, its number, and what of the item it is (card, 2, Heading). */
const NUMBERED = /^([a-z][A-Za-z]*?)(\d+)([A-Z][A-Za-z0-9]*)?$/;

/** Works out a row or grid of like items from numbered slot names such as card1Heading, card2Heading. */
export function detectRepeat(slots: SlotDef[], slide: SlideInfo): RepeatDef | undefined {
  const byId = new Map(slide.shapes.map((shape) => [shape.id, shape]));
  const items = new Map<string, Map<number, SlotDef[]>>();
  for (const slot of slots) {
    const match = NUMBERED.exec(slot.name);
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
    // A slot that points at a shape the slide does not have is for checkManifest to report; no row is made of it.
    if ([...byNumber.values()].flat().some((slot) => !byId.has(slot.shapeId))) continue;
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

    // An item takes everything in it along: with fewer items its shapes are gone, and a further
    // item is a copy of the last. So a slot whose shape lies in an item is a slot of the row when
    // every item has one under its own number (card1Icon, card2Icon, or icon1, icon2): a pictogram
    // that is kept, a sticker that is removed, a picture named by its kind. Any other slot in an
    // item holds the row to no fewer items than that one, and in the last item it means no row.
    const own = new Set([...byNumber.values()].flat());
    const itemOf = (id: string) => shapes.findIndex((ids) => chain(id).some((part) => ids.includes(part))) + 1;
    const others = slots.filter((slot) => !own.has(slot) && itemOf(slot.shapeId) > 0);
    const nameOf = (slot: SlotDef) => {
      // A value of the whole deck is one text, whatever its name: it is never a slot of the row.
      const numbered = slot.kind === 'text' && slot.deckField ? null : NUMBERED.exec(slot.name);
      return numbered && Number(numbered[2]) === itemOf(slot.shapeId) ? `${numbered[1]}{n}${numbered[3] ?? ''}` : undefined;
    };
    const joins = (name: string | undefined): name is string => name !== undefined && others.filter((slot) => nameOf(slot) === name).length === count;
    const alone = others.filter((slot) => !joins(nameOf(slot))).map((slot) => itemOf(slot.shapeId));
    if (alone.includes(count)) continue;
    const names = [...first.map((suffix) => `${prefix}{n}${suffix}`), ...new Set(others.map(nameOf).filter(joins))];

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
    // A row with an arrow, a line or a highlight among its items holds what it draws: with fewer
    // items the rest are laid out again, across what stays where it is.
    const pinned = firstRow === count && rowRoom(slide, shapes, droppedShapes(slots)).pinned;
    const repeat: RepeatDef = { item: prefix.toLowerCase(), slots: names, shapes, min: pinned ? count : Math.max(Math.min(2, count), ...alone), max: count };
    if (firstRow < count) repeat.columns = firstRow;
    return repeat;
  }
  return undefined;
}

/**
 * The slots of a pattern that are not slots of its row though their shapes lie in one of the row's
 * items, each with the number of that item. On a slide with fewer items than that, the shape is
 * gone with its item and the slide cannot be rendered.
 */
export function strandedSlots(pattern: PatternDef, slide: SlideInfo | undefined): Array<{ slot: string; item: number }> {
  const repeat = pattern.repeat;
  if (!repeat || !slide) return [];
  const parent = new Map(slide.shapes.map((shape) => [shape.id, shape.groupId]));
  const names = new Set(repeat.slots.flatMap((name) => repeat.shapes.map((_, index) => name.replace('{n}', String(index + 1)))));
  const itemOf = (id: string | undefined): number => (id === undefined ? 0 : repeat.shapes.findIndex((ids) => ids.includes(id)) + 1 || itemOf(parent.get(id)));
  return pattern.slots.filter((slot) => !names.has(slot.name)).map((slot) => ({ slot: slot.name, item: itemOf(slot.shapeId) })).filter((entry) => entry.item > 0);
}

/** The shapes a pattern takes off every slide: they are not in the way of a row. */
export const droppedShapes = (slots: SlotDef[]): Set<string> => new Set(slots.filter((slot) => slot.empty === 'drop').map((slot) => slot.shapeId));

/**
 * The slots of a row's items that are shown in another place than the slide draws them, because
 * they carry a `box`. A row is laid out from where the slide draws its items, so such a row cannot
 * be given more items or wider ones.
 */
export function movedRowSlots(slots: SlotDef[], repeat: RepeatDef): string[] {
  const names = new Set(repeat.slots.flatMap((template) => repeat.shapes.map((_, index) => template.replace('{n}', String(index + 1)))));
  return slots.filter((slot) => names.has(slot.name) && 'box' in slot && slot.box !== undefined && slot.box !== null).map((slot) => slot.name);
}

/**
 * How many items of a width fit side by side between two limits, all as fractions of the slide
 * width. Items keep their width; what tightens is the space between them, down to 3% of an item.
 */
export const rowFit = (itemWidth: number, left: number, right: number): number => Math.max(1, Math.floor((right - left - itemWidth) / (itemWidth * 1.03) + 1 + 1e-6));

/** Whether a row's items sit directly on the slide. Only such a row can spread beyond where it is drawn. */
export const rowOnSlide = (slide: SlideInfo, repeat: RepeatDef): boolean => repeat.shapes.flat().every((id) => slide.shapes.find((shape) => shape.id === id)?.groupId === undefined);

/**
 * The limits a row is laid out between: its `span` when its items sit on the slide, otherwise the
 * row as drawn. The renderer uses a span only for a row on the slide, so counting must do the same.
 */
export const rowSpan = (slide: SlideInfo, repeat: RepeatDef, area: Box): { left: number; right: number } =>
  repeat.span && rowOnSlide(slide, repeat) ? repeat.span : { left: area.x, right: area.x + area.w };

/**
 * The most items a row can take on its slide: as many as fit side by side at their drawn width,
 * spreading out to the slide's margin, the next shape or the edge of the shape it is drawn on,
 * where the row is free to. Never fewer than it draws. A grid takes what it draws, because one
 * more would need another line.
 */
export function rowCapacity(slide: SlideInfo, repeat: RepeatDef, removed: ReadonlySet<string> = new Set()): number {
  const drawn = repeat.shapes.length;
  if (repeat.columns !== undefined || drawn < 2) return drawn;
  const room = rowRoom(slide, repeat.shapes, removed);
  // Further items are copies of the last, set out by the step from the first item to the second.
  // Where that step is shorter than the widest item, as it can be in a row whose items differ in
  // width, they would lie on top of each other. And where an arrow or a line lies among the items,
  // they would be drawn across it. Such rows take what they draw.
  // Positions are read to four decimals, so two values that are equal can differ by a little more than one of those.
  if (room.pinned || room.pitch < room.itemWidth - OVERLAP) return drawn;
  const span = rowSpan(slide, repeat, room.area);
  return Math.max(drawn, rowFit(room.itemWidth, span.left, span.right), rowOnSlide(slide, repeat) ? rowFit(room.itemWidth, room.left, room.right) : 0);
}

/**
 * By how much a step must fall short of an item's width before the next item counts as lying on
 * it. Positions are read to four decimals, and an item made of several shapes has its width from
 * four of them, so equal values can differ by up to a quarter of a thousandth.
 */
export const OVERLAP = 3e-4;

/** How far the space between a row's items may tighten before the row stops looking as designed: to a third. */
const TIGHTEST = 1 / 3;

/**
 * The most items a row takes while it still looks spaced as the template draws it: the space
 * between items no less than a third of the drawn one. This is what a draft settles on by itself;
 * a person may ask for more, up to {@link rowCapacity}. Never fewer than drawn. When the count
 * holds only with the row spread into the room beside it, the span it needs is returned with it,
 * and the count is made on that span, so that asking again after taking it gives the same answer.
 */
export function rowComfort(slide: SlideInfo, repeat: RepeatDef, removed: ReadonlySet<string> = new Set()): { most: number; span?: { left: number; right: number } } {
  const drawn = repeat.shapes.length;
  if (repeat.columns !== undefined || drawn < 2) return { most: drawn };
  const room = rowRoom(slide, repeat.shapes, removed);
  if (room.pinned) return { most: drawn };
  const current = rowSpan(slide, repeat, room.area);
  const gap = Math.max(room.gap * TIGHTEST, room.itemWidth * 0.03);
  // The renderer steps by the distance from the first item to the second, and every further item
  // is a copy of the last. A step that leaves less than the gap after the widest item would break
  // the promise, so such a row stays at what it draws: items that touch, or that differ in width
  // by more than the spacing can absorb.
  if (room.pitch < room.itemWidth + gap - 1e-4) return { most: drawn };
  const spaced = (left: number, right: number) => Math.floor((right - left + gap) / (room.itemWidth + gap) + 1e-6);
  const within = Math.max(drawn, spaced(current.left, current.right));
  if (!rowOnSlide(slide, repeat)) return { most: within };
  // Rounded outward to the precision spans are stored at, less the noise of the arithmetic.
  const left = Math.max(0, Math.floor(Math.min(current.left, room.left) * 1000 + 1e-6) / 1000);
  const right = Math.min(1, Math.ceil(Math.max(current.right, room.right) * 1000 - 1e-6) / 1000);
  const spread = spaced(left, right);
  return spread > within ? { most: spread, span: { left, right } } : { most: within };
}
