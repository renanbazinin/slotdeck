// What code can work out about a template's shapes without a model: which belong to the frame,
// how much text a box holds, what a shape could be used for. Shared by the drafter, the manual
// changes and the studio view, so a slot made by hand gets the same limits as a drafted one.
import type { Box, ShapeInfo, SlideInfo, TemplateInfo } from './inspect.ts';
import type { SlotDef } from './manifest.ts';

/** What a shape is for. `fixed` means it is part of the design and not a slot. */
export const ROLES = ['title', 'subtitle', 'text', 'list', 'number', 'caption', 'photo', 'icon', 'graphic', 'chart', 'table', 'decoration', 'fixed'] as const;
export type Role = (typeof ROLES)[number];
const TEXT_ROLES = new Set<Role>(['title', 'subtitle', 'text', 'list', 'number', 'caption']);

export const pct = (value: number) => Math.round(value * 100);
export const normalize = (text: string) => text.replace(/\s+/g, ' ').trim();
export const shapeText = (shape: ShapeInfo) => normalize((shape.paragraphs ?? []).join(' '));

/** A name the manifest can use as a slot name: letters and digits, starting with a lowercase letter. */
export function slotName(raw: string, fallback: string): string {
  const words = raw.replace(/[^A-Za-z0-9]+/g, ' ').trim().split(' ').filter(Boolean);
  const name = words.map((word, index) => (index === 0 ? word.charAt(0).toLowerCase() : word.charAt(0).toUpperCase()) + word.slice(1)).join('');
  return /^[a-z]/.test(name) ? name : fallback;
}

/** How much text a box holds, estimated from its size and font, and never less than its sample holds. */
export function textLimits(shape: ShapeInfo, info: TemplateInfo, role: Role): { size: number; maxChars: number; maxParagraphs?: number } {
  const size = shape.fontSizes?.[0] ?? 14;
  const paragraphs = (shape.paragraphs ?? []).filter((paragraph) => paragraph.trim() !== '');
  const sample = paragraphs.reduce((sum, paragraph) => sum + paragraph.length, 0);
  const perLine = Math.max(1, Math.floor((shape.box.w * info.slideWidth) / 12700 / (size * 0.5)));
  // A text box often shrinks to its sample, so the lines the sample takes count as well as the box height.
  const sampleLines = paragraphs.reduce((sum, paragraph) => sum + Math.max(1, Math.ceil(paragraph.length / perLine)), 0);
  const boxLines = Math.max(1, Math.round((shape.box.h * info.slideHeight) / 12700 / (size * 1.25)));
  const estimate = Math.floor(perLine * Math.max(sampleLines, boxLines) * 0.8);
  const limits: { size: number; maxChars: number; maxParagraphs?: number } = {
    size,
    // A big figure fits what the sample shows and little more; other text can use the room there is.
    maxChars: Math.max(1, role === 'number' ? Math.max(sample, 2) : Math.max(sample, estimate)),
  };
  if (role === 'list') limits.maxParagraphs = Math.max(2, paragraphs.length);
  return limits;
}

/**
 * How many characters an area of the slide holds at a font size, when the text is in several
 * paragraphs with a gap between them. Each paragraph after the first ends half a line short on average.
 */
export function boxCapacity(box: Box, info: TemplateInfo, size: number, paragraphs = 1, gap = 0): number {
  const extra = Math.max(0, paragraphs - 1);
  const perLine = Math.max(1, Math.floor((box.w * info.slideWidth) / 12700 / (size * 0.5)));
  const height = (box.h * info.slideHeight) / 12700 - extra * gap;
  const lines = Math.max(1, Math.round(height / (size * 1.25)) - extra * 0.5);
  return Math.max(1, Math.floor(perLine * lines * 0.8));
}

/**
 * What a table holds: its columns and rows as drawn, and the characters that fit in a cell
 * before the rows grow so tall that the table leaves the slide. The room is measured from the
 * table's top down to the slide's bottom margin, or to the next shape under the table.
 * `fontSize` and `columns` give the limits for a table set differently from the sample.
 */
export function tableLimits(
  shape: ShapeInfo,
  info: TemplateInfo,
  options: { slide?: SlideInfo; fontSize?: number; columns?: number; rows?: number } = {},
): { maxColumns: number; maxRows: number; maxCellChars: number } {
  const columns = options.columns ?? shape.table?.columns ?? 4;
  const bodyRows = options.rows ?? Math.max(1, (shape.table?.rows ?? 4) - 1);
  const size = options.fontSize ?? shape.fontSizes?.[0] ?? info.defaultFontSize;
  const margins = 14.4; // a cell's two opposite margins, in points
  const points = (fraction: number, of: number) => (fraction * of) / 12700;

  let bottom = 0.93;
  for (const other of options.slide?.shapes ?? []) {
    if (other.id === shape.id || other.kind === 'group' || isFramePlaceholder(other)) continue;
    const beside = other.box.x < shape.box.x + shape.box.w && other.box.x + other.box.w > shape.box.x;
    if (beside && other.box.y > shape.box.y + 0.05) bottom = Math.min(bottom, other.box.y - 0.02);
  }
  const width = points(shape.box.w, info.slideWidth) / columns - margins;
  const height = points(Math.max(bottom - shape.box.y, 0), info.slideHeight) / (bodyRows + 1) - margins;
  const perLine = Math.max(1, Math.floor(width / (size * 0.5)));
  const lines = Math.max(1, Math.floor(height / (size * 1.2) + 0.05));
  return { maxColumns: columns, maxRows: bodyRows, maxCellChars: Math.max(4, Math.floor(perLine * lines * 0.9)) };
}

/** The roles a shape can take, by what kind of shape it is. `fixed` is always possible and not listed. */
export function rolesFor(shape: ShapeInfo): Role[] {
  if (shape.kind === 'text') return ['title', 'subtitle', 'text', 'list', 'number', 'caption', 'decoration'];
  if (shape.kind === 'picture') return ['photo', 'icon', 'graphic', 'chart', 'decoration'];
  if (shape.kind === 'table') return ['table'];
  // A drawn ornament, a line or a group can only be taken away.
  return ['decoration'];
}

/**
 * The slot a shape becomes in a given role, with limits measured from the shape. `undefined`
 * when the shape cannot take the role (a picture cannot be a title).
 */
export function slotFromShape(shape: ShapeInfo, info: TemplateInfo, role: Role, name: string, optional = false, slide?: SlideInfo): SlotDef | undefined {
  if (!rolesFor(shape).includes(role)) return undefined;
  let slot: SlotDef;
  if (shape.kind !== 'text' && shape.kind !== 'picture' && shape.kind !== 'table') {
    // Nothing fills such a shape; the slot exists so that the shape is removed.
    return { name, kind: 'text', shapeId: shape.id, empty: 'drop' };
  }
  if (shape.kind === 'text') {
    slot = { name, kind: 'text', shapeId: shape.id, ...textLimits(shape, info, role) };
  } else if (shape.kind === 'table') {
    return { name, kind: 'table', shapeId: shape.id, ...tableLimits(shape, info, slide ? { slide } : {}) };
  } else if (role === 'chart') {
    // A chart replaces the sample picture and may not sit in a group, which a native chart cannot.
    if (shape.groupId !== undefined) return undefined;
    return { name, kind: 'chart', shapeId: shape.id, maxCategories: 8, maxSeries: 3 };
  } else {
    slot = { name, kind: 'image', imageKind: role === 'photo' || role === 'icon' ? role : 'graphic', shapeId: shape.id };
  }
  if (role === 'decoration') slot.empty = 'drop';
  else if (optional) slot.empty = 'remove';
  return slot;
}

/** The longest side a pictogram has, in inches. A photograph in a layout is drawn larger than this. */
const PICTOGRAM_SIDE = 1.1;

/**
 * Whether a picture is a pictogram by what can be measured: no side longer than 1.1 inch, not a
 * JPEG (the format of photographs), and not cropped into a frame as a portrait is. A layout that
 * asks for a photograph in a place this small gets an icon instead.
 */
export function isPictogram(shape: ShapeInfo, info: TemplateInfo): boolean {
  if (shape.kind !== 'picture' || !shape.image) return false;
  const inches = (fraction: number, of: number) => (fraction * of) / 914400;
  if (inches(shape.box.w, info.slideWidth) > PICTOGRAM_SIDE || inches(shape.box.h, info.slideHeight) > PICTOGRAM_SIDE) return false;
  if (/\.jpe?g$/i.test(shape.image.part) || shape.image.cropped) return false;
  return shape.image.geometry === 'rect' || shape.image.geometry === 'custom';
}

/** Whether a shape of a role's kind carries text. */
export function isTextRole(role: Role): boolean {
  return TEXT_ROLES.has(role);
}

/**
 * A page number, date, footer or header placeholder: part of every slide, never content. A title
 * or body placeholder is content like any other text box.
 */
export function isFramePlaceholder(shape: ShapeInfo): boolean {
  return shape.placeholder === 'sldNum' || shape.placeholder === 'dt' || shape.placeholder === 'ftr' || shape.placeholder === 'hdr';
}

export interface FrameFacts {
  /** On how many slides something must sit in the same place to count as part of the frame. */
  often: number;
  /** A logo, rule, page number or other element that repeats across slides. */
  isFrame(shape: ShapeInfo): boolean;
  /** Texts that appear on many slides: one value for the whole deck. */
  repeatedTexts: string[];
  /** Every distinct sample text, in slide order. */
  texts: string[];
}

/**
 * Whatever sits in the same place on many slides is part of the frame, not content. A short deck
 * gives too little evidence of what repeats, so nothing is treated as frame there.
 */
export function frameFacts(info: TemplateInfo): FrameFacts {
  const total = info.slides.length;
  const often = total >= 6 ? Math.max(3, Math.ceil(total * 0.3)) : Infinity;
  const placeOf = (shape: ShapeInfo) => `${shape.kind}|${pct(shape.box.x)}|${pct(shape.box.y)}|${pct(shape.box.w)}|${pct(shape.box.h)}|${shape.image?.part ?? ''}`;
  const places = new Map<string, number>();
  const texts = new Map<string, number>();
  for (const slide of info.slides) {
    for (const key of new Set(slide.shapes.filter((shape) => shape.kind !== 'text').map(placeOf))) places.set(key, (places.get(key) ?? 0) + 1);
    for (const text of new Set(slide.shapes.filter((shape) => shape.kind === 'text' && !isFramePlaceholder(shape)).map(shapeText))) {
      if (text) texts.set(text, (texts.get(text) ?? 0) + 1);
    }
  }
  return {
    often,
    isFrame: (shape) => isFramePlaceholder(shape) || (shape.kind !== 'text' && (places.get(placeOf(shape)) ?? 0) >= often),
    repeatedTexts: [...texts].filter(([, count]) => count >= often).map(([text]) => text),
    texts: [...texts.keys()],
  };
}

/** One line of measured facts about a shape, as the model reads it. */
export function describeShape(shape: ShapeInfo, info: TemplateInfo): string {
  const where = `at x${pct(shape.box.x)} y${pct(shape.box.y)}, ${pct(shape.box.w)} wide ${pct(shape.box.h)} high${shape.groupId ? `, in group #${shape.groupId}` : ''}`;
  if (shape.kind === 'text') {
    const text = shapeText(shape);
    const count = shape.paragraphs?.length ?? 1;
    return `#${shape.id} text ${where} | ${shape.fontSizes?.[0] ?? 14} pt | ${count} paragraph${count === 1 ? '' : 's'}, ${text.length} characters: ${JSON.stringify(text.slice(0, 60))}`;
  }
  if (shape.kind === 'table') return `#${shape.id} table ${where} | ${shape.table?.columns ?? '?'} columns, ${shape.table?.rows ?? '?'} rows`;
  if (shape.kind !== 'picture') return `#${shape.id} ${shape.kind} ${where}`;
  const image = shape.image;
  const uses = info.slides.filter((other) => other.shapes.some((candidate) => candidate.image?.part === image?.part)).length;
  return `#${shape.id} picture ${where} | ${image?.width ?? '?'}x${image?.height ?? '?'} px ${image?.part.split('.').pop() ?? ''}${image?.cropped ? ', cropped to its frame' : ''}${image?.geometry && image.geometry !== 'rect' && image.geometry !== 'custom' ? `, ${image.geometry} frame` : ''}${uses > 1 ? `, same picture on ${uses} slides` : ''}`;
}

/** The smallest box around several shapes, as slide fractions. */
export function unionBox(boxes: Box[]): Box {
  const x = Math.min(...boxes.map((box) => box.x));
  const y = Math.min(...boxes.map((box) => box.y));
  return { x, y, w: Math.max(...boxes.map((box) => box.x + box.w)) - x, h: Math.max(...boxes.map((box) => box.y + box.h)) - y };
}

/** What {@link rowRoom} finds about a row on its slide. All lengths are fractions of the slide width. */
export interface RowRoom {
  /** The row as drawn. */
  area: Box;
  /** The width of the widest item. */
  itemWidth: number;
  /** From the first item to the second: the step the renderer sets further items out by. */
  pitch: number;
  /** The space the template draws between neighbouring items, on average. Never below zero. */
  gap: number;
  /** How far the row may spread. */
  left: number;
  right: number;
  /**
   * A shape that stays where it is lies among the items: an arrow between two of them, a line
   * through them, a highlight behind some of them. More items would be drawn across it, so the
   * row holds what it draws.
   */
  pinned: boolean;
}

/**
 * How far a row of items may spread on its slide: from the row as drawn outwards, until the
 * slide's margin, another shape that sits beside the row, or the edge of a shape the row is drawn
 * on.
 */
export function rowRoom(slide: SlideInfo, itemShapes: string[][]): RowRoom {
  const byId = new Map(slide.shapes.map((shape) => [shape.id, shape]));
  const inRow = new Set<string>();
  const mark = (id: string) => {
    inRow.add(id);
    for (const shape of slide.shapes) if (shape.groupId === id) mark(shape.id);
  };
  for (const id of itemShapes.flat()) {
    mark(id);
    // A group that holds the whole row is not beside it.
    for (let parent = byId.get(id)?.groupId; parent; parent = byId.get(parent)?.groupId) inRow.add(parent);
  }
  const boxes = itemShapes.map((ids) => unionBox(ids.flatMap((id) => (byId.get(id) ? [(byId.get(id) as ShapeInfo).box] : []))));
  const area = unionBox(boxes);
  const itemWidth = Math.max(...boxes.map((box) => box.w));

  // The margin the slide's own content keeps from its sides.
  const content = slide.shapes.filter((shape) => !isFramePlaceholder(shape) && shape.kind !== 'group' && shape.box.w < 0.9 && shape.box.x >= 0 && shape.box.x + shape.box.w <= 1);
  const margin = Math.max(0.03, Math.min(0.08, ...content.map((shape) => Math.min(shape.box.x, 1 - shape.box.x - shape.box.w)).filter((edge) => edge > 0.005)));
  const gap = 0.02;
  const edge = 0.005;
  let left = Math.min(area.x, margin);
  let right = Math.max(area.x + area.w, 1 - margin);
  let pinned = false;
  const areaEnd = area.x + area.w;
  for (const shape of slide.shapes) {
    if (inRow.has(shape.id) || isFramePlaceholder(shape) || shape.kind === 'group') continue;
    const end = shape.box.x + shape.box.w;
    // The height the shape shares with the row. A line has no height: it counts when it lies
    // inside the row's height, not when it runs along the row's top or bottom edge.
    const shared = Math.min(shape.box.y + shape.box.h, area.y + area.h) - Math.max(shape.box.y, area.y);
    if (shape.box.h > 0 ? shared <= 0 : !(shape.box.y > area.y + edge && shape.box.y < area.y + area.h - edge)) continue;
    if (end <= area.x + edge) {
      if (shape.box.w < 0.9) left = Math.max(left, Math.min(area.x, end + gap));
      continue;
    }
    if (shape.box.x >= areaEnd - edge) {
      if (shape.box.w < 0.9) right = Math.min(right, Math.max(areaEnd, shape.box.x - gap));
      continue;
    }
    // From here the shape shares width with the row. A background from edge to edge of the slide
    // limits nothing, and neither does a shape that only brushes the row's height, such as a
    // heading whose box dips into it.
    if (shape.box.x <= edge && end >= 1 - edge) continue;
    if (shape.box.h > 0 && shared < Math.min(shape.box.h, area.h) * 0.5) continue;
    const coversStart = shape.box.x <= area.x + edge;
    const coversEnd = end >= areaEnd - edge;
    const overlap = Math.min(end, areaEnd) - Math.max(shape.box.x, area.x);
    if (coversStart && coversEnd) {
      // The row is drawn on it, as on a tinted panel: the row stays a little inside its edges.
      left = Math.max(left, Math.min(area.x, shape.box.x + gap));
      right = Math.min(right, Math.max(areaEnd, end - gap));
    } else if (coversEnd && overlap < itemWidth * 0.5) {
      // A neighbour whose box reaches a little into the row's last item: as good as beside it.
      right = Math.min(right, areaEnd);
    } else if (coversStart && overlap < itemWidth * 0.5) {
      left = Math.max(left, area.x);
    } else {
      // It lies among the items and stays where it is.
      pinned = true;
      left = area.x;
      right = areaEnd;
    }
  }
  if (pinned) {
    left = area.x;
    right = areaEnd;
  }
  const ordered = boxes.slice().sort((one, other) => one.x - other.x);
  const gaps = ordered.slice(1).map((box, index) => box.x - ((ordered[index] as Box).x + (ordered[index] as Box).w));
  const between = gaps.length > 0 ? Math.max(0, gaps.reduce((sum, value) => sum + value, 0) / gaps.length) : 0;
  const pitch = boxes.length > 1 ? Math.abs((boxes[1] as Box).x - (boxes[0] as Box).x) : itemWidth;
  return { area, itemWidth, pitch, gap: between, left, right, pinned };
}
