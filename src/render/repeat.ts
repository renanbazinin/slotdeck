import { NS, childElement, childElements, descendants, path, removeNode } from '../pptx/xml.ts';
import type { Element } from '../pptx/xml.ts';
import { COPY_ID_STEP, rowLayout } from '../template/manifest.ts';
import type { PatternDef, RepeatDef, SlideBox } from '../template/manifest.ts';

/**
 * For a row that may widen its items: the shapes of the row's text slots, which widen with them.
 * `undefined` for a pattern whose row keeps the size of its items.
 */
export function rowTextIds(pattern: PatternDef): Set<string> | undefined {
  const repeat = pattern.repeat;
  if (!repeat?.widen) return undefined;
  const names = new Set(repeat.shapes.flatMap((_, index) => repeat.slots.map((name) => name.replace('{n}', String(index + 1)))));
  return new Set(pattern.slots.filter((slot) => slot.kind === 'text' && names.has(slot.name)).map((slot) => slot.shapeId));
}

/** The transform element that positions a shape within its parent. */
export function shapeTransform(shape: Element): Element | undefined {
  if (shape.localName === 'grpSp') return path(shape, [NS.p, 'grpSpPr'], [NS.a, 'xfrm']);
  if (shape.localName === 'graphicFrame') return childElement(shape, NS.p, 'xfrm');
  return path(shape, [NS.p, 'spPr'], [NS.a, 'xfrm']);
}

interface Extent {
  x: number;
  y: number;
  right: number;
  bottom: number;
}

function extent(shape: Element): Extent {
  const xfrm = shapeTransform(shape);
  const off = xfrm && childElement(xfrm, NS.a, 'off');
  const ext = xfrm && childElement(xfrm, NS.a, 'ext');
  const x = Number(off?.getAttribute('x')) || 0;
  const y = Number(off?.getAttribute('y')) || 0;
  return { x, y, right: x + (Number(ext?.getAttribute('cx')) || 0), bottom: y + (Number(ext?.getAttribute('cy')) || 0) };
}

function union(extents: Extent[]): Extent {
  return {
    x: Math.min(...extents.map((e) => e.x)),
    y: Math.min(...extents.map((e) => e.y)),
    right: Math.max(...extents.map((e) => e.right)),
    bottom: Math.max(...extents.map((e) => e.bottom)),
  };
}

function moveBy(shape: Element, dx: number, dy: number): void {
  const off = childElement(shapeTransform(shape) as Element, NS.a, 'off');
  if (!off) return;
  off.setAttribute('x', String(Math.round(Number(off.getAttribute('x') ?? 0) + dx)));
  off.setAttribute('y', String(Math.round(Number(off.getAttribute('y') ?? 0) + dy)));
}

/**
 * The place a shape states on its slide, made where it states none: a placeholder takes its place
 * from its slide layout, and one that is put somewhere else states its own from then on, as it
 * does when a person moves it in PowerPoint. `undefined` for a shape that has nowhere to state one
 * (a group without a place: its children are placed by it).
 */
function ownPlace(shape: Element): { off: Element; ext: Element } | undefined {
  const doc = shape.ownerDocument;
  // A shape or a picture states its place among its properties; a group and a frame state theirs elsewhere, and always do.
  const props = childElement(shape, NS.p, 'spPr');
  const made = (parent: Element, name: string, before: Element | null): Element => {
    const element = (doc as NonNullable<typeof doc>).createElementNS(NS.a, `${shape.lookupPrefix(NS.a) ?? 'a'}:${name}`);
    parent.insertBefore(element, before);
    return element;
  };
  const makes = doc !== null && doc !== undefined && props !== undefined;
  // The place comes first among the properties of a shape, and where it starts before how large it is.
  const xfrm = shapeTransform(shape) ?? (makes ? made(props, 'xfrm', props.firstChild as Element | null) : undefined);
  if (!xfrm) return undefined;
  const off = childElement(xfrm, NS.a, 'off') ?? (makes ? made(xfrm, 'off', xfrm.firstChild as Element | null) : undefined);
  const ext = childElement(xfrm, NS.a, 'ext') ?? (makes && off ? made(xfrm, 'ext', off.nextSibling as Element | null) : undefined);
  return off && ext ? { off, ext } : undefined;
}

/**
 * Places a shape that is not inside a group at a given area of the slide. Sizes are in EMU. A
 * shape that takes its place from its slide layout (a placeholder that states none) is given a
 * place of its own (from 0.10.2; before, it was refused as if it were inside a group).
 */
export function placeShape(shape: Element, box: SlideBox, slideWidth: number, slideHeight: number): void {
  const parent = shape.parentNode as Element | null;
  if (!parent || parent.localName !== 'spTree') throw new Error('a box can only be set for a shape that is not inside a group');
  const place = ownPlace(shape);
  if (!place) throw new Error('a box can only be set for a shape that can state a place of its own: this one has none and nowhere to state one');
  const { off, ext } = place;
  off.setAttribute('x', String(Math.round(box.x * slideWidth)));
  off.setAttribute('y', String(Math.round(box.y * slideHeight)));
  ext.setAttribute('cx', String(Math.round(box.w * slideWidth)));
  ext.setAttribute('cy', String(Math.round(box.h * slideHeight)));
}

/**
 * Places a shape that keeps its content, such as a group or an ornament. When the area is the
 * shape's own size, only its position is set, so that nothing inside a group is rescaled by the
 * rounding of the box.
 */
export function moveShape(shape: Element, box: SlideBox, slideWidth: number, slideHeight: number): void {
  const xfrm = shapeTransform(shape);
  const off = xfrm && childElement(xfrm, NS.a, 'off');
  const ext = xfrm && childElement(xfrm, NS.a, 'ext');
  const same = (value: string | null | undefined, wanted: number, whole: number) => Math.abs(Number(value ?? 0) - wanted) <= whole * 0.0002;
  if ((shape.parentNode as Element | null)?.localName === 'spTree' && off && ext && same(ext.getAttribute('cx'), box.w * slideWidth, slideWidth) && same(ext.getAttribute('cy'), box.h * slideHeight, slideHeight)) {
    off.setAttribute('x', String(Math.round(box.x * slideWidth)));
    off.setAttribute('y', String(Math.round(box.y * slideHeight)));
    return;
  }
  placeShape(shape, box, slideWidth, slideHeight);
}

/** Outlines that stay what they are when only their width changes: rectangles, with or without rounded or cut corners. */
const RECTANGLE = /^(rect|\w*Rect)$/;
const LINE = /^(line|straightConnector1)$/;

/** Whether a shape can be made wider without spoiling it. A picture would be stretched, a circle would become an oval. */
function canWiden(shape: Element): boolean {
  const xfrm = shapeTransform(shape);
  if (!xfrm || Number(xfrm.getAttribute('rot') ?? 0) !== 0) return false;
  if (shape.localName === 'grpSp') return Number(childElement(xfrm, NS.a, 'chExt')?.getAttribute('cx') ?? 0) > 0;
  if (shape.localName !== 'sp' && shape.localName !== 'cxnSp') return false;
  const props = childElement(shape, NS.p, 'spPr');
  if (!props || childElement(props, NS.a, 'custGeom')) return false;
  const preset = childElement(props, NS.a, 'prstGeom')?.getAttribute('prst') ?? (shape.localName === 'cxnSp' ? 'line' : 'rect');
  // A level line stays level; a slanted one would change its slope.
  if (LINE.test(preset)) return Number(childElement(xfrm, NS.a, 'ext')?.getAttribute('cy') ?? 0) === 0;
  return RECTANGLE.test(preset);
}

/**
 * The corners of a rounded or cut rectangle are sized by its shorter side. When widening changes
 * which side that is, or how long it is, the corner setting is scaled back so the corners stay as drawn.
 */
/**
 * Keeps the corners of a rounded or snipped rectangle as they are drawn when the shape's shorter
 * side changes: from `min(width, height)` to `min(wider, height)`.
 */
export function keepCorners(shape: Element, width: number, wider: number, height: number): void {
  const geometry = path(shape, [NS.p, 'spPr'], [NS.a, 'prstGeom']);
  const preset = geometry?.getAttribute('prst') ?? '';
  const ratio = Math.min(width, height) / Math.min(wider, height);
  if (!geometry || !/^(round|snip)\w*Rect$/.test(preset) || !(ratio > 0) || Math.abs(ratio - 1) < 1e-6) return;
  const doc = shape.ownerDocument;
  const prefix = geometry.prefix ? `${geometry.prefix}:` : '';
  let list = childElement(geometry, NS.a, 'avLst');
  const guides = list ? childElements(list, NS.a, 'gd') : [];
  if (guides.length === 0) {
    // Only the plain rounded rectangle has a default worth restating here: a sixth of the shorter side.
    if (preset !== 'roundRect' || !doc) return;
    if (!list) {
      list = doc.createElementNS(NS.a, `${prefix}avLst`);
      geometry.appendChild(list);
    }
    const guide = doc.createElementNS(NS.a, `${prefix}gd`);
    guide.setAttribute('name', 'adj');
    guide.setAttribute('fmla', `val ${Math.round(16667 * ratio)}`);
    list.appendChild(guide);
    return;
  }
  for (const guide of guides) {
    const value = /^val (-?\d+)$/.exec(guide.getAttribute('fmla') ?? '');
    if (value) guide.setAttribute('fmla', `val ${Math.round(Number(value[1]) * ratio)}`);
  }
}

/**
 * Makes the shapes of one item wider by `delta`, each in its own way. `left` and `width` are the
 * item's own extent, in the units of the shapes.
 *
 * A background, a rule or a text box that spans the item grows with it and keeps its distance to
 * both sides. Everything else keeps its size and its distance to the nearer side, or stays in the
 * middle: an icon at the start of a card stays at the start. A group grows as its children do,
 * one by one, so that a picture inside it is never stretched. `texts` names the text boxes the
 * deck writes; they grow even when they are narrow.
 */
function widenShapes(shapes: Element[], left: number, width: number, delta: number, texts: ReadonlySet<string>): void {
  for (const shape of shapes) {
    const xfrm = shapeTransform(shape);
    const off = xfrm && childElement(xfrm, NS.a, 'off');
    const ext = xfrm && childElement(xfrm, NS.a, 'ext');
    if (!xfrm || !off || !ext) continue;
    const x = Number(off.getAttribute('x') ?? 0);
    const own = Number(ext.getAttribute('cx') ?? 0);
    const before = x - left;
    const after = left + width - (x + own);
    const id = descendants(shape, NS.p, 'cNvPr')[0]?.getAttribute('id') ?? '';
    if (canWiden(shape) && (texts.has(id) || own >= Math.max(before, after) - 1)) {
      if (shape.localName === 'grpSp') {
        const inner = childElement(xfrm, NS.a, 'chExt') as Element;
        const innerWidth = Number(inner.getAttribute('cx') ?? 0);
        // The group's children are measured in its own units, which may be scaled.
        const innerDelta = own > 0 ? (delta * innerWidth) / own : delta;
        const children = childElements(shape, NS.p).filter((child) => ['sp', 'pic', 'grpSp', 'cxnSp', 'graphicFrame'].includes(child.localName ?? ''));
        widenShapes(children, Number(childElement(xfrm, NS.a, 'chOff')?.getAttribute('x') ?? 0), innerWidth, innerDelta, texts);
        inner.setAttribute('cx', String(Math.round(innerWidth + innerDelta)));
      } else {
        keepCorners(shape, own, own + delta, Number(ext.getAttribute('cy') ?? 0));
      }
      ext.setAttribute('cx', String(Math.round(own + delta)));
      continue;
    }
    const centred = Math.abs(before - after) <= Math.max(width * 0.02, 1);
    const shift = centred ? delta / 2 : before <= after ? 0 : delta;
    if (shift !== 0) off.setAttribute('x', String(Math.round(x + shift)));
  }
}

/**
 * Gives a row or grid of items the requested number: removes the surplus or copies the last
 * item, then lays the items out again around the centre of the area they occupied. Rows keep
 * the template's spacing when there is room and tighten when there is not; a grid uses as few
 * rows as the count needs, each row centred.
 *
 * With `widen` on the repeat, a line that holds fewer items than the template draws has its items
 * made wider, up to that factor, so that they fill the row's span at the template's own gaps.
 * `texts` are the ids of the row's text boxes that the deck writes. `mirrored`: the deck reads the
 * other way from the template, and a row with an item that stands out (`emphasis`) sets its items
 * out in the deck's direction.
 *
 * Copied shapes get the id of their original plus a multiple of {@link COPY_ID_STEP}.
 */
export function applyRepeat(shapes: Map<string, Element>, repeat: RepeatDef, count: number, slideWidth: number, texts: ReadonlySet<string> = new Set(), mirrored = false): void {
  const items = repeat.shapes.map((ids) =>
    ids.map((id) => {
      const shape = shapes.get(id);
      if (!shape) throw new Error(`repeat item points at shape #${id}, which is not on the template slide`);
      return shape;
    }),
  );
  const drawn = items.length;
  const first = items[0]?.[0];
  if (!first) throw new Error('a repeat needs at least one item with a shape');

  const boxes = items.map((item) => union(item.map(extent)));
  let width = Math.max(...boxes.map((box) => box.right - box.x));
  const height = Math.max(...boxes.map((box) => box.bottom - box.y));
  const area = union(boxes);

  // Spacing and direction are read off the template slide: the first two items of the first
  // row give the column pitch and the reading direction; the first items of two rows the row pitch.
  const drawnRows = rowLayout(repeat, drawn);
  const second = (drawnRows[0] ?? 0) > 1 ? boxes[1] : undefined;
  let pitch = second ? Math.abs(second.x - (boxes[0] as Extent).x) : width * 1.5;
  const rightToLeft = second ? second.x < (boxes[0] as Extent).x : true;
  const nextRow = drawnRows.length > 1 ? boxes[drawnRows[0] as number] : undefined;
  const rowPitch = nextRow ? Math.abs(nextRow.y - (boxes[0] as Extent).y) : height * 1.2;

  // A span in slide fractions only makes sense for shapes placed directly on the slide.
  const onSlide = (first.parentNode as Element | null)?.localName === 'spTree';
  const span = repeat.span && onSlide ? { left: repeat.span.left * slideWidth, right: repeat.span.right * slideWidth } : { left: area.x, right: area.right };

  for (const item of items.slice(count)) for (const shape of item) removeNode(shape);
  items.length = Math.min(drawn, count);
  boxes.length = items.length;

  const original = items[drawn - 1];
  for (let copy = 1; original && items.length < count; copy++) {
    let after = (items[items.length - 1] as Element[]).at(-1) as Element;
    const clones = original.map((shape) => {
      const clone = shape.cloneNode(true) as Element;
      for (const props of descendants(clone, NS.p, 'cNvPr')) props.setAttribute('id', String(Number(props.getAttribute('id')) + copy * COPY_ID_STEP));
      after.parentNode?.insertBefore(clone, after.nextSibling);
      after = clone;
      return clone;
    });
    items.push(clones);
    boxes.push(boxes[drawn - 1] as Extent);
  }

  const rows = rowLayout(repeat, count);
  // Fewer items on a line than drawn: with the setting, they grow to fill the line. The count
  // alone decides this, never the length of the text, so the same deck always looks the same.
  const held = Math.max(0, ...rows);
  if (repeat.widen && repeat.widen > 1 && second && held > 0 && held < (drawnRows[0] as number)) {
    const gap = pitch - width;
    const wider = Math.min(width * repeat.widen, (span.right - span.left - (held - 1) * gap) / held);
    if (wider > width + 1) {
      items.forEach((item, index) => {
        const box = boxes[index] as Extent;
        widenShapes(item, box.x, box.right - box.x, wider - width, texts);
      });
      width = wider;
      pitch = wider + gap;
    }
  }
  // In a deck that reads the other way, a row with an item that stands out sets its items out in
  // the deck's direction, each with its look; any other row keeps its places and has its values mirrored.
  const fromRight = rightToLeft !== (mirrored && repeat.emphasis !== undefined);
  // Fewer lines than drawn (an agenda of three where six are drawn): with `spread`, the space
  // between them grows, up to that factor on the drawn step, as far as the area the lines were
  // drawn in, so that they do not bunch in its middle. The count alone decides it, as for widen.
  let linePitch = rowPitch;
  if (repeat.spread !== undefined && repeat.spread > 1 && rows.length > 1 && rows.length < drawnRows.length) {
    linePitch = Math.max(rowPitch, Math.min(rowPitch * repeat.spread, (area.bottom - area.y - height) / (rows.length - 1)));
  }
  const totalHeight = (rows.length - 1) * linePitch + height;
  const top = (area.y + area.bottom) / 2 - totalHeight / 2;
  const centre = (area.x + area.right) / 2;
  let index = 0;
  rows.forEach((size, row) => {
    const step = size > 1 ? Math.min(pitch, (span.right - span.left - width) / (size - 1)) : 0;
    const total = (size - 1) * step + width;
    const start = Math.min(Math.max(centre - total / 2, span.left), Math.max(span.left, span.right - total));
    for (let position = 0; position < size; position++, index++) {
      const box = boxes[index] as Extent;
      const x = start + (fromRight ? size - 1 - position : position) * step;
      const y = top + row * linePitch;
      for (const shape of items[index] as Element[]) moveBy(shape, x - box.x, y - box.y);
    }
  });
}
