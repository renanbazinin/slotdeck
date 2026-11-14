import { NS, childElement, descendants, path, removeNode } from '../pptx/xml.ts';
import type { Element } from '../pptx/xml.ts';
import { COPY_ID_STEP, rowLayout } from '../template/manifest.ts';
import type { RepeatDef, SlideBox } from '../template/manifest.ts';

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

/** Places a shape that is not inside a group at a given area of the slide. Sizes are in EMU. */
export function placeShape(shape: Element, box: SlideBox, slideWidth: number, slideHeight: number): void {
  const parent = shape.parentNode as Element | null;
  const xfrm = shapeTransform(shape);
  const off = xfrm && childElement(xfrm, NS.a, 'off');
  const ext = xfrm && childElement(xfrm, NS.a, 'ext');
  if (!parent || parent.localName !== 'spTree' || !off || !ext) throw new Error('a box can only be set for a shape that is not inside a group');
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

/**
 * Gives a row or grid of items the requested number: removes the surplus or copies the last
 * item, then lays the items out again around the centre of the area they occupied. Rows keep
 * the template's spacing when there is room and tighten when there is not; a grid uses as few
 * rows as the count needs, each row centred.
 *
 * Copied shapes get the id of their original plus a multiple of {@link COPY_ID_STEP}.
 */
export function applyRepeat(shapes: Map<string, Element>, repeat: RepeatDef, count: number, slideWidth: number): void {
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
  const width = Math.max(...boxes.map((box) => box.right - box.x));
  const height = Math.max(...boxes.map((box) => box.bottom - box.y));
  const area = union(boxes);

  // Spacing and direction are read off the template slide: the first two items of the first
  // row give the column pitch and the reading direction; the first items of two rows the row pitch.
  const drawnRows = rowLayout(repeat, drawn);
  const second = (drawnRows[0] ?? 0) > 1 ? boxes[1] : undefined;
  const pitch = second ? Math.abs(second.x - (boxes[0] as Extent).x) : width * 1.5;
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
  const totalHeight = (rows.length - 1) * rowPitch + height;
  const top = (area.y + area.bottom) / 2 - totalHeight / 2;
  const centre = (area.x + area.right) / 2;
  let index = 0;
  rows.forEach((size, row) => {
    const step = size > 1 ? Math.min(pitch, (span.right - span.left - width) / (size - 1)) : 0;
    const total = (size - 1) * step + width;
    const start = Math.min(Math.max(centre - total / 2, span.left), Math.max(span.left, span.right - total));
    for (let position = 0; position < size; position++, index++) {
      const box = boxes[index] as Extent;
      const x = start + (rightToLeft ? size - 1 - position : position) * step;
      const y = top + row * rowPitch;
      for (const shape of items[index] as Element[]) moveBy(shape, x - box.x, y - box.y);
    }
  });
}
