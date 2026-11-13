// What a studio screen needs to draw a template for editing by hand: every slide's shapes with
// their place on the slide, what each one is in the manifest, and what it could become.
import type { Template } from '../render/render.ts';
import { describeSlot } from './change.ts';
import { inspectTemplate } from './inspect.ts';
import type { Box, ShapeKind } from './inspect.ts';
import type { EmptyPolicy, SlotDef } from './manifest.ts';
import { frameFacts, rolesFor, shapeText } from './measure.ts';
import type { Role } from './measure.ts';

export interface StudioShape {
  /** The id a change names the shape by. */
  id: string;
  kind: ShapeKind;
  /** Where the shape is, as fractions of the slide: lay it over a picture of the slide. */
  box: Box;
  /** The group the shape sits in, when it does. */
  groupId?: string;
  /** Sample text, for a shape that has some. */
  text?: string;
  /** Largest font size of the sample text, in points. */
  fontSize?: number;
  /** The shape repeats across slides (logo, header rule, page number), so it is rarely content. */
  frame: boolean;
  /** The roles `addSlot` accepts for this shape. */
  roles: Role[];
  /** The slots that point at the shape, one for each pattern made from the slide that uses it. */
  slots: Array<{ pattern: string; name: string; kind: SlotDef['kind']; empty: EmptyPolicy; summary: string }>;
}

export interface StudioSlide {
  /** 1-based slide number in the template. */
  slide: number;
  /** Ids of the patterns made from this slide. Empty when the slide is not used. */
  patterns: string[];
  shapes: StudioShape[];
}

export interface StudioView {
  /** Width of a slide divided by its height. */
  aspectRatio: number;
  slides: StudioSlide[];
}

/**
 * Joins the pptx and the manifest into one view for a studio screen. Nothing here is stored:
 * call it again after every change.
 */
export function studioView(template: Template): StudioView {
  const info = inspectTemplate(template.pkg);
  const frame = frameFacts(info);
  const slides = info.slides.map((slide): StudioSlide => {
    const patterns = template.manifest.patterns.filter((pattern) => pattern.sourceSlide === slide.number);
    const shapes = slide.shapes.map((shape): StudioShape => {
      const view: StudioShape = {
        id: shape.id,
        kind: shape.kind,
        box: shape.box,
        frame: frame.isFrame(shape),
        roles: rolesFor(shape),
        slots: patterns.flatMap((pattern) =>
          pattern.slots
            .filter((slot) => slot.shapeId === shape.id)
            .map((slot) => ({ pattern: pattern.id, name: slot.name, kind: slot.kind, empty: slot.empty ?? 'error', summary: describeSlot(slot) })),
        ),
      };
      if (shape.groupId !== undefined) view.groupId = shape.groupId;
      const text = shapeText(shape);
      if (text) view.text = text;
      const size = shape.fontSizes?.[0];
      if (size !== undefined) view.fontSize = size;
      return view;
    });
    return { slide: slide.number, patterns: patterns.map((pattern) => pattern.id), shapes };
  });
  return { aspectRatio: info.slideHeight > 0 ? info.slideWidth / info.slideHeight : 16 / 9, slides };
}
