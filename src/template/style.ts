// A manifest made from a pptx alone, with no model and nobody's judgement: enough of one for the
// software to take the template's style (its frame, its text styles, its colours) and lay slides
// out itself. Every slide becomes a layout whose texts are slots, named by what can be measured:
// the title is the largest text near the top, a number is a short figure set large, the rest is
// text. It does not say what a slide is for, so only three of its layouts are meant to be used as
// they are: a cover, a closing slide, and the frame that free slides stand on.
//
// Style-only mode (from 0.10.0): see docs/style-only.md.
import { pageColor, paletteOf, schemeOf } from '../color/recolor.ts';
import { freeKitOrReason } from '../free/kit.ts';
import { PptxPackage } from '../pptx/package.ts';
import { NS, childElement, childElements, parseXml, path, rootElement } from '../pptx/xml.ts';
import type { Element } from '../pptx/xml.ts';
import { TextStyles } from '../text/styles.ts';
import { inspectTemplate } from './inspect.ts';
import type { Box, ShapeInfo, SlideInfo, TemplateInfo } from './inspect.ts';
import type { Direction, PatternDef, SlotDef, TemplateManifest } from './manifest.ts';
import { frameFacts, inFrameBand, isFramePlaceholder, isPictogram, shapeText, slotFromShape } from './measure.ts';

/** What `styleManifest` is told. */
export interface StyleManifestOptions {
  /** The manifest's name. Left out, `style`. */
  name?: string;
}

/** What `styleManifest` makes of a pptx. */
export interface StyleManifest {
  /** A manifest like any other: `loadTemplate(pptx, manifest)` takes it, and `checkManifest` passes it. */
  manifest: TemplateManifest;
  /**
   * The layouts a deck in the template's style is planned from: the cover where there is one,
   * `free` (the slides the software lays out) where the template can carry them, and the closing
   * slide where there is one. Give it as `patterns` to a brief or to `presentationTools`: the
   * manifest's other layouts are the template's slides as they are drawn, with slots that no
   * person named, and are there for their text styles.
   */
  patterns: string[];
  /**
   * Why the template cannot carry slides the software lays out, in a sentence; left out where it
   * can. With a reason, `patterns` has no `free` in it and there is no style-only mode for this
   * file: it needs a manifest that a person prepares (`prepareManifest`).
   */
  reason?: string;
  /** The id of the layout a deck opens with, where the template's first slide can be one. */
  cover?: string;
  /** The id of the layout a deck ends with, where the template's last slide can be one. */
  closing?: string;
  /** The deck's own texts the frame shows (a department, a date), by the name of the field each takes, with the sample the template shows there. */
  fields: Record<string, string>;
  /** What was decided and why, in sentences: for a person who looks at the result. */
  notes: string[];
}

const RTL = /[\p{Script=Hebrew}\p{Script=Arabic}]/u;
const bottom = (box: Box) => box.y + box.h;
const right = (box: Box) => box.x + box.w;
const round = (value: number) => Math.round(value * 10000) / 10000;

/** A short figure: digits with their signs and units, and no words. */
const FIGURE = /^[\s\d.,:%+\-−–/$€₪£xXKMBkmb°]+$/;

/**
 * Makes a manifest for a template from the pptx alone, by rules: no model is called and nobody
 * prepares anything (from 0.10.0; see docs/style-only.md). The software takes the template's
 * style (the frame of its slides, its text styles, its colours) and lays the slides out itself,
 * between the template's own cover and closing slide.
 *
 * ```ts
 * const style = styleManifest(pptx, { name: 'Our deck' });
 * if (style.reason) throw new Error(style.reason);          // the file needs a prepared manifest
 * const template = loadTemplate(pptx, style.manifest);
 * const { pptx: deck } = await generateDeck(template, { prompt, patterns: style.patterns, fields: { field1: 'Sales' } }, { llm, layout: 'adaptive' });
 * ```
 *
 * Deterministic: the same file always gives the same manifest, so it may be stored or made again.
 */
export function styleManifest(pptx: Uint8Array | PptxPackage, options: StyleManifestOptions = {}): StyleManifest {
  const pkg = pptx instanceof PptxPackage ? pptx : PptxPackage.load(pptx);
  const read = inspectTemplate(pkg);
  const styles = new TextStyles(pkg);
  const notes: string[] = [];

  // A placeholder that states no place and no size of text takes them from its layout and its
  // master: it is seen here where it stands and at the size it is set, as a text box is.
  let inherits = 0;
  const seen = (slide: SlideInfo, shape: ShapeInfo): ShapeInfo => {
    if (shape.kind !== 'text' || shape.placeholder === undefined) return shape;
    let box = shape.box;
    if (!(box.w > 0 || box.h > 0)) {
      const from = styles.inheritedBox(slide.number, shape.id);
      if (from && from.cx > 0 && from.cy > 0) box = { x: round(from.x / read.slideWidth), y: round(from.y / read.slideHeight), w: round(from.cx / read.slideWidth), h: round(from.cy / read.slideHeight) };
    }
    const size = shape.fontSizes?.length ? undefined : styles.frame(slide.number, shape.id)?.paragraphs[0]?.run.size;
    if (box === shape.box && size === undefined) return shape;
    inherits++;
    return { ...shape, box, ...(size === undefined ? {} : { fontSizes: [Math.round(size * 2) / 2] }) };
  };
  const info: TemplateInfo = { ...read, slides: read.slides.map((slide) => ({ ...slide, shapes: slide.shapes.map((shape) => seen(slide, shape)) })) };
  const frame = frameFacts(info);
  const placeOf = (_slide: SlideInfo, shape: ShapeInfo): Box => shape.box;

  // The direction and the language of the template: by the letters of its sample texts.
  let rtl = 0;
  let ltr = 0;
  for (const text of frame.texts) {
    for (const [letter] of text.matchAll(/\p{L}/gu)) {
      if (RTL.test(letter)) rtl++;
      else ltr++;
    }
  }
  const direction: Direction = rtl > ltr ? 'rtl' : 'ltr';
  const language = direction === 'rtl' ? (frame.texts.some((text) => /\p{Script=Arabic}/u.test(text)) && !frame.texts.some((text) => /\p{Script=Hebrew}/u.test(text)) ? 'ar' : 'he') : 'en';

  // The deck's own texts in the frame: one field for each distinct sample.
  const fields: Record<string, string> = {};
  const fieldOf = new Map<string, string>();
  const fieldFor = (text: string): string => {
    let name = fieldOf.get(text);
    if (!name) {
      name = `field${fieldOf.size + 1}`;
      fieldOf.set(text, name);
      fields[name] = text;
    }
    return name;
  };
  for (const text of frame.fieldTexts) fieldFor(text);

  // The size running text has: the size most characters of the template are set at, titles aside.
  const weight = new Map<number, number>();
  for (const slide of info.slides) {
    for (const shape of slide.shapes) {
      if (shape.kind !== 'text' || isFramePlaceholder(shape) || inFrameBand(shape.box)) continue;
      const size = shape.fontSizes?.[0];
      if (size) weight.set(size, (weight.get(size) ?? 0) + shapeText(shape).length);
    }
  }
  const bodySize = [...weight].sort((a, b) => b[1] - a[1] || a[0] - b[0])[0]?.[0] ?? info.defaultFontSize;

  const patterns: PatternDef[] = [];
  for (const slide of info.slides) {
    const slots: SlotDef[] = [];
    const used = new Set<string>();
    const name = (base: string): string => {
      for (let n = 1; ; n++) {
        const candidate = n === 1 && (base === 'title' || base === 'subtitle') ? base : `${base}${n}`;
        if (!used.has(candidate)) {
          used.add(candidate);
          return candidate;
        }
      }
    };
    const texts = slide.shapes.filter((shape) => shape.kind === 'text' && !isFramePlaceholder(shape));
    const content = texts.filter((shape) => !frame.fieldTexts.includes(shapeText(shape)) && (shapeText(shape) !== '' || shape.placeholder !== undefined));

    // The title: a title placeholder, or else the largest text whose middle is in the upper part
    // of the slide; of several alike, the one that stands highest.
    const sizeOf = (shape: ShapeInfo) => shape.fontSizes?.[0] ?? info.defaultFontSize;
    const placeholderTitle = content.find((shape) => shape.placeholder === 'title' || shape.placeholder === 'ctrTitle');
    // A title is a shape of its own: one inside a group moves with its group.
    const upper = content.filter((shape) => {
      if (shape.groupId !== undefined) return false;
      const box = placeOf(slide, shape);
      return box.y + box.h / 2 < 0.5 && shapeText(shape).length <= 120;
    });
    // A slide with nothing up there (a cover, a closing slide) has its title wherever its largest
    // text is, when that is clearly larger than running text and is no figure.
    const anywhere = content.filter((shape) => shape.groupId === undefined && sizeOf(shape) >= 1.5 * bodySize && shapeText(shape).length <= 120 && !FIGURE.test(shapeText(shape)));
    const largest = (list: ShapeInfo[]) => list.slice().sort((a, b) => sizeOf(b) - sizeOf(a) || placeOf(slide, a).y - placeOf(slide, b).y)[0];
    const title = placeholderTitle ?? largest(upper.filter((shape) => !FIGURE.test(shapeText(shape)))) ?? largest(anywhere);
    const titleBox = title ? placeOf(slide, title) : undefined;
    // The line under it: the nearest text below the title, beside nothing, smaller than it.
    const subtitle =
      title && titleBox
        ? content
            .filter((shape) => shape !== title && sizeOf(shape) < sizeOf(title) && (shape.paragraphs ?? []).filter(Boolean).length <= 1 && shapeText(shape).length <= 140)
            .map((shape) => ({ shape, box: placeOf(slide, shape) }))
            .filter(({ box }) => box.y >= titleBox.y + titleBox.h * 0.5 && box.y <= bottom(titleBox) + 0.1 && box.x < right(titleBox) && right(box) > titleBox.x)
            .sort((a, b) => a.box.y - b.box.y)[0]?.shape
        : undefined;

    // A short line with a longer text right under it, in its column, is the heading of that text.
    const headed = new Set<string>();
    for (const head of content) {
      const text = shapeText(head);
      if (head === title || head === subtitle || text.length === 0 || text.length > 40 || (head.paragraphs ?? []).filter(Boolean).length > 1 || FIGURE.test(text)) continue;
      const box = head.box;
      const under = content.some((other) => {
        if (other === head || other === title || shapeText(other).length <= text.length) return false;
        const overlap = Math.min(right(box), right(other.box)) - Math.max(box.x, other.box.x);
        return other.box.y >= bottom(box) - 0.01 && other.box.y - bottom(box) <= 0.05 && overlap >= 0.6 * Math.min(box.w, other.box.w) && sizeOf(other) <= sizeOf(head);
      });
      if (under) headed.add(head.id);
    }

    for (const shape of slide.shapes) {
      if (shape.kind === 'group') continue;
      // Pictures and tables inside a group stay with their group; its texts are filled where they are.
      if (shape.groupId !== undefined && shape.kind !== 'text') continue;
      if (shape.kind === 'text') {
        const text = shapeText(shape);
        if (isFramePlaceholder(shape)) {
          // A footer, a header or a date that the slide writes out is the deck's own text too; the page number is the software's.
          if (shape.placeholder !== 'sldNum' && text !== '') slots.push({ name: name('frame'), kind: 'text', shapeId: shape.id, deckField: fieldFor(text), empty: 'remove' });
          continue;
        }
        const field = fieldOf.get(text);
        if (field) {
          // A text of the deck's own that the frame shows: filled from the deck's fields, and taken off where the deck has none.
          slots.push({ name: name('frame'), kind: 'text', shapeId: shape.id, deckField: field, empty: 'remove' });
          continue;
        }
        if (!content.includes(shape)) continue;
        if (shape === title && shape.groupId === undefined) {
          const slot = slotFromShape(shape, info, 'title', name('title'), false, slide);
          if (slot) slots.push(slot);
          continue;
        }
        if (shape === subtitle) {
          const slot = slotFromShape(shape, info, 'subtitle', name('subtitle'), true, slide);
          if (slot) slots.push(slot);
          continue;
        }
        const paragraphs = (shape.paragraphs ?? []).filter((paragraph) => paragraph.trim() !== '');
        const size = sizeOf(shape);
        // A big number: a few characters with a digit among them, set far larger than running text ("87%", "2.4M", "6 min").
        const figure = /\d/.test(text) && ((text.length <= 8 && FIGURE.test(text) && size >= 1.5 * bodySize) || (text.length <= 10 && size >= 2 * bodySize));
        const heading = !figure && headed.has(shape.id);
        const role = figure ? 'number' : paragraphs.length >= 2 ? 'list' : 'text';
        const slot = slotFromShape(shape, info, role, name(heading ? 'heading' : role), true, slide);
        if (!slot || slot.kind !== 'text') continue;
        // A figure holds a figure, and a heading a line, whatever width their boxes have.
        if (figure) slot.maxChars = Math.min(slot.maxChars ?? 8, 8);
        else if (heading || (role === 'text' && paragraphs.length <= 1 && text.length <= 40 && size >= bodySize && (shape.box.h * info.slideHeight) / 12700 < 2.6 * size)) slot.maxChars = Math.min(slot.maxChars ?? 60, 60);
        slots.push(slot);
      } else if (shape.kind === 'picture') {
        if (frame.isFrame(shape) || isPictogram(shape, info)) continue;
        // A photograph keeps the template's own picture until a deck gives it another.
        if (shape.box.w * shape.box.h >= 0.03) slots.push({ name: name('photo'), kind: 'image', imageKind: 'photo', shapeId: shape.id, empty: 'keep' });
      } else if (shape.kind === 'table') {
        const slot = slotFromShape(shape, info, 'table', name('table'), true, slide);
        if (slot) slots.push(slot);
      }
    }
    if (slots.length === 0) continue;
    patterns.push({ id: `slide${slide.number}`, sourceSlide: slide.number, description: `Slide ${slide.number} of the template, as it is drawn`, slots });
  }

  // The cover and the closing slide: the template's first and last slide, where each has a title.
  const titled = (pattern: PatternDef | undefined) => pattern?.slots.some((slot) => slot.name === 'title') === true;
  const first = patterns.find((pattern) => pattern.sourceSlide === 1);
  const last = patterns.find((pattern) => pattern.sourceSlide === info.slides.length);
  const result: StyleManifest = { manifest: { version: 1, name: options.name ?? 'style', language, direction, patterns }, patterns: [], fields, notes };
  if (first && titled(first)) {
    first.id = 'cover';
    first.description = 'The opening slide: the title of the deck and one line under it';
    result.cover = 'cover';
  } else notes.push('the first slide has no text that reads as a title, so the deck has no cover of the template');
  if (last && last !== first && titled(last)) {
    last.id = 'closing';
    last.description = 'The closing slide: a few words and one line under them';
    result.closing = 'closing';
  } else notes.push('the last slide has no text that reads as a title, so the deck has no closing slide of the template');
  // The frame of the slides the software lays out: of the slides whose title sits where most
  // titles sit, the one whose room under the title is clearest. A slide can show a drawing there
  // that is no shape of its own (it is drawn by its slide layout) or that repeats on many slides,
  // and both would stay under everything that is laid out: a brush stroke behind three cards.
  const titleBoxOf = (pattern: PatternDef): Box | undefined => {
    const slot = pattern.slots.find((entry) => entry.name === 'title');
    const slide = info.slides[pattern.sourceSlide - 1];
    const shape = slot && slide?.shapes.find((candidate) => candidate.id === slot.shapeId);
    return slide && shape ? placeOf(slide, shape) : undefined;
  };
  const keyOf = (box: Box) => [box.x, box.y, box.w, box.h].map((value) => Math.round(value * 50)).join('|');
  const frames = patterns.filter((pattern) => pattern !== first && pattern !== last).flatMap((pattern) => {
    const box = titleBoxOf(pattern);
    // A title low on the slide leaves no room under it: such a slide is no frame.
    return box && bottom(box) <= 0.45 ? [{ pattern, box, key: keyOf(box) }] : [];
  });
  const counts = new Map<string, number>();
  for (const entry of frames) counts.set(entry.key, (counts.get(entry.key) ?? 0) + 1);
  const most = Math.max(0, ...counts.values());
  const drawn = new Map<string, Box[]>();
  const layoutDraws = (part: string): Box[] => {
    let boxes = drawn.get(part);
    if (!boxes) drawn.set(part, (boxes = part && pkg.has(part) ? drawnBoxes(rootElement(parseXml(pkg.readText(part))), info) : []));
    return boxes;
  };
  const clutter = (pattern: PatternDef, title: Box): number => {
    const slide = info.slides[pattern.sourceSlide - 1] as SlideInfo;
    const room: Box = { x: 0.03, y: bottom(title), w: 0.94, h: Math.max(0, 0.9 - bottom(title)) };
    const filled = new Set(pattern.slots.map((slot) => slot.shapeId));
    const stays = slide.shapes.filter((shape) => shape.kind !== 'text' && shape.groupId === undefined && !filled.has(shape.id) && frame.isFrame(shape)).map((shape) => shape.box);
    return [...layoutDraws(slide.layout), ...stays].filter((box) => !(box.w >= 0.95 && box.h >= 0.95)).reduce((sum, box) => sum + overlap(box, room), 0);
  };
  // A page that is one colour but for a glow in a corner is the template's own look, and fills a
  // corner that a low row leaves empty: of frames that are clear alike, one with such a page is
  // taken before a plain one, and the calmest of them first (the one most of whose page is plain).
  const glowOf = (pattern: PatternDef): number => {
    const slide = info.slides[pattern.sourceSlide - 1] as SlideInfo;
    return pageColor(pkg, slide.part, rootElement(parseXml(pkg.readText(slide.part)))).glow ?? 0;
  };
  const best = frames
    .filter((entry) => counts.get(entry.key) === most)
    .map((entry) => ({ ...entry, clutter: clutter(entry.pattern, entry.box), glow: glowOf(entry.pattern), shapes: (info.slides[entry.pattern.sourceSlide - 1] as SlideInfo).shapes.length }))
    .sort((a, b) => a.clutter - b.clutter || b.glow - a.glow || a.shapes - b.shapes || a.pattern.sourceSlide - b.pattern.sourceSlide)[0];
  if (best) {
    result.manifest.free = { base: best.pattern.id };
    notes.push(`free slides stand on the frame of slide ${best.pattern.sourceSlide}: its title sits where the titles of ${most} slides sit${best.clutter > 0.002 ? `, and ${Math.round(best.clutter * 100)}% of the slide under it is drawn on by its layout` : ', and nothing is drawn under it'}${best.glow > 0 ? `; its page is one colour over ${Math.round(best.glow * 100)}% of it, with the template's glow in the rest` : ''}`);
  } else {
    // With no such slide the engine would stand them on a cover, under a title in the middle of the slide: no place for content.
    result.manifest.free = false;
    result.reason = 'no slide has a title in its upper part with room under it: slides the software lays out need one slide with a title at the top and content under it';
  }
  // The template's colours are found by how much its slides use each. A deck whose slides draw
  // nothing in a colour (titles and bullets alone) has its colours in its theme: the first two accents.
  const first1 = info.slides[0];
  if (first1 && patterns.length > 0 && !paletteOf(pkg, { patterns }).accent) {
    const scheme = schemeOf(pkg, first1.part);
    const accent = scheme.resolve('accent1')?.toUpperCase();
    const accent2 = scheme.resolve('accent2')?.toUpperCase();
    if (accent) {
      result.manifest.colors = { accent, ...(accent2 && accent2 !== accent ? { accent2 } : {}) };
      notes.push(`no slide draws anything in a colour, so the main colour is the first accent of the theme, ${accent}`);
    }
  }
  // A template drawn with small running text (an agency's 9 pt) is drawn for a designer's eye. The
  // slides the software lays out carry a person's own words, to be read: their running text is
  // enlarged to what 14 pt is on a slide 13.33 inches wide, by no more than half.
  const comfortable = 14 * (info.slideWidth / 12192000);
  if (bodySize < comfortable - 0.25) {
    result.manifest.textScale = Math.min(1.5, Math.round((comfortable / bodySize) * 20) / 20);
    notes.push(`running text is ${bodySize} pt on a slide ${Math.round((info.slideWidth / 914400) * 100) / 100} inches wide: it is enlarged ${result.manifest.textScale} times on every slide`);
  }
  notes.push(`${patterns.length} of ${info.slides.length} slides were read; running text is set at ${bodySize} pt; the template reads ${direction === 'rtl' ? 'right to left' : 'left to right'}`);
  if (inherits > 0) notes.push(`${inherits} texts are placeholders that take their place or their size from the slide layout: they were read where they stand and at the size they are set`);
  if (frame.often === Infinity) notes.push('the template has fewer than six slides: too few to tell what repeats, so nothing but page numbers and footers is taken as its frame');
  // Whether the template can carry the slides the software lays out is the engine's to say: it
  // needs a frame to stand on and a text style it can copy.
  if (!result.reason) {
    const kit = patterns.length > 0 ? freeKitOrReason({ pkg, manifest: result.manifest }) : 'no slide of the file has a text to take a style from';
    if (typeof kit === 'string') result.reason = kit;
  }
  if (result.reason) {
    result.manifest.free = false;
    notes.push(`the template cannot carry slides the software lays out: ${result.reason}`);
  }
  result.patterns = [result.cover, result.reason ? undefined : 'free', result.closing].filter((id): id is string => id !== undefined);
  return result;
}

const overlap = (a: Box, b: Box): number => Math.max(0, Math.min(right(a), right(b)) - Math.max(a.x, b.x)) * Math.max(0, Math.min(bottom(a), bottom(b)) - Math.max(a.y, b.y));

/**
 * Where a slide layout draws something of its own: the boxes of its pictures, shapes and groups
 * that are no placeholders, as fractions of the slide. Every slide on the layout shows them, and
 * no slide can take them off.
 */
function drawnBoxes(root: Element, info: TemplateInfo): Box[] {
  const tree = path(root, [NS.p, 'cSld'], [NS.p, 'spTree']);
  if (!tree) return [];
  const boxes: Box[] = [];
  for (const child of childElements(tree)) {
    if (!['sp', 'pic', 'grpSp', 'cxnSp', 'graphicFrame'].includes(child.localName ?? '')) continue;
    const nonVisual = childElements(child).find((entry) => /^nv\w+Pr$/.test(entry.localName ?? ''));
    const application = nonVisual && childElement(nonVisual, NS.p, 'nvPr');
    if (application && childElement(application, NS.p, 'ph')) continue;
    const properties = childElement(child, NS.p, child.localName === 'grpSp' ? 'grpSpPr' : 'spPr');
    const transform = properties ? childElement(properties, NS.a, 'xfrm') : child.localName === 'graphicFrame' ? childElement(child, NS.p, 'xfrm') : undefined;
    const off = transform && childElement(transform, NS.a, 'off');
    const ext = transform && childElement(transform, NS.a, 'ext');
    if (!off || !ext) continue;
    const box = { x: Number(off.getAttribute('x') ?? 0) / info.slideWidth, y: Number(off.getAttribute('y') ?? 0) / info.slideHeight, w: Number(ext.getAttribute('cx') ?? 0) / info.slideWidth, h: Number(ext.getAttribute('cy') ?? 0) / info.slideHeight };
    if (box.w > 0 && box.h > 0) boxes.push(box);
  }
  return boxes;
}
