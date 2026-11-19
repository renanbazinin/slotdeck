// A new layout made from the template's own parts. The model designs it on a grid laid over the
// template's content area, with the template's own text styles and ornaments; the code builds a
// real slide from that: the frame of a slide the template already has (background, header, logo,
// page number, its title where every title is), and copies of the template's own shapes in the
// cells the design gives them. The slide is added at the end of the pptx and becomes a pattern.
import type { LlmProvider } from '../llm/provider.ts';
import { ContentTypes, isRelType, nextRelId, readRels, relativeTarget, writeRels } from '../pptx/package.ts';
import type { PptxPackage, Relationship } from '../pptx/package.ts';
import { readPresentation } from '../pptx/presentation.ts';
import { NS, childElement, childElements, descendants, parseXml, path, removeNode, rootElement, serializeXml } from '../pptx/xml.ts';
import type { Document, Element } from '../pptx/xml.ts';
import { keepCorners } from '../render/repeat.ts';
import { indexShapes, loadTemplate, renderDeck } from '../render/render.ts';
import type { Template } from '../render/render.ts';
import { setShapeText } from '../render/text.ts';
import type { TextOptions } from '../render/text.ts';
import { TextStyles } from '../text/styles.ts';
import { changeWith } from './change.ts';
import type { ManifestChange } from './change.ts';
import { checkManifest } from './check.ts';
import { inspectTemplate } from './inspect.ts';
import type { Box, ShapeInfo, SlideInfo, TemplateInfo } from './inspect.ts';
import type { PatternDef, SlotDef, TemplateManifest } from './manifest.ts';
import { measuredBudgets } from './budgets.ts';
import { frameFacts, inFrameBand, isFramePlaceholder, shapeText } from './measure.ts';
import { sampleDeck } from './sample.ts';

/** The columns and rows of the grid a design is drawn on. */
export const DESIGN_GRID = 12;
const GRID = DESIGN_GRID;

/** What a part of a composed layout is for. A decoration is drawn and never filled. */
export const DESIGN_ROLES = ['title', 'subtitle', 'text', 'list', 'number', 'caption', 'decoration'] as const;
export type DesignRole = (typeof DESIGN_ROLES)[number];

/** One part of a composed layout: a copy of one of the template's shapes, in some cells of the grid. */
export interface DesignElement {
  /** The slot's name, such as `point1Heading`. Empty for a decoration. */
  name: string;
  /** A text style (`T1`, `T2`, ...) or a decoration (`D1`, ...) of the template, as {@link designKit} lists them. */
  style: string;
  role: DesignRole;
  /** A deck may leave the text out; its shape is then removed. */
  optional: boolean;
  /** Sample text in the template's language, paragraphs separated by line breaks. Empty for a decoration. */
  sample: string;
  /** The first column and row it takes, from 1, and how many columns and rows. Column 1 is at the slide's left edge. */
  column: number;
  row: number;
  width: number;
  height: number;
}

/** A layout as the model designs it. */
export interface PatternDesign {
  /** Id of the new pattern, such as `three-points`. */
  id: string;
  /** What the layout is for, written for the planner that chooses layouts. */
  description: string;
  /** Slots of the base slide that stay where they are, such as `title`. */
  keep: string[];
  /**
   * The layout whose slide gives the frame and the title. Left out, the code picks the one whose
   * title sits where most titles do. Recorded, so that the design is built on the same frame again.
   */
  base?: string;
  /** In drawing order: a decoration listed before a text lies behind it. */
  elements: DesignElement[];
}

/** A text style or decoration of the template that a design can use. */
export interface KitStyle {
  key: string;
  kind: 'text' | 'decoration';
  /** Where its shape is: the composed shape is a copy of it. */
  slide: number;
  shapeId: string;
  /** One line about it, as the model reads it. */
  text: string;
  /** Text: font size in points, and how many grid rows one line takes. */
  size?: number;
  lineRows?: number;
  bullet?: boolean;
  /** Decoration: a rule across, a rule down, or an area such as a card. */
  shape?: 'across' | 'down' | 'area';
  /** Where its shape is drawn, as fractions of the slide. */
  drawn: Box;
}

/** What a design is drawn with: the grid over the template's content area, the base slide, the styles. */
export interface DesignKit {
  /** The pattern whose slide gives the frame: background, header, logo, page number, title. */
  base: string;
  baseSlide: number;
  /** The content area the grid covers, as fractions of the slide. */
  area: Box;
  /** Slots of the base slide that a design may keep where they are, with the cells they cover. */
  keepable: Array<{ slot: string; role: string; cells: { column: number; row: number; width: number; height: number } }>;
  /** Cells taken by what stays on every slide (header, logo, page number): nothing may be put there. */
  taken: boolean[][];
  styles: KitStyle[];
}

export interface ComposeOptions {
  llm: LlmProvider;
  /**
   * The layout to stand on, such as the one the new layout replaces in a deck: its frame and its
   * title stay. It needs a title and no row of items; otherwise the code picks one as usual.
   */
  base?: string;
  /** `low` asks the model for light reasoning: much faster, and it was enough in the runs measured. */
  effort?: 'low';
  signal?: AbortSignal;
}

export interface ComposedPattern {
  /** The template pptx with the new slide added after the last one. Store it in place of the template's. */
  pptx: Uint8Array;
  /** The manifest with the new pattern added at the end. */
  manifest: TemplateManifest;
  pattern: string;
  /** The new slide's number in the pptx. */
  slide: number;
  /** What was made, in a sentence or two. */
  summary: string;
  /** The design it was built from. `buildPattern` builds it again. */
  design: PatternDesign;
  /** What the code noticed and allowed, such as a sample text longer than its box holds. */
  warnings: string[];
}

/** A design that cannot be built, with every reason. */
export class DesignError extends Error {
  readonly problems: string[];
  constructor(problems: string[]) {
    super(`The layout cannot be built: ${problems.join(' ')}`);
    this.name = 'DesignError';
    this.problems = problems;
  }
}

// ---------------------------------------------------------------------------
// The kit

const EMU_PER_INCH = 914400;
const round = (value: number) => Math.round(value * 10000) / 10000;
const right = (box: Box) => box.x + box.w;
const bottom = (box: Box) => box.y + box.h;
const overlap = (a: Box, b: Box) => Math.max(0, Math.min(right(a), right(b)) - Math.max(a.x, b.x)) * Math.max(0, Math.min(bottom(a), bottom(b)) - Math.max(a.y, b.y));
const median = (values: number[]) => {
  const sorted = values.slice().sort((a, b) => a - b);
  return sorted.length === 0 ? Number.NaN : (sorted[Math.floor((sorted.length - 1) / 2)] as number);
};
const key = (box: Box) => `${Math.round(box.x * 50)}|${Math.round(box.y * 50)}|${Math.round(box.w * 50)}|${Math.round(box.h * 50)}`;
const hasRelationship = (element: Element) =>
  [element, ...Array.from(element.getElementsByTagNameNS('*', '*'))].some((el) => Array.from({ length: el.attributes.length }, (_, i) => el.attributes.item(i)).some((attr) => attr?.namespaceURI === NS.r));

/** The box a slot shows its shape in: its own box when the pattern moves it. */
function slotBox(slot: SlotDef, shape: ShapeInfo): Box {
  return 'box' in slot && slot.box ? slot.box : shape.box;
}

const isTitle = (slot: SlotDef) => slot.kind === 'text' && !slot.deckField && /^(title|heading|slideTitle)$/i.test(slot.name);

/** Picks the slide whose frame and title a composed layout stands on: the one whose title sits where most titles do. */
function pickBase(manifest: TemplateManifest, info: TemplateInfo, wanted?: string): { pattern: PatternDef; slide: SlideInfo; title: SlotDef } | undefined {
  // A layout asked for by name may move its shapes: its title is kept where that layout shows it.
  const asked = wanted ? manifest.patterns.find((pattern) => pattern.id === wanted) : undefined;
  if (asked && !asked.repeat) {
    const slide = info.slides[asked.sourceSlide - 1];
    const title = asked.slots.find(isTitle);
    if (slide && title && slide.shapes.some((shape) => shape.id === title.shapeId)) return { pattern: asked, slide, title };
  }
  const candidates: Array<{ pattern: PatternDef; slide: SlideInfo; title: SlotDef; box: Box }> = [];
  for (const pattern of manifest.patterns) {
    const slide = info.slides[pattern.sourceSlide - 1];
    const title = pattern.slots.find(isTitle);
    const shape = title && slide?.shapes.find((candidate) => candidate.id === title.shapeId);
    // A layout that moves its shapes is a variant: its slide shows them elsewhere.
    if (!slide || !title || !shape || pattern.repeat || pattern.slots.some((slot) => 'box' in slot && slot.box)) continue;
    candidates.push({ pattern, slide, title, box: slotBox(title, shape) });
  }
  if (candidates.length === 0) return undefined;
  const counts = new Map<string, number>();
  for (const candidate of candidates) counts.set(key(candidate.box), (counts.get(key(candidate.box)) ?? 0) + 1);
  const content = (slide: SlideInfo) => slide.shapes.filter((shape) => shape.kind !== 'group').length;
  candidates.sort((a, b) => (counts.get(key(b.box)) ?? 0) - (counts.get(key(a.box)) ?? 0) || content(a.slide) - content(b.slide) || a.slide.number - b.slide.number);
  const best = candidates[0] as (typeof candidates)[number];
  return { pattern: best.pattern, slide: best.slide, title: best.title };
}

/**
 * Shape ids of a slide that stay on a composed slide whatever the design: the frame and the deck's
 * own texts. And the shapes the grid is laid around (see {@link contentArea}): those that stayed in
 * 0.5.0, when a sample the template repeats under its titles stayed too, so that a layout made
 * then is built in the same place after the manifest was put right.
 */
function staying(base: PatternDef, slide: SlideInfo, info: TemplateInfo): { stays: Set<string>; laid: Set<string> } {
  const frame = frameFacts(info);
  const stays = new Set<string>();
  const laid = new Set<string>();
  // A text the pattern fills is content, even when its sample is the same on many slides.
  const filled = new Set(base.slots.filter((slot) => slot.kind === 'text' && !slot.deckField && slot.empty !== 'drop').map((slot) => slot.shapeId));
  for (const shape of slide.shapes) {
    const text = shape.kind === 'text' ? shapeText(shape) : '';
    const repeated = frame.repeatedTexts.includes(text);
    // A sample repeated in the content, such as "Subtitle", which a draft made a deck field and prepareManifest a slot.
    if (repeated && !frame.fieldTexts.includes(text)) laid.add(shape.id);
    if (filled.has(shape.id)) continue;
    if (frame.isFrame(shape) || repeated) laid.add(shape.id);
    if (frame.isFrame(shape) || frame.fieldTexts.includes(text)) stays.add(shape.id);
  }
  // A deck field stays where the frame shows it. One in the content is a sample the template
  // repeats under its titles, which a draft made a field: it goes with the rest of the content.
  for (const slot of base.slots) {
    const shape = slot.kind === 'text' && slot.deckField ? slide.shapes.find((candidate) => candidate.id === slot.shapeId) : undefined;
    if (!shape) continue;
    laid.add(slot.shapeId);
    if (inFrameBand(shape.box)) stays.add(slot.shapeId);
  }
  return { stays, laid };
}

/**
 * Whether a thin shape of the frame runs through the content area, as a divider beside a column of
 * titles does on some slides. It belongs to how those slides lay out their content: the grid does
 * not see a line, so the parts of a new layout would lie across it.
 */
function runsThrough(shape: ShapeInfo, area: Box): boolean {
  if (shape.kind !== 'line' && shape.kind !== 'shape') return false;
  const margin = 0.01;
  const along = (from: number, size: number, start: number, length: number) => Math.min(from + size, start + length) - Math.max(from, start);
  if (shape.box.w <= 0.005) {
    const x = shape.box.x + shape.box.w / 2;
    return x > area.x + margin && x < right(area) - margin && along(shape.box.y, shape.box.h, area.y, area.h) > area.h * 0.25;
  }
  if (shape.box.h <= 0.005) {
    const y = shape.box.y + shape.box.h / 2;
    return y > area.y + margin && y < bottom(area) - margin && along(shape.box.x, shape.box.w, area.x, area.w) > area.w * 0.25;
  }
  return false;
}

/** The area the composed content may use: within the margins the template's content keeps, under the header and above the footer. */
function contentArea(manifest: TemplateManifest, info: TemplateInfo, slide: SlideInfo, stays: Set<string>, title: Box): Box {
  const lefts: number[] = [];
  const rights: number[] = [];
  const used = new Set(manifest.patterns.map((pattern) => pattern.sourceSlide));
  const frame = frameFacts(info);
  for (const number of used) {
    const shapes = (info.slides[number - 1]?.shapes ?? []).filter(
      (shape) => shape.kind !== 'group' && !frame.isFrame(shape) && !isFramePlaceholder(shape) && shape.box.w > 0 && shape.box.w < 0.9 && shape.box.x > 0.005 && right(shape.box) < 0.995,
    );
    if (shapes.length === 0) continue;
    lefts.push(Math.min(...shapes.map((shape) => shape.box.x)));
    rights.push(Math.max(...shapes.map((shape) => right(shape.box))));
  }
  const x = Math.max(0.02, Math.min(median(lefts) || 0.06, title.w > 0 ? title.x : 1));
  const xEnd = Math.min(0.98, Math.max(median(rights) || 0.94, title.w > 0 ? right(title) : 0));
  let top = 0.04;
  let end = 0.95;
  for (const shape of slide.shapes) {
    if (!stays.has(shape.id) || shape.kind === 'group' || shape.box.w >= 0.98 && shape.box.h >= 0.98) continue;
    if (bottom(shape.box) <= 0.3) top = Math.max(top, bottom(shape.box) + 0.02);
    else if (shape.box.y >= 0.7) end = Math.min(end, shape.box.y - 0.02);
  }
  if (title.w > 0) top = Math.min(top, title.y);
  return { x: round(x), y: round(top), w: round(xEnd - x), h: round(end - top) };
}

interface Cells {
  column: number;
  row: number;
  width: number;
  height: number;
}

function cellsOf(box: Box, area: Box): Cells | undefined {
  const cw = area.w / GRID;
  const ch = area.h / GRID;
  const c0 = Math.max(0, Math.floor((box.x - area.x) / cw + 0.25));
  const c1 = Math.min(GRID, Math.ceil((right(box) - area.x) / cw - 0.25));
  const r0 = Math.max(0, Math.floor((box.y - area.y) / ch + 0.25));
  const r1 = Math.min(GRID, Math.ceil((bottom(box) - area.y) / ch - 0.25));
  if (c1 <= c0 || r1 <= r0) return undefined;
  return { column: c0 + 1, row: r0 + 1, width: c1 - c0, height: r1 - r0 };
}

export interface ShapeLook {
  element: Element;
  fill: string;
  line: string;
  geometry: string;
  align: string;
  color: string;
}

const colorOf = (parent: Element | undefined): string => {
  const fill = parent && childElement(parent, NS.a, 'solidFill');
  const color = fill && childElements(fill)[0];
  return color ? (color.getAttribute('val') ?? '') : '';
};

export function lookOf(element: Element): ShapeLook {
  const spPr = childElement(element, NS.p, 'spPr');
  const ln = spPr && childElement(spPr, NS.a, 'ln');
  const run = descendants(element, NS.a, 'rPr')[0] ?? descendants(element, NS.a, 'endParaRPr')[0];
  const pPr = descendants(element, NS.a, 'pPr')[0];
  return {
    element,
    fill: spPr && childElement(spPr, NS.a, 'solidFill') ? colorOf(spPr) || 'theme colour' : '',
    line: ln && !childElement(ln, NS.a, 'noFill') && childElement(ln, NS.a, 'solidFill') ? `${colorOf(ln) || 'theme colour'}${ln.getAttribute('w') ? ` ${Math.round((Number(ln.getAttribute('w')) / 12700) * 4) / 4} pt` : ''}` : '',
    geometry: (spPr && childElement(spPr, NS.a, 'prstGeom')?.getAttribute('prst')) ?? (element.localName === 'cxnSp' ? 'line' : 'custom'),
    align: pPr?.getAttribute('algn') ?? '',
    color: colorOf(run),
  };
}

const ALIGN: Record<string, string> = { l: 'left-aligned', r: 'right-aligned', ctr: 'centred', just: 'justified' };

/**
 * What a layout can be composed of, read from the template: the grid over its content area, the
 * slide whose frame it stands on, and every distinct text style and ornament its layouts use. No
 * model, and nothing is changed. `undefined` when the template has no layout with a title to
 * stand on.
 */
export function designKit(template: Template, inspected?: TemplateInfo, base?: string): DesignKit | undefined {
  return kitWith(template, inspected ?? inspectTemplate(template.pkg), base)?.kit;
}

export interface KitParts {
  kit: DesignKit;
  info: TemplateInfo;
  base: PatternDef;
  slide: SlideInfo;
  stays: Set<string>;
}

/**
 * The frame a new slide stands on, as a composed layout has it: the layout whose slide gives it,
 * the shapes of that slide that stay (frame and deck's own texts), and the kit. For a slide that
 * the software lays out itself; `undefined` when the template has no layout with a title.
 */
export function frameParts(template: Template, base?: string, inspected?: TemplateInfo): KitParts | undefined {
  return kitWith(template, inspected ?? inspectTemplate(template.pkg), base);
}

/** The name a composed slide carries in the pptx, so that a kit leaves it out. */
const COMPOSED = 'composed ';
const COMPOSED_SLIDE = /<(?:\w+:)?cSld\b[^>]*\bname="composed /;

/** Numbers of the slides that were composed: they copy the template's styles and would only repeat them. */
function composedSlides(pkg: PptxPackage): Set<number> {
  const out = new Set<number>();
  readPresentation(pkg).slides.forEach((ref, index) => {
    if (pkg.has(ref.part) && COMPOSED_SLIDE.test(pkg.readText(ref.part))) out.add(index + 1);
  });
  return out;
}

function kitWith(template: Template, info: TemplateInfo, wanted?: string): KitParts | undefined {
  // Layouts composed earlier are left out, so that composing one more never changes the kit: a
  // design made for the template is built the same way after another layout was added.
  const composed = composedSlides(template.pkg);
  const manifest = { ...template.manifest, patterns: template.manifest.patterns.filter((pattern) => !composed.has(pattern.sourceSlide)) };
  template = { ...template, manifest };
  const picked = pickBase(manifest, info, wanted);
  if (!picked) return undefined;
  const { pattern: base, slide, title } = picked;
  const { stays, laid } = staying(base, slide, info);
  const titleShape = slide.shapes.find((shape) => shape.id === title.shapeId) as ShapeInfo;
  const area = contentArea(manifest, info, slide, laid, slotBox(title, titleShape));
  for (const shape of slide.shapes) if (stays.has(shape.id) && runsThrough(shape, area)) stays.delete(shape.id);

  const taken = Array.from({ length: GRID }, () => Array.from({ length: GRID }, () => false));
  for (const shape of slide.shapes) {
    if (!stays.has(shape.id) || shape.kind === 'group' || (shape.box.w >= 0.98 && shape.box.h >= 0.98)) continue;
    mark(taken, shape.box, area);
  }
  const keepable: DesignKit['keepable'] = [];
  for (const slot of base.slots) {
    if (slot.kind !== 'text' || slot.deckField || slot.empty === 'drop' || stays.has(slot.shapeId)) continue;
    const shape = slide.shapes.find((candidate) => candidate.id === slot.shapeId);
    if (!shape || !(isTitle(slot) || /subtitle|subheading/i.test(slot.name))) continue;
    const cells = cellsOf(slotBox(slot, shape), area);
    if (cells) keepable.push({ slot: slot.name, role: isTitle(slot) ? 'title' : 'subtitle', cells });
  }

  const styles = kitStyles(template, info, area);
  return { kit: { base: base.id, baseSlide: slide.number, area, keepable, taken, styles }, info, base, slide, stays };
}

function mark(taken: boolean[][], box: Box, area: Box): void {
  const cw = area.w / GRID;
  const ch = area.h / GRID;
  for (let row = 0; row < GRID; row++) {
    for (let column = 0; column < GRID; column++) {
      const cell = { x: area.x + column * cw, y: area.y + row * ch, w: cw, h: ch };
      if (overlap(cell, box) > cw * ch * 0.25) (taken[row] as boolean[])[column] = true;
    }
  }
}

/** Every distinct text style and ornament on the slides the template's patterns use. */
function kitStyles(template: Template, info: TemplateInfo, area: Box): KitStyle[] {
  const { manifest } = template;
  const frame = frameFacts(info);
  const styles = new TextStyles(template.pkg);
  const pres = readPresentation(template.pkg);
  const used = [...new Set(manifest.patterns.map((pattern) => pattern.sourceSlide))].sort((a, b) => a - b);
  const slotNames = new Map<string, string[]>();
  for (const pattern of manifest.patterns) {
    for (const slot of pattern.slots) {
      const at = `${pattern.sourceSlide}|${slot.shapeId}`;
      if (!slotNames.has(at)) slotNames.set(at, []);
      (slotNames.get(at) as string[]).push(`${slot.name} in ${pattern.id}`);
    }
  }
  const deckShapes = new Set(manifest.patterns.flatMap((pattern) => pattern.slots.filter((slot) => slot.kind === 'text' && (slot.deckField || slot.empty === 'drop')).map((slot) => `${pattern.sourceSlide}|${slot.shapeId}`)));
  const rowHeight = (area.h * info.slideHeight) / 12700 / GRID;

  const texts = new Map<string, { style: KitStyle; count: number; uses: Set<string>; sample: string }>();
  const ornaments = new Map<string, { style: KitStyle; count: number; uses: Set<string> }>();
  for (const number of used) {
    const slide = info.slides[number - 1];
    const ref = pres.slides[number - 1];
    if (!slide || !ref) continue;
    const elements = indexShapes(rootElement(parseXml(template.pkg.readText(ref.part))));
    const boxes = slide.shapes.filter((shape) => shape.kind === 'text').map((shape) => shape.box);
    const pictures = slide.shapes.filter((shape) => shape.kind === 'picture').map((shape) => shape.box);
    for (const shape of slide.shapes) {
      const element = elements.get(shape.id);
      if (!element || shape.placeholder !== undefined || frame.isFrame(shape) || hasRelationship(element)) continue;
      const at = `${number}|${shape.id}`;
      const names = slotNames.get(at) ?? [];
      const look = lookOf(element);
      if (shape.kind === 'text' && element.localName === 'sp') {
        if (deckShapes.has(at) || frame.repeatedTexts.includes(shapeText(shape))) continue;
        const text = styles.frame(number, shape.id);
        const paragraph = text?.paragraphs[0];
        if (!text || !paragraph || text.vertical) continue;
        const run = paragraph.run;
        const size = Math.round(run.size * 2) / 2;
        // Body text is shown larger in every deck when the manifest says so.
        const shown = size <= 12 ? Math.round(size * (manifest.textScale ?? 1) * 2) / 2 : size;
        const signature = [size, run.bold, run.italic, run.latin.name, look.color, paragraph.bullet, look.fill, look.line, look.geometry, look.align].join('|');
        const known = texts.get(signature);
        const role = names[0]?.split(' in ')[0] ?? '';
        if (known) {
          known.count++;
          for (const name of names) known.uses.add(name);
          continue;
        }
        const parts = [
          `${size} pt${shown !== size ? ` (shown at ${shown} pt)` : ''}${run.bold ? ' bold' : ''}${run.italic ? ' italic' : ''}`,
          run.latin.name || run.cs.name ? `${run.latin.name || run.cs.name}` : '',
          look.color ? `colour ${look.color}` : '',
          ALIGN[look.align] ?? '',
          paragraph.bullet ? 'with bullets' : '',
          look.fill ? `on a filled box (${look.fill})` : '',
          look.line ? `with an outline (${look.line})` : '',
        ].filter(Boolean);
        texts.set(signature, {
          style: { key: '', kind: 'text', slide: number, shapeId: shape.id, drawn: shape.box, text: parts.join(', '), size, lineRows: Math.round(((shown * 1.2) / rowHeight) * 100) / 100, bullet: paragraph.bullet },
          count: 1,
          uses: new Set(names),
          sample: shapeText(shape).slice(0, 40) || role,
        });
      } else if ((shape.kind === 'shape' || shape.kind === 'line') && (element.localName === 'sp' || element.localName === 'cxnSp')) {
        if (shape.box.w >= 0.9 && shape.box.h >= 0.9) continue;
        const inches = (fraction: number, of: number) => Math.round(((fraction * of) / EMU_PER_INCH) * 100) / 100;
        const w = inches(shape.box.w, info.slideWidth);
        const h = inches(shape.box.h, info.slideHeight);
        const kind: KitStyle['shape'] = h <= 0.06 || w >= h * 12 ? 'across' : w <= 0.06 || h >= w * 12 ? 'down' : 'area';
        if (!look.fill && !look.line && shape.kind !== 'line') continue;
        // A disc or frame behind a pictogram is half of a picture: without one it is an empty shape.
        if (pictures.some((box) => box.x >= shape.box.x - 0.005 && right(box) <= right(shape.box) + 0.005 && box.y >= shape.box.y - 0.005 && bottom(box) <= bottom(shape.box) + 0.005)) continue;
        const holds = kind === 'area' && boxes.some((box) => box.x >= shape.box.x - 0.005 && right(box) <= right(shape.box) + 0.005 && box.y >= shape.box.y - 0.005 && bottom(box) <= bottom(shape.box) + 0.005);
        const signature = [kind, look.geometry, look.fill, look.line, kind === 'area' ? '' : Math.round((kind === 'across' ? h : w) * 100)].join('|');
        const known = ornaments.get(signature);
        const what = kind === 'across' ? 'a rule across' : kind === 'down' ? 'a rule down' : holds ? `a ${look.geometry === 'roundRect' ? 'rounded ' : ''}panel that text sits on, such as a card` : `a ${look.geometry === 'roundRect' ? 'rounded rectangle' : look.geometry === 'ellipse' ? 'circle or oval' : look.geometry}`;
        const use = `on slide ${number}${holds ? ' behind text' : ''}`;
        if (known) {
          known.count++;
          known.uses.add(use);
          continue;
        }
        const parts = [what, look.fill ? `filled ${look.fill}` : '', look.line ? `outline ${look.line}` : '', kind === 'area' ? `${w} x ${h} in as drawn` : ''].filter(Boolean);
        ornaments.set(signature, { style: { key: '', kind: 'decoration', slide: number, shapeId: shape.id, drawn: shape.box, text: parts.join(', '), shape: kind }, count: 1, uses: new Set([use]) });
      }
    }
  }

  const out: KitStyle[] = [];
  const sortedTexts = [...texts.values()].sort((a, b) => (b.style.size ?? 0) - (a.style.size ?? 0) || b.count - a.count).slice(0, 20);
  sortedTexts.forEach((entry, index) => {
    const uses = [...entry.uses].slice(0, 4).join(', ');
    out.push({ ...entry.style, key: `T${index + 1}`, text: `${entry.style.text} | ${uses ? `used for ${uses}` : 'design text'}${entry.count > 1 ? ` (${entry.count} shapes)` : ''} | e.g. ${JSON.stringify(entry.sample)} | one line takes ${entry.style.lineRows} rows` });
  });
  [...ornaments.values()]
    .sort((a, b) => b.count - a.count)
    .slice(0, 12)
    .forEach((entry, index) => out.push({ ...entry.style, key: `D${index + 1}`, text: `${entry.style.text} | ${[...entry.uses].slice(0, 3).join(', ')}${entry.count > 1 ? ` (${entry.count} shapes)` : ''}` }));
  return out;
}

// ---------------------------------------------------------------------------
// Checking a design

interface Placed {
  element: DesignElement;
  style: KitStyle;
  box: Box;
  cells: Cells;
}

const inside = (inner: Cells, outer: Cells) => inner.column >= outer.column && inner.row >= outer.row && inner.column + inner.width <= outer.column + outer.width && inner.row + inner.height <= outer.row + outer.height;
const cellsOverlap = (a: Cells, b: Cells) => a.column < b.column + b.width && b.column < a.column + a.width && a.row < b.row + b.height && b.row < a.row + a.height;
const describeCells = (cells: Cells) => `columns ${cells.column}-${cells.column + cells.width - 1}, rows ${cells.row}-${cells.row + cells.height - 1}`;
/** Space between neighbouring parts, and between a panel's edge and the text on it, as shares of the slide's width. */
const GAP = 0.012;
const PAD = 0.012;

/**
 * The box of some cells, kept a little apart from its neighbours; a text on a panel keeps a margin
 * from the panel's edges. `aspect` is the slide's width over its height, so that a gap is as wide
 * as it is high.
 */
function boxOf(cells: Cells, area: Box, aspect: number, panel?: Cells): Box {
  const cw = area.w / GRID;
  const ch = area.h / GRID;
  const gapX = GAP / 2;
  const gapY = (GAP / 2) * aspect;
  let x = area.x + (cells.column - 1) * cw;
  let y = area.y + (cells.row - 1) * ch;
  let xEnd = x + cells.width * cw;
  let yEnd = y + cells.height * ch;
  if (cells.column > 1) x += gapX;
  if (cells.column + cells.width - 1 < GRID) xEnd -= gapX;
  if (cells.row > 1) y += gapY;
  if (cells.row + cells.height - 1 < GRID) yEnd -= gapY;
  if (panel) {
    const padX = PAD;
    const padY = PAD * aspect;
    if (cells.column === panel.column) x += padX;
    if (cells.column + cells.width === panel.column + panel.width) xEnd -= padX;
    if (cells.row === panel.row) y += padY;
    if (cells.row + cells.height === panel.row + panel.height) yEnd -= padY;
  }
  return { x: round(x), y: round(y), w: round(xEnd - x), h: round(yEnd - y) };
}

const NAME = /^[a-z][A-Za-z0-9]{0,47}$/;

/** What is wrong with a design, as sentences the model can act on; and where each part goes when nothing is. */
function review(design: PatternDesign, parts: KitParts, manifest: TemplateManifest): { problems: string[]; soft: string[]; placed: Placed[] } {
  const { kit } = parts;
  const problems: string[] = [];
  const soft: string[] = [];
  const styles = new Map(kit.styles.map((style) => [style.key, style]));
  if (!/^[a-z0-9][a-z0-9-]{0,47}$/.test(design.id)) problems.push(`The id "${design.id}" must be lower case letters, digits and hyphens.`);
  else if (manifest.patterns.some((pattern) => pattern.id === design.id)) problems.push(`The template already has a layout called "${design.id}"; choose another id.`);
  if (!design.description.trim()) problems.push('The description is empty; say what content the layout suits.');
  const keepable = new Map(kit.keepable.map((entry) => [entry.slot, entry]));
  for (const name of design.keep) if (!keepable.has(name)) problems.push(`"${name}" cannot be kept; the base slide offers ${[...keepable.keys()].map((slot) => `"${slot}"`).join(', ') || 'nothing'}.`);
  const kept = design.keep.filter((name) => keepable.has(name)).map((name) => (keepable.get(name) as DesignKit['keepable'][number]).cells);
  // Names the new pattern already has: the kept slots and the deck's own texts.
  const names = new Set<string>([...design.keep, ...parts.base.slots.filter((slot) => parts.stays.has(slot.shapeId)).map((slot) => slot.name)]);

  // Every pair of parts is compared below, and a part takes at least one cell of the grid: a
  // design of more parts than the grid has cells is refused before that, in one sentence.
  if (design.elements.length > GRID * GRID) {
    return { problems: [`The layout has ${design.elements.length} parts; a layout takes at most ${GRID * GRID}.`], soft: [], placed: [] };
  }
  const placed: Placed[] = [];
  design.elements.forEach((element, index) => {
    const label = `Part ${index + 1}${element.name ? ` ("${element.name}")` : ''}`;
    const style = styles.get(element.style);
    if (!style) return void problems.push(`${label} uses style "${element.style}", which the template does not have.`);
    if ((style.kind === 'decoration') !== (element.role === 'decoration')) {
      return void problems.push(`${label} is a ${element.role} with ${style.key}, which is ${style.kind === 'decoration' ? 'a decoration' : 'a text style'}; text roles take T styles and decorations take D styles.`);
    }
    const cells: Cells = { column: element.column, row: element.row, width: element.width, height: element.height };
    if (![cells.column, cells.row, cells.width, cells.height].every((value) => Number.isInteger(value) && value >= 1) || cells.column + cells.width - 1 > GRID || cells.row + cells.height - 1 > GRID) {
      return void problems.push(`${label} is at ${describeCells(cells)}, outside the grid of ${GRID} columns and ${GRID} rows.`);
    }
    if (element.role !== 'decoration') {
      if (!NAME.test(element.name)) problems.push(`${label}: a text needs a name of letters and digits that starts with a lower case letter, such as "point1Text".`);
      else if (names.has(element.name)) problems.push(`${label}: the name "${element.name}" is taken; every text needs its own.`);
      names.add(element.name);
      if (!element.sample.trim()) problems.push(`${label} has no sample text.`);
      const lines = Math.max(1, element.sample.split('\n').length);
      if (style.lineRows !== undefined && cells.height < Math.max(1, style.lineRows * 1.1)) {
        problems.push(`${label} is ${cells.height} rows high, but one line of ${style.key} takes ${style.lineRows} rows; give it at least ${Math.ceil(style.lineRows * 1.1)}.`);
      } else if (style.lineRows !== undefined && cells.height < style.lineRows * lines) {
        soft.push(`${label}'s sample has ${lines} paragraphs, which take about ${Math.ceil(style.lineRows * lines)} rows; it has ${cells.height}.`);
      }
    }
    for (let row = cells.row; row < cells.row + cells.height; row++) {
      for (let column = cells.column; column < cells.column + cells.width; column++) {
        if (kit.taken[row - 1]?.[column - 1]) {
          problems.push(`${label} covers row ${row}, column ${column}, which the slide's header or footer takes.`);
          row = GRID + 1;
          break;
        }
      }
    }
    for (const [slot, entry] of keepable) {
      if (design.keep.includes(slot) && cellsOverlap(cells, entry.cells)) problems.push(`${label} at ${describeCells(cells)} covers the kept ${slot} at ${describeCells(entry.cells)}.`);
    }
    placed.push({ element, style, cells, box: { x: 0, y: 0, w: 0, h: 0 } });
  });

  // Texts may not lie on each other; a rule may not cross a text; a panel holds the texts on it whole.
  for (let i = 0; i < placed.length; i++) {
    for (let j = i + 1; j < placed.length; j++) {
      const a = placed[i] as Placed;
      const b = placed[j] as Placed;
      if (!cellsOverlap(a.cells, b.cells)) continue;
      const one = `"${a.element.name || a.style.key}" (${describeCells(a.cells)})`;
      const other = `"${b.element.name || b.style.key}" (${describeCells(b.cells)})`;
      const aDeco = a.element.role === 'decoration';
      const bDeco = b.element.role === 'decoration';
      if (!aDeco && !bDeco) problems.push(`The texts ${one} and ${other} lie on each other.`);
      else if (aDeco !== bDeco) {
        const [deco, text] = aDeco ? [a, b] : [b, a];
        if (deco.style.shape !== 'area') problems.push(`The rule ${aDeco ? one : other} crosses the text ${aDeco ? other : one}.`);
        else if (!inside(text.cells, deco.cells)) problems.push(`The text ${text === a ? one : other} reaches over the edge of the panel ${deco === a ? one : other}; put it wholly on the panel or off it.`);
        else if (placed.indexOf(deco) > placed.indexOf(text)) problems.push(`The panel ${deco === a ? one : other} is listed after the text on it and would hide it; list it first.`);
      }
    }
  }
  if (!placed.some((part) => part.element.role !== 'decoration') && kept.length === 0) problems.push('The layout has no text for a deck to fill.');

  // The slide should look full: the parts together reach across the content area and down most of it.
  const free = Array.from({ length: GRID }, (_, row) => row).filter((row) => !(kit.taken[row] as boolean[]).every(Boolean));
  // The kept title reaches across the slide by itself; what counts is where the new parts are.
  const all = placed.map((part) => part.cells);
  const keptRows = new Set(kept.flatMap((cells) => Array.from({ length: cells.height }, (_, i) => cells.row + i)));
  const columns = new Set<number>();
  const rows = new Set<number>();
  for (const cells of all) {
    for (let c = cells.column; c < cells.column + cells.width; c++) columns.add(c);
    for (let r = cells.row; r < cells.row + cells.height; r++) rows.add(r);
  }
  const emptyColumns = Array.from({ length: GRID }, (_, c) => c + 1).filter((c) => !columns.has(c));
  // A column left free between items is a gutter; a free column at the side is a hole.
  if (emptyColumns.length > 2 || emptyColumns.includes(1) || emptyColumns.includes(GRID)) soft.push(`Nothing is in column${emptyColumns.length === 1 ? '' : 's'} ${ranges(emptyColumns)}: that part of the slide stays empty. Spread the parts over the whole width.`);
  const emptyRows = free.map((row) => row + 1).filter((row) => !rows.has(row) && !keptRows.has(row));
  if (emptyRows.length > Math.max(2, free.length * 0.25)) soft.push(`Nothing is in row${emptyRows.length === 1 ? '' : 's'} ${ranges(emptyRows)}: that part of the slide stays empty. Use more of the height.`);

  for (const part of placed) {
    const panel = part.element.role === 'decoration' ? undefined : placed.find((other) => other.element.role === 'decoration' && other.style.shape === 'area' && inside(part.cells, other.cells))?.cells;
    part.box = boxOf(part.cells, kit.area, parts.info.slideWidth / parts.info.slideHeight, panel);
  }
  if (problems.length === 0) {
    evenItems(placed);
    // Where everything is drawn, a rule may not run through a text or cut through a panel, as one
    // between cards does when the cards are spaced evenly and it was not in a gap. It is left out.
    for (const rule of placed.filter((part) => part.element.role === 'decoration' && part.style.shape !== 'area')) {
      const hit = placed.find((other) => other !== rule && (other.element.role === 'decoration' ? other.style.shape === 'area' && cuts(rule, other.box) : runsAcross(rule, other.box)));
      if (!hit) continue;
      placed.splice(placed.indexOf(rule), 1);
      const what = hit.element.role === 'decoration' ? `the panel ${hit.style.key} (${describeCells(hit.cells)})` : `the text "${hit.element.name}" (${describeCells(hit.cells)})`;
      soft.push(`The rule ${rule.style.key} (${describeCells(rule.cells)}) runs through ${what} where they are drawn, so it is left out. Put a rule in the gap between two items, or beside them, or leave it out.`);
    }
  }
  return { problems, soft, placed };
}

/** The line a rule is drawn as: along the middle of its box, as {@link placeCopy} draws it. */
function ruleLine(rule: Placed): Box {
  const { x, y, w, h } = rule.box;
  return rule.style.shape === 'down' ? { x: x + w / 2, y, w: 0, h } : { x, y: y + h / 2, w, h: 0 };
}

/** Whether a rule runs through the inside of a box, and not only along its edge. */
function runsAcross(rule: Placed, box: Box): boolean {
  const line = ruleLine(rule);
  const tol = 0.002;
  if (line.w === 0) return line.x > box.x + tol && line.x < right(box) - tol && Math.min(bottom(line), bottom(box)) - Math.max(line.y, box.y) > tol;
  return line.y > box.y + tol && line.y < bottom(box) - tol && Math.min(right(line), right(box)) - Math.max(line.x, box.x) > tol;
}

/** Whether a rule cuts through a panel: it runs through the inside and out past its edge. A rule within a card is part of the card. */
function cuts(rule: Placed, panel: Box): boolean {
  if (!runsAcross(rule, panel)) return false;
  const line = ruleLine(rule);
  const tol = 0.005;
  return line.w === 0 ? line.y < panel.y - tol || bottom(line) > bottom(panel) + tol : line.x < panel.x - tol || right(line) > right(panel) + tol;
}

const union = (boxes: Box[]): Box => {
  const x = Math.min(...boxes.map((box) => box.x));
  const y = Math.min(...boxes.map((box) => box.y));
  return { x, y, w: Math.max(...boxes.map(right)) - x, h: Math.max(...boxes.map(bottom)) - y };
};

/**
 * Makes the numbered items of a design (point1..., point2...) equally large and equally far
 * apart, over the room the model drew them in. Twelve columns do not divide into five equal
 * items, and the software lays a row out at an even step from its first item, so items of
 * different widths would run into each other in a deck. A decoration goes with the item it holds
 * or lies within; a rule between two items stays in the middle of the gap.
 */
function evenItems(placed: Placed[]): void {
  const numberOf = (name: string) => /^[a-z][A-Za-z]*?(\d+)(?=[A-Z]|$)/.exec(name)?.[1];
  const items = new Map<string, Placed[]>();
  for (const part of placed) {
    const n = part.element.role === 'decoration' ? undefined : numberOf(part.element.name);
    if (n) items.set(n, [...(items.get(n) ?? []), part]);
  }
  if (items.size < 2) return;
  const groups = [...items.values()];
  const texts = groups.map((group) => union(group.map((part) => part.box)));
  const loose: Placed[] = [];
  const tol = 0.01;
  const holds = (outer: Box, inner: Box) => outer.x <= inner.x + tol && right(outer) >= right(inner) - tol && outer.y <= inner.y + tol && bottom(outer) >= bottom(inner) - tol;
  // First the panels that hold an item's texts; then what lies within an item with its panel, such
  // as a rule across a card, which reaches nearer the card's edge than the texts on it do.
  const rest: Placed[] = [];
  for (const part of placed) {
    if (part.element.role !== 'decoration') continue;
    const owner = texts.findIndex((box) => holds(part.box, box));
    if (owner >= 0) (groups[owner] as Placed[]).push(part);
    else rest.push(part);
  }
  const owned = groups.map((group) => union(group.map((part) => part.box)));
  for (const part of rest) {
    const owner = owned.findIndex((box) => holds(box, part.box));
    if (owner >= 0) (groups[owner] as Placed[]).push(part);
    else loose.push(part);
  }
  const boxes = groups.map((group) => union(group.map((part) => part.box)));
  const sharesDown = (a: Box, b: Box) => Math.min(bottom(a), bottom(b)) - Math.max(a.y, b.y) > 0;
  const sharesAcross = (a: Box, b: Box) => Math.min(right(a), right(b)) - Math.max(a.x, b.x) > 0;
  const across = boxes.every((a) => boxes.every((b) => a === b || sharesDown(a, b)));
  const down = !across && boxes.every((a) => boxes.every((b) => a === b || sharesAcross(a, b)));
  if (!across && !down) return;

  const start = (box: Box) => (across ? box.x : box.y);
  const size = (box: Box) => (across ? box.w : box.h);
  const order = groups.map((group, index) => ({ group, box: boxes[index] as Box })).sort((a, b) => start(a.box) - start(b.box));
  const gaps = order.slice(1).map((entry, index) => start(entry.box) - (start((order[index] as { box: Box }).box) + size((order[index] as { box: Box }).box)));
  if (gaps.some((gap) => gap < -tol)) return;
  const sizes = order.map((entry) => size(entry.box));
  if (Math.max(...sizes) - Math.min(...sizes) < 0.004 && Math.max(...gaps) - Math.min(...gaps) < 0.004) return;
  const from = start((order[0] as { box: Box }).box);
  const to = Math.max(...order.map((entry) => start(entry.box) + size(entry.box)));
  const gap = Math.max(GAP, median(gaps));
  const each = (to - from - gap * (order.length - 1)) / order.length;
  if (!(each > 0)) return;

  const moves = order.map((entry, index) => ({ was: entry.box, at: from + index * (each + gap) }));
  // A loose rule between two items goes to the middle of their new gap.
  for (const part of loose) {
    const middle = across ? part.box.x + part.box.w / 2 : part.box.y + part.box.h / 2;
    const index = moves.findIndex((move, i) => i < moves.length - 1 && middle >= start(move.was) + size(move.was) - tol && middle <= start((moves[i + 1] as { was: Box }).was) + tol);
    if (index < 0) continue;
    const centre = (moves[index] as { at: number }).at + each + gap / 2;
    if (across) part.box = { ...part.box, x: round(centre - part.box.w / 2) };
    else part.box = { ...part.box, y: round(centre - part.box.h / 2) };
  }
  order.forEach((entry, index) => {
    const move = moves[index] as { was: Box; at: number };
    const scale = each / size(move.was);
    for (const part of entry.group) {
      if (across) part.box = { ...part.box, x: round(move.at + (part.box.x - move.was.x) * scale), w: round(part.box.w * scale) };
      else part.box = { ...part.box, y: round(move.at + (part.box.y - move.was.y) * scale), h: round(part.box.h * scale) };
    }
  });
}

const ranges = (numbers: number[]): string => {
  const out: string[] = [];
  for (let i = 0; i < numbers.length; ) {
    let j = i;
    while (j + 1 < numbers.length && (numbers[j + 1] as number) === (numbers[j] as number) + 1) j++;
    out.push(i === j ? String(numbers[i]) : `${numbers[i]}-${numbers[j]}`);
    i = j + 1;
  }
  return out.join(', ');
};

// ---------------------------------------------------------------------------
// Building

/**
 * Builds a designed layout: a new slide at the end of the template's pptx, made from the base
 * slide's frame and copies of the template's own shapes, and a pattern for it. No model. Throws
 * {@link DesignError} with every reason when the design cannot be built.
 */
export function buildPattern(template: Template, design: PatternDesign): ComposedPattern {
  const info = inspectTemplate(template.pkg);
  const parts = kitWith(template, info, design.base);
  if (!parts) throw new DesignError(['The template has no layout with a title to stand on.']);
  const checked = review(design, parts, template.manifest);
  if (checked.problems.length > 0) throw new DesignError(checked.problems);
  return assemble(template, design, parts, checked.placed, checked.soft);
}

// Remembered for the manifest object the layouts were built under: a template that was given
// another manifest builds them again. A manifest object edited in place is not seen.
const built = new WeakMap<Template, { manifest: TemplateManifest; known: Map<string, Template> }>();

/**
 * The template with layouts made for one deck built onto it, in order: what a deck that carries
 * `layouts` is rendered, checked and edited with. The template itself is not changed. A layout
 * whose id the template already has is not built again: the template's own is used, as after the
 * layout was added to the template in a studio. Built once for a template and a list of layouts.
 * Throws {@link DesignError} when a layout no longer fits the template, or when `layouts` is no
 * list of layouts.
 */
export function withLayouts(template: Template, layouts: readonly PatternDesign[] | undefined): Template {
  // Layouts read from a stored deck may be anything: what is no list of layouts is said, not tripped over.
  const given: unknown = layouts;
  if (given !== undefined && given !== null && !(Array.isArray(given) && given.every((design) => typeof design === 'object' && design !== null && !Array.isArray(design)))) {
    throw new DesignError(["The deck's layouts are no list of layouts, such as [{ \"id\": \"three-goals\", \"elements\": [...] }]"]);
  }
  const wanted = (layouts ?? []).filter((design) => !template.manifest.patterns.some((pattern) => pattern.id === design.id));
  if (wanted.length === 0) return template;
  const key = JSON.stringify(wanted);
  let cache = built.get(template);
  if (!cache || cache.manifest !== template.manifest) built.set(template, (cache = { manifest: template.manifest, known: new Map() }));
  const known = cache.known.get(key);
  if (known) return known;
  let result = template;
  for (const design of wanted) {
    try {
      const made = buildPattern(result, design);
      result = loadTemplate(made.pptx, made.manifest, template.fonts ? { fonts: template.fonts } : {});
    } catch (error) {
      const reasons = error instanceof DesignError ? error.problems : [error instanceof Error ? error.message : String(error)];
      throw new DesignError(reasons.map((reason) => `The deck's layout "${design.id}": ${reason}`));
    }
  }
  cache.known.set(key, result);
  return result;
}

function assemble(template: Template, design: PatternDesign, parts: KitParts, placed: Placed[], soft: string[]): ComposedPattern {
  const { base, slide, stays, info } = parts;
  // A sample is written in the template's own language and direction.
  const { language, direction } = template.manifest;
  const writing: TextOptions = { flipAlign: false, direction, templateDirection: direction, ...(language ? { language } : {}) };
  const pkg = template.pkg.clone();
  const pres = readPresentation(pkg);
  const baseRef = pres.slides[slide.number - 1];
  if (!baseRef) throw new DesignError([`The template has no slide ${slide.number}.`]);
  const doc = parseXml(pkg.readText(baseRef.part));
  const root = rootElement(doc);
  const tree = path(root, [NS.p, 'cSld'], [NS.p, 'spTree']) as Element;

  // The base slide keeps its frame, the deck's own texts and the slots the design keeps; nothing else.
  const keepIds = new Set(stays);
  for (const name of design.keep) {
    const slot = base.slots.find((candidate) => candidate.name === name);
    if (slot) keepIds.add(slot.shapeId);
  }
  const ids = (element: Element) => descendants(element, NS.p, 'cNvPr').map((props) => props.getAttribute('id') ?? '');
  for (const child of childElements(tree, NS.p)) {
    if (child.localName === 'nvGrpSpPr' || child.localName === 'grpSpPr') continue;
    if (!ids(child).some((id) => keepIds.has(id))) removeNode(child);
  }
  // Animations name shapes that are gone.
  for (const timing of childElements(root, NS.p, 'timing')) removeNode(timing);
  path(root, [NS.p, 'cSld'])?.setAttribute('name', `${COMPOSED}${design.id}`);

  let nextId = Math.max(1, ...descendants(root, NS.p, 'cNvPr').map((props) => Number(props.getAttribute('id') ?? 0)).filter(Number.isFinite)) + 1;
  const sources = new Map<number, Map<string, Element>>();
  const sourceShape = (number: number, id: string): Element => {
    if (!sources.has(number)) sources.set(number, indexShapes(rootElement(parseXml(pkg.readText((pres.slides[number - 1] as { part: string }).part)))));
    return (sources.get(number) as Map<string, Element>).get(id) as Element;
  };
  const made: Array<{ part: Placed; id: string }> = [];
  for (const part of placed) {
    const copy = doc.importNode(sourceShape(part.style.slide, part.style.shapeId), true) as Element;
    const props = descendants(copy, NS.p, 'cNvPr')[0] as Element;
    const id = String(nextId++);
    props.setAttribute('id', id);
    props.setAttribute('name', part.element.name || `${part.style.key} decoration`);
    props.removeAttribute('descr');
    for (const ext of childElements(props, NS.a, 'extLst')) removeNode(ext);
    placeCopy(copy, part, info);
    if (part.element.role !== 'decoration') setShapeText(copy, part.element.sample.split('\n').map((line) => line.trim()).filter(Boolean), writing);
    tree.appendChild(copy);
    made.push({ part, id });
  }

  // The new slide goes after the last one.
  const types = ContentTypes.read(pkg);
  const dir = baseRef.part.slice(0, baseRef.part.lastIndexOf('/'));
  let index = pres.slides.length + 1;
  while (pkg.has(`${dir}/slide${index}.xml`)) index++;
  const part = `${dir}/slide${index}.xml`;
  const usedRels = new Set<string>();
  for (const element of [root, ...Array.from(root.getElementsByTagNameNS('*', '*'))]) {
    for (let i = 0; i < element.attributes.length; i++) {
      const attr = element.attributes.item(i);
      if (attr?.namespaceURI === NS.r) usedRels.add(attr.value);
    }
  }
  const rels = readRels(pkg, baseRef.part).filter((rel) => !isRelType(rel, 'notesSlide') && (isRelType(rel, 'slideLayout') || usedRels.has(rel.id)));
  pkg.write(part, serializeXml(doc));
  writeRels(pkg, part, rels);
  types.setOverride(part, 'application/vnd.openxmlformats-officedocument.presentationml.slide+xml');
  types.save(pkg);
  appendToPresentation(pkg, pres.part, part, baseRef.relId);

  // The pattern: the base slots of every shape that stayed, then a slot for each text.
  const number = pres.slides.length + 1;
  const manifest = structuredClone(template.manifest);
  const kept = new Set(descendants(root, NS.p, 'cNvPr').map((props) => props.getAttribute('id') ?? ''));
  const pattern: PatternDef = {
    id: design.id,
    sourceSlide: number,
    description: design.description.trim(),
    slots: base.slots.filter((slot) => kept.has(slot.shapeId) && (keepIds.has(slot.shapeId) || slot.kind === 'text' && (slot.deckField !== undefined || slot.empty === 'drop'))).map((slot) => structuredClone(slot)),
  };
  manifest.patterns.push(pattern);
  const built = loadTemplate(pkg.save(), manifest, template.fonts ? { fonts: template.fonts } : {});
  const changes: ManifestChange[] = made
    .filter(({ part: placedPart }) => placedPart.element.role !== 'decoration')
    .map(({ part: placedPart, id }) => ({ op: 'addSlot', pattern: design.id, shape: id, role: placedPart.element.role as Exclude<DesignRole, 'decoration'>, name: placedPart.element.name, optional: placedPart.element.optional }));
  if (placed.some((placedPart) => placedPart.element.role === 'number')) changes.push({ op: 'setPattern', pattern: design.id, figures: true });
  const builtInfo = inspectTemplate(built.pkg);
  let result = changeWith(built, changes, builtInfo);
  // A budget is what the box holds, not what the model's sample happened to be: measured with the
  // template's fonts when it has them, estimated from the box when it does not.
  const shapesNow = new Map((builtInfo.slides[number - 1]?.shapes ?? []).map((shape) => [shape.id, shape]));
  const measured = new Map<string, number>();
  if (template.fonts) {
    for (const budget of measuredBudgets({ ...built, manifest: result.manifest })) if (budget.pattern === design.id) measured.set(budget.slot, budget.suggested);
  }
  const budgets: ManifestChange[] = [];
  for (const { part: placedPart, id } of made) {
    const role = placedPart.element.role;
    const shape = shapesNow.get(id);
    if (role === 'decoration' || !shape || !result.manifest.patterns.find((candidate) => candidate.id === design.id)?.slots.some((slot) => slot.name === placedPart.element.name)) continue;
    const room = measured.get(placedPart.element.name) ?? estimate(shape, builtInfo, placedPart.style.size ?? shape.fontSizes?.[0] ?? 14);
    const sample = placedPart.element.sample.replace(/\s+/g, ' ').trim().length;
    const budget: ManifestChange = { op: 'setSlot', pattern: design.id, slot: placedPart.element.name, maxChars: Math.max(1, role === 'number' ? Math.max(room, sample, 2) : room) };
    // A box drawn for more than its sample is filled to the sample's length, as the design meant;
    // a deck that writes a third of that leaves a card mostly empty.
    if ((role === 'text' || role === 'list') && sample >= 20 && room >= sample * 1.5) budget.description = `Drawn for about ${Math.round(sample / 10) * 10} characters`;
    budgets.push(budget);
  }
  if (budgets.length > 0) {
    const set = changeWith({ ...built, manifest: result.manifest }, budgets, builtInfo);
    result = { ...set, applied: [...result.applied, ...set.applied], rejected: [...result.rejected, ...set.rejected] };
  }
  const warnings = [...soft, ...result.rejected.map((entry) => `A change to the new layout was refused: ${entry.reason}`)];
  // Numbered names make a row: the planner may then ask for more or fewer of them.
  const numbered = new Set(placed.map((placedPart) => /^[a-z]+?(\d+)[A-Z]?/.exec(placedPart.element.name)?.[1]).filter(Boolean));
  let row = '';
  if (numbered.size >= 2) {
    const repeat = changeWith({ ...built, manifest: result.manifest }, [{ op: 'setRepeat', pattern: design.id }], builtInfo);
    if (repeat.applied.length > 0) {
      result = repeat;
      row = ` ${repeat.applied[0]?.message ?? ''}`;
      // With fewer items than drawn, the items of a row grow wider, as prepareManifest lets rows do.
      // A grid keeps its cards as drawn: what is inside a card is not laid out for a wider one.
      const grid = result.manifest.patterns.find((candidate) => candidate.id === design.id)?.repeat?.columns !== undefined;
      const wider = grid ? undefined : changeWith({ ...built, manifest: result.manifest }, [{ op: 'setRepeat', pattern: design.id, widen: 2 }], builtInfo);
      if (wider && wider.applied.length > 0) result = wider;
    } else warnings.push(`The numbered parts were not made a row: ${repeat.rejected[0]?.reason ?? 'no reason given'}.`);
  }
  const final = { ...built, manifest: result.manifest };
  const errors = checkManifest(final, builtInfo).filter((issue) => issue.severity === 'error' && issue.pattern === design.id);
  if (errors.length > 0) throw new DesignError(errors.map((issue) => `${issue.slot ?? issue.pattern}: ${issue.message}`));
  try {
    const only = { ...final, manifest: { ...final.manifest, patterns: final.manifest.patterns.filter((candidate) => candidate.id === design.id) } };
    for (const items of ['min', 'max'] as const) {
      const { deck, assets } = sampleDeck(only, { items });
      renderDeck(only, deck, { assets, layout: 'adaptive' });
    }
  } catch (error) {
    throw new DesignError([`The sample of the new layout could not be rendered: ${error instanceof Error ? error.message : String(error)}`]);
  }

  const texts = placed.filter((placedPart) => placedPart.element.role !== 'decoration').length;
  const ornaments = placed.length - texts;
  return {
    pptx: final.pkg.save(),
    manifest: final.manifest,
    pattern: design.id,
    slide: number,
    summary: `"${design.id}" is a new layout on slide ${number}, on the frame of "${base.id}"${design.keep.length ? ` with its ${design.keep.join(' and ')}` : ''}: ${texts} text${texts === 1 ? '' : 's'}${ornaments ? ` and ${ornaments} decoration${ornaments === 1 ? '' : 's'}` : ''} in the template's own styles.${row}`,
    design,
    warnings,
  };
}

/**
 * How many characters a composed text box holds, without fonts to measure with: whole lines
 * only, inside the box's usual margins, with letters a little wider than the estimate for a
 * template's own boxes. The sample text is the model's, so it does not count.
 */
function estimate(shape: ShapeInfo, info: TemplateInfo, size: number): number {
  const width = (shape.box.w * info.slideWidth) / 12700 - 14.4;
  const height = (shape.box.h * info.slideHeight) / 12700 - 7.2;
  const perLine = Math.max(1, Math.floor(width / (size * 0.55)));
  const lines = Math.max(1, Math.floor(height / (size * 1.25)));
  return Math.max(1, Math.floor(perLine * lines * 0.85));
}

/** Sets a copied shape's place. A rule keeps its own thickness and runs along the middle of its cells. */
function placeCopy(copy: Element, part: Placed, info: TemplateInfo): void {
  const spPr = childElement(copy, NS.p, 'spPr');
  const xfrm = spPr && childElement(spPr, NS.a, 'xfrm');
  const off = xfrm && childElement(xfrm, NS.a, 'off');
  const ext = xfrm && childElement(xfrm, NS.a, 'ext');
  if (!xfrm || !off || !ext) throw new DesignError([`Style ${part.style.key} has no position of its own to copy.`]);
  xfrm.removeAttribute('rot');
  if (part.element.role !== 'decoration') {
    xfrm.removeAttribute('flipH');
    xfrm.removeAttribute('flipV');
  }
  let { x, y, w, h } = part.box;
  // A rounded panel keeps the corners it is drawn with, whatever its new size.
  const tall = info.slideHeight / info.slideWidth;
  const was = Math.min(part.style.drawn.w, part.style.drawn.h * tall);
  const now = Math.min(w, h * tall);
  if (part.style.shape === 'area' && was > 0 && now > 0) keepCorners(copy, was, now, Number.POSITIVE_INFINITY);
  if (part.style.shape === 'across') {
    const thick = part.style.drawn.h;
    y = y + h / 2 - thick / 2;
    h = thick;
  } else if (part.style.shape === 'down') {
    const thick = part.style.drawn.w;
    x = x + w / 2 - thick / 2;
    w = thick;
  }
  off.setAttribute('x', String(Math.round(x * info.slideWidth)));
  off.setAttribute('y', String(Math.round(y * info.slideHeight)));
  ext.setAttribute('cx', String(Math.max(0, Math.round(w * info.slideWidth))));
  ext.setAttribute('cy', String(Math.max(0, Math.round(h * info.slideHeight))));
}

/** Adds a slide part at the end of the presentation's slide list, and to its last section when it has sections. */
function appendToPresentation(pkg: PptxPackage, presPart: string, slidePart: string, baseRelId: string): void {
  const rels = readRels(pkg, presPart);
  const type = rels.find((rel) => rel.id === baseRelId)?.type ?? rels.find((rel) => isRelType(rel, 'slide'))?.type ?? '';
  const relId = nextRelId(rels);
  const fresh: Relationship = { id: relId, type, target: relativeTarget(presPart, slidePart), external: false };
  writeRels(pkg, presPart, [...rels, fresh]);

  const doc = parseXml(pkg.readText(presPart));
  const root = rootElement(doc);
  const list = childElement(root, NS.p, 'sldIdLst') as Element;
  const entries = childElements(list, NS.p, 'sldId');
  const id = String(Math.max(255, ...entries.map((entry) => Number(entry.getAttribute('id') ?? 0))) + 1);
  const entry = doc.createElementNS(NS.p, `${root.prefix ? `${root.prefix}:` : ''}sldId`);
  entry.setAttribute('id', id);
  entry.setAttributeNS(NS.r, `${root.lookupPrefix(NS.r) ?? 'r'}:id`, relId);
  list.appendChild(entry);
  const P14 = 'http://schemas.microsoft.com/office/powerpoint/2010/main';
  const sections = descendants(root, P14, 'section');
  const last = sections.at(-1);
  const lastList = last && childElement(last, P14, 'sldIdLst');
  if (lastList) {
    const item = (doc as Document).createElementNS(P14, `${lastList.prefix ? `${lastList.prefix}:` : ''}sldId`);
    item.setAttribute('id', id);
    lastList.appendChild(item);
  }
  pkg.write(presPart, serializeXml(doc));
}

// ---------------------------------------------------------------------------
// Asking the model

const SYSTEM = [
  'You design one new slide layout for an existing PowerPoint template, so that it looks like it was always part of the template.',
  'You place copies of the template\'s own text styles and decorations on a grid over the slide\'s content area. You never invent a colour, a font or a size: you choose among the styles listed.',
  'Answer only with JSON that follows the schema.',
].join(' ');

const GUIDE = `How to design it:
- Keep the template's title where every title is: put "title" in "keep" unless the person asks for a slide without the usual title. Keep "subtitle" too when the template's content slides have one.
- Fill the content area. The parts together should reach from column 1 to column ${GRID} and use most of the free rows. A layout that leaves half the slide empty is wrong. Make text boxes large rather than small: a deck writes into them.
- Use the template's styles for what they are used for: heading styles for headings, body styles for text, a style with bullets for a list.
- The role "number" is only for a figure that measures something (85%, 2.4M, 12 days), and only when the person asks for figures. The number of a step or an item (01, 02, 03) is a "caption", even in a big number style: a deck without figures cannot use a layout whose role is "number".
- Use the template's decorations as it does: a panel behind a heading and its text makes a card, a rule separates. Give every panel the same size in a set of cards. List a panel before the texts that sit on it.
- Repeated items (columns, cards, steps) get numbered names: point1Heading, point1Text, point2Heading, point2Text. Make the items the same size and spread them evenly.
- Align parts to each other: items in a row start on the same row and are equally high; a column of parts shares its edges.
- Leave one empty column or row between items, not more.
- For a right-to-left template, reading starts at column ${GRID}: the first item goes on the right.
- A text needs rows for its lines: a style's line takes the rows the list says. Give each text room for about a third more than its sample.
- Write the sample texts in the template's language, as a real slide would read, at the length a real slide would have.
- Do not repeat a layout the template already has.
- "id" is short and in lower case with hyphens. "description" tells the planner what content the layout suits, in one or two sentences.`;

function gridMap(kit: DesignKit): string {
  const letters = new Map<string, string>();
  const keep: string[][] = Array.from({ length: GRID }, () => Array.from({ length: GRID }, () => ''));
  kit.keepable.forEach((entry, index) => {
    const letter = entry.role === 'title' ? 't' : String.fromCharCode(115 + index); // s, ...
    letters.set(entry.slot, letter);
    for (let r = entry.cells.row; r < entry.cells.row + entry.cells.height; r++) for (let c = entry.cells.column; c < entry.cells.column + entry.cells.width; c++) (keep[r - 1] as string[])[c - 1] = letter;
  });
  const lines = [`      ${Array.from({ length: GRID }, (_, c) => String(c + 1).padStart(3)).join('')}`];
  for (let r = 0; r < GRID; r++) {
    lines.push(`row ${String(r + 1).padStart(2)}${Array.from({ length: GRID }, (_, c) => (kit.taken[r]?.[c] ? '  #' : (keep[r]?.[c] ?? '') !== '' ? `  ${keep[r]?.[c]}` : '  .')).join('')}`);
  }
  const legend = ['# is the frame (header, logo, page number): nothing goes there.', ...kit.keepable.map((entry) => `${letters.get(entry.slot)} is the ${entry.slot} of the base slide, at ${describeCells(entry.cells)}: free when you do not keep it.`)];
  return `${lines.join('\n')}\n${legend.join('\n')}`;
}

function designPrompt(template: Template, parts: KitParts, instruction: string): string {
  const { kit, info } = parts;
  const { manifest } = template;
  const inches = (fraction: number, of: number) => Math.round(((fraction * of) / EMU_PER_INCH) * 100) / 100;
  const texts = kit.styles.filter((style) => style.kind === 'text');
  const ornaments = kit.styles.filter((style) => style.kind === 'decoration');
  return [
    `The template "${manifest.name}" is ${inches(1, info.slideWidth)} x ${inches(1, info.slideHeight)} inches, written in ${manifest.language ?? 'its own language'}, reading ${manifest.direction === 'rtl' ? 'right to left' : 'left to right'}.`,
    `The new slide stands on the frame of the layout "${kit.base}" (slide ${kit.baseSlide}): its background, header, logo and page number stay. The content area is a grid of ${GRID} columns and ${GRID} rows, each cell ${inches(kit.area.w / GRID, info.slideWidth)} x ${inches(kit.area.h / GRID, info.slideHeight)} inches:\n${gridMap(kit)}`,
    `The layouts the template has:\n${manifest.patterns.map((pattern) => `- ${pattern.id}: ${pattern.description.slice(0, 160)}`).join('\n')}`,
    `Text styles (biggest first):\n${texts.map((style) => `${style.key}: ${style.text}`).join('\n')}`,
    ornaments.length > 0 ? `Decorations:\n${ornaments.map((style) => `${style.key}: ${style.text}`).join('\n')}` : 'The template has no decorations to use.',
    GUIDE,
    `The person asks:\n${instruction.trim()}`,
  ].join('\n\n');
}

function schemaFor(kit: DesignKit) {
  return {
    type: 'object',
    properties: {
      id: { type: 'string' },
      description: { type: 'string' },
      keep: { type: 'array', items: { type: 'string', enum: kit.keepable.length > 0 ? kit.keepable.map((entry) => entry.slot) : ['none'] } },
      elements: {
        type: 'array',
        items: {
          type: 'object',
          properties: {
            role: { type: 'string', enum: [...DESIGN_ROLES] },
            style: { type: 'string', enum: kit.styles.map((style) => style.key) },
            name: { type: 'string' },
            sample: { type: 'string' },
            optional: { type: 'boolean' },
            column: { type: 'integer' },
            row: { type: 'integer' },
            width: { type: 'integer' },
            height: { type: 'integer' },
          },
          required: ['role', 'style', 'name', 'sample', 'optional', 'column', 'row', 'width', 'height'],
        },
      },
    },
    required: ['id', 'description', 'keep', 'elements'],
  };
}

function toDesign(raw: unknown): PatternDesign {
  const answer = (raw && typeof raw === 'object' ? raw : {}) as Record<string, unknown>;
  const text = (value: unknown) => (typeof value === 'string' ? value : '');
  const int = (value: unknown) => (typeof value === 'number' ? Math.round(value) : Number.NaN);
  return {
    id: text(answer.id).trim().toLowerCase(),
    description: text(answer.description).trim(),
    keep: (Array.isArray(answer.keep) ? answer.keep : []).map(text).filter((name) => name && name !== 'none'),
    elements: (Array.isArray(answer.elements) ? answer.elements : []).map((item): DesignElement => {
      const element = (item && typeof item === 'object' ? item : {}) as Record<string, unknown>;
      const role = DESIGN_ROLES.includes(element.role as DesignRole) ? (element.role as DesignRole) : 'text';
      return {
        name: role === 'decoration' ? '' : text(element.name).trim(),
        style: text(element.style).trim(),
        role,
        optional: element.optional === true,
        sample: role === 'decoration' ? '' : text(element.sample),
        column: int(element.column),
        row: int(element.row),
        width: int(element.width),
        height: int(element.height),
      };
    }),
  };
}

/**
 * Asks a model for a new layout in the template's own look, and builds it with
 * {@link buildPattern}: a new slide at the end of the pptx and a pattern for it. The model
 * places copies of the template's text styles and decorations on a grid over the content area;
 * the code keeps the frame and the title where the template has them, works out every box, and
 * measures every budget. A design with problems is sent back once with the reasons. Throws
 * {@link DesignError} when the second answer cannot be built either.
 */
export async function composePattern(template: Template, instruction: string, options: ComposeOptions): Promise<ComposedPattern> {
  if (!instruction.trim()) throw new Error('Say what the new layout should be.');
  const info = inspectTemplate(template.pkg);
  const parts = kitWith(template, info, options.base);
  if (!parts) throw new DesignError(['The template has no layout with a title to stand on.']);
  const schema = schemaFor(parts.kit);
  const prompt = designPrompt(template, parts, instruction);
  const ask = async (text: string): Promise<PatternDesign> => ({
    ...toDesign(await options.llm.generateJson({ system: SYSTEM, prompt: text, schema, maxOutputTokens: 6000, ...(options.effort ? { effort: options.effort } : {}), ...(options.signal ? { signal: options.signal } : {}) })),
    base: parts.kit.base,
  });

  let design = await ask(prompt);
  let checked = review(design, parts, template.manifest);
  if (checked.problems.length > 0 || checked.soft.length > 0) {
    const second = await ask(
      [
        prompt,
        `You answered:\n${JSON.stringify(design)}`,
        `The software found these problems:\n${[...checked.problems, ...checked.soft].map((problem) => `- ${problem}`).join('\n')}`,
        'Answer with the whole layout again, with these problems solved.',
      ].join('\n\n'),
    );
    const again = review(second, parts, template.manifest);
    // The second answer is taken unless it is worse than the first.
    if (again.problems.length === 0 || again.problems.length <= checked.problems.length) {
      design = second;
      checked = again;
    }
  }
  if (checked.problems.length > 0) throw new DesignError(checked.problems);
  return assemble(template, design, parts, checked.placed, checked.soft);
}
