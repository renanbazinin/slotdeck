// What a free slide is made with, read from the template: the layout whose frame it stands on,
// the room its blocks take, and which of the template's own shapes each part is a copy of (a
// heading, running text, a big number, a caption, a card's panel). No model. The frame is the
// one a composed layout stands on; everything else is chosen here, by rules that look at how the
// template's own layouts use their shapes.
import { contrast, luminance } from '../color/space.ts';
import { chartColors, fillColor, isColor, literalOf, pageColor, paletteOf, schemeOf } from '../color/recolor.ts';
import type { Scheme } from '../color/recolor.ts';
import type { TemplatePalette } from '../color/palette.ts';
import type { PptxPackage } from '../pptx/package.ts';
import { readPresentation } from '../pptx/presentation.ts';
import { NS, childElement, childElements, descendants, parseXml, path, rootElement } from '../pptx/xml.ts';
import type { Element } from '../pptx/xml.ts';
import { indexShapes } from '../render/render.ts';
import type { Template } from '../render/render.ts';
import { frameParts } from '../template/compose.ts';
import { inspectTemplate } from '../template/inspect.ts';
import type { Box, ShapeInfo, TemplateInfo } from '../template/inspect.ts';
import { templateTextSetting } from '../template/manifest.ts';
import type { ChartSlotDef, Direction, PatternDef, SlideBox, SlotDef, TableSlotDef, TemplateManifest, TextSlotDef } from '../template/manifest.ts';
import { inFrameBand } from '../template/measure.ts';
import { TextStyles } from '../text/styles.ts';
import type { Spacing } from '../text/styles.ts';

/** A text style of the template: the shape a part of a free slide is a copy of, and what the code changes on the copy. */
export interface FreeTextStyle {
  slide: number;
  shapeId: string;
  /** Font size in points as the template draws it. */
  size: number;
  bold: boolean;
  italic: boolean;
  /** The source paragraph has a bullet. */
  bullet: boolean;
  /** The colour of its text, as six hex digits. */
  color: string;
  /** The fill of the text's own shape, for a text on a filled box. */
  fill?: string;
  /**
   * How far apart the template sets the lines of this text, where that can be more than the usual
   * 1.2 times the size: a share of the usual above 1 (1.3 for a line spacing of 130%), or a height
   * in points. A copy keeps the line spacing of its source, so its lines are counted at it. For
   * the engine's own counting: no part of the package's interface.
   */
  leading?: Spacing;
  /** What the copy is given that the source does not have, for a role the template has no shape for. */
  set?: { bold?: boolean; italic?: boolean; bullets?: boolean; size?: number; color?: string };
}

/** A shape of the template that is drawn and not filled with text: a panel, a rule. */
export interface FreeShapeStyle {
  slide: number;
  shapeId: string;
  /** Its fill, when it has one. */
  fill?: string;
  /** Its box as the template draws it. */
  drawn: SlideBox;
}

export const FREE_STYLE_ROLES = ['body', 'heading', 'figure', 'caption', 'list', 'quote', 'note'] as const;
export type FreeStyleRole = (typeof FREE_STYLE_ROLES)[number];

/**
 * What free slides of a template are made with. The engine's working data: the package does not
 * export it, and its fields change with the engine.
 */
export interface FreeKit {
  /** The layout whose slide gives the frame, and that slide's number. */
  base: string;
  baseSlide: number;
  /** Slide size in EMU. */
  slideWidth: number;
  slideHeight: number;
  /** The first id a shape made for a free slide gets: above every id of the base slide as the template has it. */
  firstId: number;
  /** Ids of the base slide's shapes that stay on a free slide: the frame and the deck's own texts. */
  stays: string[];
  /** Slots of the base that stay with it: the deck's own texts. */
  frameSlots: SlotDef[];
  /** The title: the base's title shape, the box it is shown in, its size in points and what it holds. */
  title?: { shapeId: string; slot: TextSlotDef; box: SlideBox; place: boolean; size: number; maxChars: number };
  /** The content area, as a composed layout has it. */
  area: SlideBox;
  /** Where the template's own layouts put their content, top and bottom, as fractions of the slide's height. */
  band: { top: number; bottom: number };
  /** The most room a crowded slide may take. */
  dense: { top: number; bottom: number };
  /** The space the template keeps between neighbouring items, as a fraction of the slide's width. */
  gap: number;
  /** The colour behind the base slide. */
  background: string;
  /** The colour of text on the base slide that states none: the theme's. */
  ink: string;
  /** The reading direction the template was drawn for. */
  direction: Direction;
  /** The factor the template puts on body text. */
  textScale: number;
  /** How wide PowerPoint sets a character of the template's running text, as a share of its size, where its manifest states it (`TemplateManifest.characterWidth`). */
  characterWidth?: number;
  /** The least size, in points, of a long text set dense on a free slide, where the manifest states it (`TemplateManifest.leastTextSize`, up to a half point); otherwise by the slide's width (`leastSizeOf` in plan.ts). */
  leastTextSize?: number;
  styles: Record<Exclude<FreeStyleRole, 'note'>, FreeTextStyle> & { note?: FreeTextStyle };
  /** The background of a card: a tinted or outlined panel that the template sets text on. */
  panel?: FreeShapeStyle;
  /** A strongly filled panel for a note, with the style of the text the template sets on it. */
  notePanel?: FreeShapeStyle & { text: FreeTextStyle };
  /** A thin line across. */
  rule?: FreeShapeStyle;
  /** Whether the template sets its big numbers on a panel. */
  figureOnPanel: boolean;
  /** A picture of the template whose outline (rounded corners) the photos of a free slide take. */
  picture?: { slide: number; shapeId: string };
  /** A table of the template to copy, with the slot that says how its text is set. */
  table?: { slide: number; shapeId: string; slot: TableSlotDef; rowHeight: number; size: number; margin: number };
  /** How a chart is drawn: the colours of its series, its text and its lines. */
  chart: Pick<ChartSlotDef, 'colors' | 'text' | 'lineColor'>;
  palette: TemplatePalette;
  /**
   * The band of colour the blocks of a slide stand on (`FreeArrange.tone`): its fill, the room
   * kept between its edge and the blocks as a fraction of the slide's width, and the template's
   * panel it is a copy of, for its corners, where the template has one. Only on the kit a toned
   * slide is planned with (`tonedKit` in plan.ts). `icons` says whether a card on it may carry its
   * icon: not on a band in the template's main colour, where an icon drawn for a light page does
   * not read.
   */
  ground?: { fill: string; pad: number; icons: boolean; from?: FreeShapeStyle };
}

const right = (box: Box) => box.x + box.w;
const bottom = (box: Box) => box.y + box.h;
const TOL = 0.005;
const within = (inner: Box, outer: Box) => inner.x >= outer.x - TOL && right(inner) <= right(outer) + TOL && inner.y >= outer.y - TOL && bottom(inner) <= bottom(outer) + TOL;
const median = (values: number[]): number | undefined => {
  const sorted = values.slice().sort((a, b) => a - b);
  return sorted.length === 0 ? undefined : sorted[Math.floor((sorted.length - 1) / 2)];
};
const round = (value: number) => Math.round(value * 10000) / 10000;
const clamp = (value: number, low: number, high: number) => Math.min(high, Math.max(low, value));
/**
 * A shape with no place of its own on the slide: a placeholder that takes its position from the
 * slide's layout. The template is read as showing it as nothing at the corner.
 */
const placeless = (box: Box) => !(box.w > 0 || box.h > 0);

const TITLE = /^(title|heading|slideTitle)$/i;
/** A slot that holds the line under a title. */
const SUBTITLE = /^sub(title|heading)\d*$/i;
const COMPOSED_SLIDE = /<(?:\w+:)?cSld\b[^>]*\bname="composed /;
const hasRelationship = (element: Element) =>
  [element, ...Array.from(element.getElementsByTagNameNS('*', '*'))].some((el) => {
    for (let i = 0; i < el.attributes.length; i++) if (el.attributes.item(i)?.namespaceURI === NS.r) return true;
    return false;
  });

interface Candidate extends FreeTextStyle {
  face: string;
  /** Slot names that use it, the most its slots hold, and what all of them hold together. */
  names: string[];
  budget: number;
  budgets: number;
  uses: number;
  /** The colour it is drawn on: a panel's fill, or the slide's background. */
  ground: string;
  /** It lies on a panel: a filled or outlined shape that is not the whole slide. */
  onPanel: boolean;
  /** The fill of the filled shape it lies on, when it lies on one. */
  panelFill?: string;
}

interface Area extends FreeShapeStyle {
  preset: string;
  outlined: boolean;
  holds: number;
  uses: number;
}

interface Scan {
  info: TemplateInfo;
  candidates: Candidate[];
  areas: Area[];
  rules: Area[];
  /** Per slide number: the colour behind it, and the colour of text that states none. */
  backgrounds: Map<number, string>;
  inks: Map<number, string>;
  /** How many texts of the layouts were passed over as a style to copy because they are placeholders. */
  placeholders: number;
  /** Where a shape of a slide stands: its own box, or, for a placeholder that states none, the one its layout or its master gives it. */
  stands: (slide: number, shape: ShapeInfo) => Box;
  /**
   * Top and bottom of the content of each pattern slide, and its title box when it has one. The
   * line under a title (a slot named as a subtitle) belongs to the title: `top` is where the
   * content under it starts, and `under` where the first thing under the title starts, that line
   * counted in.
   */
  content: Array<{ slide: number; top: number; under: number; bottom: number; title?: Box }>;
  gaps: number[];
  picture?: { slide: number; shapeId: string };
  table?: FreeKit['table'];
  chart?: FreeKit['chart'];
}

const scans = new WeakMap<PptxPackage, Map<string, Scan>>();

const runColor = (shape: Element, scheme: Scheme): string => {
  for (const props of [...descendants(shape, NS.a, 'rPr'), ...descendants(shape, NS.a, 'endParaRPr')]) {
    const fill = childElement(props, NS.a, 'solidFill');
    const color = fill && childElements(fill, NS.a).find(isColor);
    const hex = color && literalOf(color, scheme);
    if (hex) return hex;
  }
  return scheme.resolve('tx1') ?? '000000';
};

/** Reads every pattern slide once: its text styles as the slots use them, its panels and rules, where its content sits. */
function scan(template: Template): Scan {
  const { pkg, manifest } = template;
  // Everything read below comes from the package and from the manifest's patterns: slots with
  // their shapes, budgets, colours and text sizes. A scan is good for exactly these patterns.
  const key = JSON.stringify(manifest.patterns);
  let cache = scans.get(pkg);
  if (!cache) scans.set(pkg, (cache = new Map()));
  const known = cache.get(key);
  if (known) return known;

  const info = inspectTemplate(pkg);
  const pres = readPresentation(pkg);
  const styles = new TextStyles(pkg);
  // The usual title of a real template is a placeholder with no position of its own on the slide.
  const placeOf = (number: number, shape: ShapeInfo): Box => {
    if (!placeless(shape.box)) return shape.box;
    const from = styles.inheritedBox(number, shape.id);
    if (!from || !(from.cx > 0 && from.cy > 0)) return shape.box;
    return { x: round(from.x / info.slideWidth), y: round(from.y / info.slideHeight), w: round(from.cx / info.slideWidth), h: round(from.cy / info.slideHeight) };
  };
  const result: Scan = { info, candidates: [], areas: [], rules: [], backgrounds: new Map(), inks: new Map(), placeholders: 0, stands: placeOf, content: [], gaps: [] };
  const bySignature = new Map<string, Candidate>();
  const seen = new Set<string>();
  const areaBySignature = new Map<string, Area>();
  const ruleBySignature = new Map<string, Area>();
  const pictures = new Map<string, { slide: number; shapeId: string; uses: number }>();
  const numbers = [...new Set(manifest.patterns.map((pattern) => pattern.sourceSlide))].sort((a, b) => a - b);

  for (const number of numbers) {
    const ref = pres.slides[number - 1];
    const slide = info.slides[number - 1];
    if (!ref || !slide || !pkg.has(ref.part)) continue;
    const xml = pkg.readText(ref.part);
    // A layout made for a deck copies the template's shapes: reading it again would only repeat them.
    if (COMPOSED_SLIDE.test(xml)) continue;
    const root = rootElement(parseXml(xml));
    const elements = indexShapes(root);
    const scheme = schemeOf(pkg, ref.part);
    // The colour the page shows where content stands: of a page that is white but for a glow in a corner, white.
    const background = pageColor(pkg, ref.part, root).color ?? 'FFFFFF';
    result.backgrounds.set(number, background);
    result.inks.set(number, scheme.resolve('tx1') ?? '000000');
    const patterns = manifest.patterns.filter((pattern) => pattern.sourceSlide === number);
    const shapeOf = new Map(slide.shapes.map((shape) => [shape.id, shape]));

    // Panels and rules: drawn shapes that are not the whole slide.
    const filled: Array<{ shape: ShapeInfo; fill: string }> = [];
    const texts = slide.shapes.filter((shape) => shape.kind === 'text');
    const slideAreas: Array<{ shape: ShapeInfo; area: Area }> = [];
    for (const shape of slide.shapes) {
      const element = elements.get(shape.id);
      if (!element || (shape.kind !== 'shape' && shape.kind !== 'line' && shape.kind !== 'text') || shape.placeholder !== undefined) continue;
      if (element.localName !== 'sp' && element.localName !== 'cxnSp') continue;
      const props = childElement(element, NS.p, 'spPr');
      const fill = fillColor(props, scheme);
      const line = props && childElement(props, NS.a, 'ln');
      const outlined = !!line && !childElement(line, NS.a, 'noFill') && !!childElement(line, NS.a, 'solidFill');
      if (fill && shape.kind !== 'line') filled.push({ shape, fill });
      if (shape.kind === 'text' || hasRelationship(element) || (shape.box.w >= 0.9 && shape.box.h >= 0.9)) continue;
      const preset = (props && childElement(props, NS.a, 'prstGeom')?.getAttribute('prst')) ?? (element.localName === 'cxnSp' ? 'line' : 'custom');
      const thin = shape.box.h <= 0.008 || shape.box.w <= 0.008;
      if (thin) {
        if (!(shape.box.w > shape.box.h) || (!outlined && !fill)) continue;
        const signature = `${preset}|${fill ?? ''}|${outlined ? (fillColor(line ?? undefined, scheme) ?? 'line') : ''}`;
        const rule = ruleBySignature.get(signature);
        if (rule) rule.uses++;
        else ruleBySignature.set(signature, { slide: number, shapeId: shape.id, drawn: shape.box, preset, outlined, holds: 0, uses: 1, ...(fill ? { fill } : {}) });
        continue;
      }
      if (!fill && !outlined) continue;
      if (!/^(rect|roundRect|round\w*Rect|snip\w*Rect)$/.test(preset)) continue;
      const holds = texts.filter((text) => within(text.box, shape.box)).length;
      const signature = `${preset}|${fill ?? ''}|${outlined}`;
      let area = areaBySignature.get(signature);
      if (area) {
        area.uses++;
        area.holds += holds;
      } else {
        area = { slide: number, shapeId: shape.id, drawn: shape.box, preset, outlined, holds, uses: 1, ...(fill ? { fill } : {}) };
        areaBySignature.set(signature, area);
      }
      slideAreas.push({ shape, area });
    }
    // The space between neighbours that share a top edge is the template's own gutter.
    const tops = new Map<string, ShapeInfo[]>();
    for (const { shape } of slideAreas) {
      const top = `${Math.round(shape.box.y * 200)}|${Math.round(shape.box.h * 200)}`;
      tops.set(top, [...(tops.get(top) ?? []), shape]);
    }
    for (const row of tops.values()) {
      const sorted = row.slice().sort((a, b) => a.box.x - b.box.x);
      for (let i = 1; i < sorted.length; i++) {
        const space = (sorted[i] as ShapeInfo).box.x - right((sorted[i - 1] as ShapeInfo).box);
        if (space > 0.004 && space < 0.08) result.gaps.push(space);
      }
    }

    // Text styles, as the slots of the layouts use them.
    const slots = new Map<string, { names: string[]; budget: number; title: boolean; styleRun: 'first' | 'last' }>();
    let titleBox: Box | undefined;
    const skip = new Set<string>();
    for (const pattern of patterns) {
      for (const slot of pattern.slots) {
        if (slot.kind === 'image' && slot.imageKind === 'photo') {
          const shape = shapeOf.get(slot.shapeId);
          const element = elements.get(slot.shapeId);
          const geometry = element && path(element, [NS.p, 'spPr'], [NS.a, 'prstGeom']);
          // A photo that runs off the slide is cut by the slide, not by a shape of its own.
          if (shape && geometry && shape.groupId === undefined && shape.box.x > TOL && shape.box.y > TOL && right(shape.box) < 1 - TOL && bottom(shape.box) < 1 - TOL) {
            const adjust = descendants(geometry, NS.a, 'gd').map((guide) => guide.getAttribute('fmla')).join(',');
            const signature = `${geometry.getAttribute('prst')}|${adjust}`;
            const entry = pictures.get(signature);
            if (entry) entry.uses++;
            else pictures.set(signature, { slide: number, shapeId: slot.shapeId, uses: 1 });
          }
        }
        if (slot.kind === 'table' && !result.table) {
          const element = elements.get(slot.shapeId);
          const cells = element ? descendants(element, NS.a, 'tc') : [];
          const merged = cells.some((cell) => ['gridSpan', 'rowSpan', 'hMerge', 'vMerge'].some((name) => cell.hasAttribute(name)));
          const rows = element ? descendants(element, NS.a, 'tr') : [];
          if (element && element.localName === 'graphicFrame' && rows.length > 0 && !merged && !hasRelationship(element)) {
            const body = rows[Math.min(1, rows.length - 1)] as Element;
            const stated = Number(descendants(body, NS.a, 'rPr')[0]?.getAttribute('sz') ?? descendants(body, NS.a, 'endParaRPr')[0]?.getAttribute('sz') ?? 0) / 100;
            // The most a cell of the sample keeps free at a side, in points: 0.1 inch where it states nothing.
            const margin = cells.reduce((most, cell) => {
              const props = childElement(cell, NS.a, 'tcPr');
              return Math.max(most, ...['marL', 'marR'].map((name) => Number(props?.getAttribute(name) ?? 91440) / 12700));
            }, 0);
            result.table = { slide: number, shapeId: slot.shapeId, slot, rowHeight: Number(body.getAttribute('h') ?? 0), size: slot.bodyText?.size ?? (stated > 0 ? stated : 12), margin: Number.isFinite(margin) && margin > 0 ? margin : 7.2 };
          }
        }
        if (slot.kind === 'chart' && !result.chart && slot.colors && slot.colors.length > 0) {
          result.chart = { colors: slot.colors, ...(slot.text ? { text: slot.text } : {}), ...(slot.lineColor ? { lineColor: slot.lineColor } : {}) };
        }
        if (slot.kind !== 'text') continue;
        if (slot.deckField || slot.empty === 'drop' || slot.empty === 'keep') {
          skip.add(slot.shapeId);
          continue;
        }
        const entry = slots.get(slot.shapeId) ?? { names: [], budget: 0, title: false, styleRun: slot.styleRun ?? 'first' };
        entry.names.push(slot.name);
        entry.budget = Math.max(entry.budget, slot.maxChars ?? 0);
        if (TITLE.test(slot.name)) {
          entry.title = true;
          const shape = shapeOf.get(slot.shapeId);
          const stands = shape ? placeOf(number, shape) : undefined;
          if (stands && !('box' in slot && slot.box)) titleBox ??= stands;
        }
        slots.set(slot.shapeId, entry);
      }
    }
    for (const [shapeId, slot] of slots) {
      const shape = shapeOf.get(shapeId);
      const element = elements.get(shapeId);
      // A placeholder takes its look from its layout and master: there is nothing on the slide to copy it from.
      if (shape && !skip.has(shapeId) && shape.placeholder !== undefined && !slot.title) result.placeholders++;
      if (!shape || !element || skip.has(shapeId) || element.localName !== 'sp' || shape.placeholder !== undefined || hasRelationship(element)) continue;
      const frame = styles.frame(number, shapeId, slot.styleRun);
      const paragraph = frame?.paragraphs[0];
      if (!frame || !paragraph || frame.vertical) continue;
      const fill = fillColor(childElement(element, NS.p, 'spPr'), scheme);
      const under = filled.filter((entry) => entry.shape.id !== shapeId && within(shape.box, entry.shape.box)).sort((a, b) => a.shape.box.w * a.shape.box.h - b.shape.box.w * b.shape.box.h)[0];
      const onPanel = slideAreas.some((entry) => within(shape.box, entry.shape.box));
      const size = Math.round(paragraph.run.size * 2) / 2;
      const color = runColor(element, scheme);
      const face = paragraph.run.latin.name || paragraph.run.cs.name;
      const signature = [size, paragraph.run.bold, paragraph.run.italic, face, color, paragraph.bullet, fill ?? '', slot.title].join('|');
      const at = `${number}|${shapeId}`;
      if (seen.has(at)) continue;
      seen.add(at);
      const known = bySignature.get(signature);
      if (known) {
        known.uses++;
        known.budgets += slot.budget;
        known.budget = Math.max(known.budget, slot.budget);
        for (const name of slot.names) if (!known.names.includes(name)) known.names.push(name);
        continue;
      }
      const candidate: Candidate = {
        slide: number,
        shapeId,
        size,
        bold: paragraph.run.bold,
        italic: paragraph.run.italic,
        bullet: paragraph.bullet,
        color,
        face,
        names: slot.names.slice(),
        budget: slot.budget,
        budgets: slot.budget,
        uses: 1,
        ground: fill ?? under?.fill ?? background,
        onPanel,
      };
      if (fill) candidate.fill = fill;
      // Lines set closer than usual are counted as usual: only what can make a line higher is kept.
      const spacing = paragraph.lineSpacing;
      if ('points' in spacing ? spacing.points > 0 : spacing.percent > 1) candidate.leading = spacing;
      if (under) candidate.panelFill = under.fill;
      if (slot.title) candidate.names.unshift('title');
      bySignature.set(signature, candidate);
    }

    // Where this slide's content sits: under its title, apart from what repeats on every slide.
    // A subtitle is the title's second line, not content: blocks that started on its line would
    // stand pressed under the title, where the template's own layouts leave a line free.
    const title = titleBox;
    const framed = new Set(patterns.flatMap((pattern) => pattern.slots.filter((slot) => slot.kind === 'text' && slot.deckField).map((slot) => slot.shapeId)));
    const subtitles = new Set(patterns.flatMap((pattern) => pattern.slots.filter((slot) => slot.kind === 'text' && SUBTITLE.test(slot.name)).map((slot) => slot.shapeId)));
    const under = slide.shapes.filter(
      (shape) => shape.kind !== 'group' && shape.placeholder === undefined && !framed.has(shape.id) && shape.box.w < 0.98 && shape.box.w > 0 && shape.box.h > 0 && !(title && shape.box === title) && (!title || shape.box.y >= bottom(title) - TOL) && !inFrameBand(shape.box),
    );
    const body = under.filter((shape) => !subtitles.has(shape.id));
    if (title && body.length > 0) result.content.push({ slide: number, top: Math.min(...body.map((shape) => shape.box.y)), under: Math.min(...under.map((shape) => shape.box.y)), bottom: Math.max(...body.map((shape) => bottom(shape.box))), title });
  }

  result.candidates = [...bySignature.values()];
  result.areas = [...areaBySignature.values()];
  result.rules = [...ruleBySignature.values()];
  const picture = [...pictures.values()].sort((a, b) => b.uses - a.uses || a.slide - b.slide)[0];
  if (picture) result.picture = { slide: picture.slide, shapeId: picture.shapeId };
  cache.set(key, result);
  return result;
}

const style = (candidate: Candidate, set?: FreeTextStyle['set']): FreeTextStyle => {
  const out: FreeTextStyle = { slide: candidate.slide, shapeId: candidate.shapeId, size: candidate.size, bold: candidate.bold, italic: candidate.italic, bullet: candidate.bullet, color: candidate.color };
  if (candidate.fill) out.fill = candidate.fill;
  if (candidate.leading) out.leading = candidate.leading;
  if (set) out.set = set;
  return out;
};

const order = (a: Candidate, b: Candidate) => a.slide - b.slide || Number(a.shapeId) - Number(b.shapeId);
const named = (candidate: Candidate, pattern: RegExp) => candidate.names.some((name) => pattern.test(name));
/** Named as a heading: `card1Heading`, `columnHead`. A sub-heading is the line under one. */
const isHeading = (candidate: Candidate) => candidate.names.some((name) => /heading|head\d*$/i.test(name) && !/sub/i.test(name));

/**
 * Chooses a shape of the template for every part of a free slide. A style is taken only where
 * its text reads on what it will be drawn on: the base slide's background.
 */
function chooseStyles(found: Scan, base: { titleSize: number; titleShape: { slide: number; shapeId: string } | undefined; background: string; rtl: boolean }): FreeKit['styles'] | undefined {
  const T = base.titleSize;
  const reads = (candidate: Candidate) => contrast(candidate.color, base.background) >= 3;
  const plain = found.candidates.filter((candidate) => !candidate.fill && !candidate.names.includes('title') && reads(candidate));
  const best = <T>(list: T[], ...by: Array<(a: T, b: T) => number>): T | undefined => list.slice().sort((a, b) => by.reduce((result, compare) => result || compare(a, b), 0))[0];

  const bodies = plain.filter((candidate) => !candidate.bullet && !candidate.bold && !candidate.italic && candidate.size <= 0.6 * T);
  const body = best(bodies.filter((candidate) => candidate.budget >= 80), (a, b) => b.budgets - a.budgets, (a, b) => b.uses - a.uses, order) ?? best(bodies, (a, b) => b.budgets - a.budgets, (a, b) => b.uses - a.uses, order);
  if (!body) {
    // A template whose running text is all bold, bulleted or large still has a text to stand in.
    const any = best(plain.filter((candidate) => candidate.size <= 0.7 * T), (a, b) => b.budgets - a.budgets, order);
    if (!any) return undefined;
    return fallbackStyles(found, any, T, reads, base.rtl);
  }
  return rolesFrom(found, body, T, reads, base.rtl);
}

function fallbackStyles(found: Scan, any: Candidate, T: number, reads: (candidate: Candidate) => boolean, rtl: boolean): FreeKit['styles'] {
  // The stand-in is the candidate made plain: not bold, not slanted, and without the bullet it
  // may have. The roles that are derived from it are plain too; a list is still taken from the
  // template's own bulleted text.
  const plain = { ...any, bold: false, italic: false, bullet: false };
  return rolesFrom(found, plain, T, reads, rtl, style(plain, { bold: false, italic: false, ...(any.bullet ? { bullets: false } : {}) }));
}

function rolesFrom(found: Scan, body: Candidate, T: number, reads: (candidate: Candidate) => boolean, rtl: boolean, bodyStyle = style(body)): FreeKit['styles'] {
  const best = <V>(list: V[], ...by: Array<(a: V, b: V) => number>): V | undefined => list.slice().sort((a, b) => by.reduce((result, compare) => result || compare(a, b), 0))[0];
  const plain = found.candidates.filter((candidate) => !candidate.fill && !candidate.names.includes('title') && reads(candidate));
  const first = (pattern: RegExp) => (a: Candidate, b: Candidate) => Number(named(b, pattern)) - Number(named(a, pattern));
  const near = (size: number) => (a: Candidate, b: Candidate) => Math.abs(a.size - size) - Math.abs(b.size - size);

  const heading = best(
    plain.filter((candidate) => !candidate.bullet && (candidate.bold || candidate.face !== body.face) && candidate.size >= body.size && candidate.size <= 0.8 * T && (candidate.budget === 0 || candidate.budget <= 60) && !named(candidate, /^subtitle$/i)),
    (a, b) => Number(isHeading(b)) - Number(isHeading(a)),
    (a, b) => b.uses - a.uses,
    near(1.25 * body.size),
    order,
  );
  const headingStyle = heading ? style(heading) : { ...bodyStyle, set: { ...bodyStyle.set, bold: true, size: Math.round(body.size * 1.25 * 2) / 2 } };
  const headingSize = heading?.size ?? body.size * 1.25;

  const figure = best(
    plain.filter((candidate) => !candidate.bullet && candidate.size >= 1.3 * headingSize && candidate.budget > 0 && candidate.budget <= 8),
    (a, b) => b.uses - a.uses,
    (a, b) => a.size - b.size,
    order,
  );
  const figureStyle = figure ? style(figure) : { ...headingStyle, set: { ...headingStyle.set, bold: true, size: Math.round(body.size * 2.2 * 2) / 2 } };

  const caption = best(
    plain.filter((candidate) => !candidate.bullet && candidate !== body && (candidate.size <= 0.9 * body.size || (named(candidate, /caption|footnote|note|source|label/i) && candidate.size <= 1.2 * body.size && !candidate.bold))),
    first(/caption|footnote|note|source|label/i),
    (a, b) => Number(b.italic) - Number(a.italic),
    (a, b) => a.size - b.size,
    order,
  );
  const captionStyle = caption ? style(caption) : { ...bodyStyle, set: { ...bodyStyle.set, size: Math.max(5, Math.round(body.size * 0.85 * 2) / 2) } };

  const list = best(plain.filter((candidate) => candidate.bullet), near(body.size), (a, b) => b.uses - a.uses, order);
  const listStyle = list ? style(list, list.size === body.size ? undefined : { size: body.size }) : { ...bodyStyle, set: { ...bodyStyle.set, bullets: true } };

  const quote = best(plain.filter((candidate) => !candidate.bullet && candidate.names.some((name) => /quot/i.test(name) && !/author|source|name|by$/i.test(name))), (a, b) => b.size - a.size, order) ?? best(plain.filter((candidate) => !candidate.bullet && candidate.italic && candidate.size >= 1.3 * body.size && candidate.size <= 0.95 * T), near(1.6 * body.size), order);
  // A template with no quotation of its own has one made from its heading: light, in the colour
  // of running text, so that it does not read as one more heading, and slanted where the script
  // has a slant. A template drawn to be read from the right is set in letters that have none:
  // a slant there is the machine's, and the quotation stands upright.
  const quoteStyle = quote ? style(quote) : { ...headingStyle, set: { ...headingStyle.set, italic: !rtl, bold: false, color: bodyStyle.set?.color ?? body.color, size: Math.round(body.size * 1.6 * 2) / 2 } };

  const styles: FreeKit['styles'] = { body: bodyStyle, heading: headingStyle, figure: figureStyle, caption: captionStyle, list: listStyle, quote: quoteStyle };
  // A note: a text the template sets on a box of its own, lightly filled.
  const notes = found.candidates.filter((candidate) => candidate.fill && luminance(candidate.fill) >= 0.5 && contrast(candidate.color, candidate.fill) >= 3 && !candidate.names.includes('title'));
  const note = best(notes, first(/takeaway|note|bottomLine|conclusion|summary/i), (a, b) => b.uses - a.uses, order);
  if (note) styles.note = style(note);
  return styles;
}

/**
 * The least room a frame leaves for blocks: from where they start under its title down to the
 * bottom of the content area, as a share of the slide's height. A layout made for content leaves
 * more than half the slide (0.54 and up on the templates this was measured on). A cover, a
 * section divider and a closing slide, whose titles sit in the middle of the slide or low on it,
 * leave 0.40 at most. 0.45 lies between the two: no title page is taken for a frame, and no
 * content layout is refused. The engine's own measure, exported for its tests: the package does
 * not export it. What a caller can ask is `freeBaseRefusal`.
 */
export const FREE_FRAME_ROOM = 0.45;
/**
 * A template none of whose layouts leaves that much is stood on all the same, on the layout that
 * leaves most, as long as that is this share of the slide's height. Under it a row of cards has
 * no room for a heading and a line of text.
 */
const FREE_FRAME_LEAST = 0.15;

/** The room a kit's frame leaves for blocks, as a share of the slide's height. */
const roomOf = (kit: FreeKit): number => kit.dense.bottom - kit.band.top;

/** What was worked out for a template, and the manifest it was worked out under. */
interface Bench {
  manifest: TemplateManifest;
  /** The kit, or why there is none, by the frame that was asked for. */
  kits: Map<string, FreeKit | string>;
  /** The kit on the frame of one layout, whatever room it leaves, by the layout's id. */
  frames: Map<string, FreeKit | string>;
}

// A kit is good for the manifest object it was worked out under. A manifest that `changeManifest`
// gave back is another object, and starts afresh whether it is loaded on the open package or put
// on the same template. A manifest object that is edited in place is not seen: that is not supported.
const benches = new WeakMap<Template, Bench>();

function benchOf(template: Template): Bench {
  let bench = benches.get(template);
  if (!bench || bench.manifest !== template.manifest) benches.set(template, (bench = { manifest: template.manifest, kits: new Map(), frames: new Map() }));
  return bench;
}

/** Why a template cannot carry a free slide, as a sentence; or its kit. */
export function freeKitOrReason(template: Template, base?: string): FreeKit | string {
  const wanted = base ?? (template.manifest.free ? template.manifest.free.base : undefined);
  const { kits } = benchOf(template);
  const key = wanted ?? '';
  const known = kits.get(key);
  if (known !== undefined) return known;
  const result = build(template, wanted);
  kits.set(key, result);
  return result;
}

/** What free slides of a template are made with, or `undefined` when the template cannot carry one. */
export function freeKit(template: Template, base?: string): FreeKit | undefined {
  const kit = freeKitOrReason(template, base);
  return typeof kit === 'string' ? undefined : kit;
}

/** The kit on the frame of one layout, worked out once. Whether the frame leaves room enough is for the caller to say. */
function frameOf(template: Template, id: string | undefined): FreeKit | string {
  const { frames } = benchOf(template);
  const key = id ?? '';
  let made = frames.get(key);
  if (made === undefined) frames.set(key, (made = frameKit(template, scan(template), id)));
  return made;
}

const percent = (share: number): string => `${Math.max(0, Math.round(share * 100))}%`;

/**
 * Why a layout whose title sits low is no frame, as a sentence. `instead` is the frame free
 * slides stand on when that one leaves less than a frame should too (a template of title pages
 * only): the sentence then compares with it, and does not speak of a room no layout has.
 */
const lowTitle = (made: FreeKit, instead?: FreeKit): string =>
  `the layout "${made.base}" has its title low on the slide: blocks would get ${percent(roomOf(made))} of the slide's height, and ${instead ? `"${instead.base}" leaves them ${percent(roomOf(instead))}` : `a frame leaves them ${percent(FREE_FRAME_ROOM)} or more`}`;

/** The kit on the frame of one layout when it leaves room enough for blocks, or why it is no frame. */
function usable(template: Template, id: string | undefined, least = FREE_FRAME_ROOM): FreeKit | string {
  const made = frameOf(template, id);
  if (typeof made === 'string' || roomOf(made) >= least - 1e-9) return made;
  return lowTitle(made);
}

const titled = (pattern: PatternDef): boolean => pattern.slots.some((slot) => slot.kind === 'text' && !slot.deckField && TITLE.test(slot.name));

/**
 * Why a layout cannot be the frame of free slides, as a sentence; `undefined` when it can. For a
 * frame that a slide names (`base`) or the manifest does (`free.base`): a slide that names one
 * that cannot be a frame stands on the template's usual one, and this says why. Part of the
 * package's interface: a studio asks it before it offers a layout as the frame.
 */
export function freeBaseRefusal(template: Template, base: string): string | undefined {
  const pattern = template.manifest.patterns.find((entry) => entry.id === base);
  if (!pattern) return `the template has no layout "${base}"`;
  if (pattern.repeat) return `the layout "${base}" has a row of ${pattern.repeat.item ?? 'item'}s, which the software lays out itself`;
  if (!titled(pattern)) return `the layout "${base}" has no title`;
  // What a slide that names it stands on. On a template none of whose layouts leaves the room
  // of a frame that is the layout that leaves most, and it may be this one.
  const stands = freeKitOrReason(template, base);
  if (typeof stands !== 'string' && stands.base === base) return undefined;
  const made = frameOf(template, base);
  if (typeof made === 'string') return made;
  // A layout that was made for one deck, or whose title shape is gone, gives no frame of its own.
  if (made.base !== base) return `the layout "${base}" gives no frame of its own`;
  // It gives a frame, and that frame was passed over: its title sits low.
  return lowTitle(made, typeof stands !== 'string' && roomOf(stands) < FREE_FRAME_ROOM - 1e-9 ? stands : undefined);
}

function build(template: Template, wanted: string | undefined): FreeKit | string {
  const { manifest } = template;
  // In turn: the frame that was asked for, the one the manifest names, and the layout whose title
  // sits where most titles do. A frame that cannot carry a free slide gives way to the next.
  let reason = '';
  const tried = new Set<string>();
  for (const id of [wanted, manifest.free ? manifest.free.base : undefined, undefined]) {
    if (tried.has(id ?? '')) continue;
    tried.add(id ?? '');
    const made = usable(template, id);
    if (typeof made !== 'string') return made;
    reason = made;
  }
  // None of them is a frame: no running text reads on its background, or its title sits low.
  // Then the layout with a title that leaves most room is stood on, the first in the manifest
  // of several alike; and a template with no layout that leaves room enough, on the one that
  // leaves most all the same.
  for (const least of [FREE_FRAME_ROOM, FREE_FRAME_LEAST]) {
    let best: FreeKit | undefined;
    for (const pattern of manifest.patterns as PatternDef[]) {
      if (pattern.repeat || !titled(pattern)) continue;
      const other = usable(template, pattern.id, least);
      if (typeof other !== 'string' && other.base === pattern.id && (!best || roomOf(other) > roomOf(best) + 1e-6)) best = other;
    }
    if (best) return best;
  }
  return reason;
}

/** The kit a free slide has on the frame of one layout, or why that layout gives it none. */
function frameKit(template: Template, found: Scan, id: string | undefined): FreeKit | string {
  const { pkg, manifest } = template;
  const { info } = found;
  const pres = readPresentation(pkg);
  const aspect = info.slideWidth / info.slideHeight;

  const parts = frameParts(template, id, info);
  if (!parts) return 'the template has no layout with a title to stand on';
  const { base, slide, kit } = parts;
  const title = base.slots.find((slot) => slot.kind === 'text' && !slot.deckField && TITLE.test(slot.name));
  const titleShape = title && slide.shapes.find((shape) => shape.id === title.shapeId);
  const background = found.backgrounds.get(slide.number) ?? 'FFFFFF';
  // A title that takes its place from its layout stands where the layout, or the master, puts it,
  // and is never moved: there is no position on the slide to change.
  const fixed = !!titleShape && placeless(titleShape.box);
  const titleBox: Box | undefined = !titleShape ? undefined : fixed ? found.stands(slide.number, titleShape) : title && title.kind === 'text' && title.box ? title.box : titleShape.box;
  if (titleBox && placeless(titleBox)) return `the title of the layout "${base.id}" takes its place from the slide layout, and neither that nor the master states one`;
  const titleSize = (title && title.kind === 'text' ? (title.fontSize ?? title.size) : undefined) ?? titleShape?.fontSizes?.[0] ?? 28;
  const styles = chooseStyles(found, { titleSize, titleShape: titleShape ? { slide: slide.number, shapeId: titleShape.id } : undefined, background, rtl: manifest.direction === 'rtl' });
  if (!styles) {
    // Text in placeholders is passed over, so a template that keeps all its running text in them has no style to offer.
    const boxed = found.candidates.some((candidate) => !candidate.fill && !candidate.names.includes('title'));
    if (!boxed && found.placeholders > 0) return "the template's running text is in placeholders, which a free slide cannot copy: there is no text style to take unless some of it is set in text boxes of their own";
    return "the template's layouts have no running text that reads on this layout's background, so there is no text style to copy";
  }

  // What stays: the frame as a composed layout keeps it, and on a small template, where too few
  // slides say what repeats, whatever of the slide no layout fills and lies in its margins.
  const stays = new Set(parts.stays);
  const pointed = new Set(manifest.patterns.filter((pattern) => pattern.sourceSlide === slide.number).flatMap((pattern) => pattern.slots.map((slot) => slot.shapeId)));
  for (const shape of slide.shapes) {
    if (shape.groupId !== undefined || pointed.has(shape.id) || shape.kind === 'text' || stays.has(shape.id)) continue;
    if (inFrameBand(shape.box) || (shape.box.w >= 0.98 && shape.box.h >= 0.98)) stays.add(shape.id);
  }
  if (title) stays.add(title.shapeId);
  const ref = pres.slides[slide.number - 1];
  const root = ref ? rootElement(parseXml(pkg.readText(ref.part))) : undefined;
  const ids = root ? descendants(root, NS.p, 'cNvPr').map((props) => Number(props.getAttribute('id') ?? 0)).filter(Number.isFinite) : [];

  // The room for blocks, across: the content area of a composed layout. Where that area leaves
  // clearly more margin at the side reading ends than at the side it starts (a hundredth of the
  // slide's width or more), a free slide would sit off centre under a frame that spans both
  // margins. Its room is then widened at that side: as far as the margin at the start side, no
  // further than what stays above the title reaches (a logo, a rule, the page number), and only
  // when nothing that stays lies in the strip that is added.
  const area = widened(
    kit.area,
    manifest.direction,
    slide.shapes.filter((shape) => stays.has(shape.id) && shape.kind !== 'group' && shape.id !== titleShape?.id && !(shape.box.w >= 0.98 && shape.box.h >= 0.98)),
    titleBox,
  );
  // The band the template's own layouts keep their content in. Its top is where the first thing
  // under their title starts. Where the frame of a free slide draws something on the line a
  // subtitle takes (a deck's own text, a rule) that line is the title's: the band starts under
  // it, and what that takes from the room is given back at its foot, as far down as the lowest
  // quarter of those layouts take their content. A free slide has no subtitle of its own, so
  // where nothing stays on that line the band starts on it, as the layouts' first lines do.
  const titleBottom = titleBox ? bottom(titleBox) : area.y;
  const same = found.content.filter((entry) => titleBox && entry.title && Math.abs(entry.title.y - titleBox.y) <= 0.01 && Math.abs(entry.title.h - titleBox.h) <= 0.02);
  const areaBottom = bottom(area);
  const topOf = (tops: number[]) => clamp(median(tops) ?? titleBottom + 0.035 * aspect, titleBottom + 0.012 * aspect, titleBottom + 0.09);
  const first = topOf(same.map((entry) => entry.under));
  const underLine = topOf(same.map((entry) => entry.top));
  const onLine =
    underLine > first + TOL &&
    slide.shapes.some((shape) => stays.has(shape.id) && shape.id !== titleShape?.id && shape.kind !== 'group' && !(shape.box.w >= 0.98 && shape.box.h >= 0.98) && shape.box.w > 0 && shape.box.h > 0 && shape.box.y < underLine - TOL && bottom(shape.box) > first + TOL && shape.box.x < right(area) - TOL && right(shape.box) > area.x + TOL);
  const top = onLine ? underLine : first;
  const lowered = Math.max(0, top - first);
  const bottoms = same.map((entry) => entry.bottom).sort((a, b) => a - b);
  const usual = median(bottoms) ?? areaBottom - 0.03 * aspect;
  const low = bottoms.length > 0 ? (bottoms[Math.min(bottoms.length - 1, Math.ceil(0.75 * (bottoms.length - 1)))] as number) : usual;
  const end = Math.min(areaBottom, Math.max(top + 0.3, Math.min(low, usual + lowered)));
  const gap = clamp(median(found.gaps) ?? 0.02, 0.012, 0.03);

  const result: FreeKit = {
    base: base.id,
    baseSlide: slide.number,
    slideWidth: info.slideWidth,
    slideHeight: info.slideHeight,
    firstId: Math.max(1, ...ids) + 1,
    stays: [...stays],
    frameSlots: base.slots.filter((slot) => slot.kind === 'text' && slot.deckField !== undefined && stays.has(slot.shapeId)).map((slot) => structuredClone(slot)),
    area,
    band: { top: round(top), bottom: round(end) },
    dense: { top: round(Math.min(top, titleBottom + 0.02)), bottom: round(areaBottom) },
    gap: round(gap),
    background,
    ink: found.inks.get(slide.number) ?? '000000',
    direction: manifest.direction,
    textScale: manifest.textScale ?? 1,
    styles,
    figureOnPanel: false,
    chart: found.chart ?? {},
    palette: paletteOf(pkg, manifest),
  };
  const characterWidth = templateTextSetting(manifest, 'characterWidth');
  if (characterWidth !== undefined) result.characterWidth = characterWidth;
  const leastTextSize = templateTextSetting(manifest, 'leastTextSize');
  if (leastTextSize !== undefined) result.leastTextSize = leastTextSize;
  if (titleShape && titleBox) {
    // A title drawn in a narrow box beside other content has the whole width on a free slide,
    // when nothing that stays stands in its way.
    const others = slide.shapes.filter((shape) => stays.has(shape.id) && shape.id !== titleShape.id && shape.kind !== 'group' && !(shape.box.w >= 0.98 && shape.box.h >= 0.98));
    const free = !others.some((shape) => shape.box.y < bottom(titleBox) - TOL && bottom(shape.box) > titleBox.y + TOL && right(shape.box) > area.x + TOL && shape.box.x < right(area) - TOL);
    const movable = !fixed && titleShape.groupId === undefined;
    const narrow = titleBox.w < 0.6 * area.w && free && movable;
    const box = narrow ? { x: area.x, y: titleBox.y, w: area.w, h: titleBox.h } : titleBox;
    const perLine = Math.max(1, Math.floor((box.w * info.slideWidth) / 12700 / (titleSize * 0.55)));
    const linesHeld = Math.max(1, Math.floor((box.h * info.slideHeight) / 12700 / (titleSize * 1.2)));
    const drawn = title && title.kind === 'text' ? title.maxChars : undefined;
    const placed = narrow || (movable && title?.kind === 'text' && title.box !== undefined);
    result.title = {
      shapeId: titleShape.id,
      slot: structuredClone(title as TextSlotDef),
      box: { x: round(box.x), y: round(box.y), w: round(box.w), h: round(box.h) },
      place: placed,
      size: titleSize,
      maxChars: narrow ? Math.max(drawn ?? 0, Math.floor(perLine * linesHeld * 0.85)) : (drawn ?? Math.floor(perLine * linesHeld * 0.85)),
    };
  }
  if (!result.chart.colors || result.chart.colors.length === 0) {
    const colors = chartColors(result.palette);
    if (colors.length > 0) result.chart = { ...result.chart, colors };
  }

  // Panels: a tinted or outlined box that the template sets text on is a card; a strongly filled one a note.
  const holding = found.areas.filter((entry) => entry.holds > 0);
  const calm = holding
    .filter((entry) => (entry.fill ? luminance(entry.fill) >= 0.6 && contrast(styles.body.color, entry.fill) >= 3 : entry.outlined))
    .sort((a, b) => b.uses - a.uses || a.slide - b.slide || Number(a.shapeId) - Number(b.shapeId))[0];
  if (calm) result.panel = { slide: calm.slide, shapeId: calm.shapeId, drawn: calm.drawn, ...(calm.fill ? { fill: calm.fill } : {}) };
  if (!styles.note) {
    const strong = holding.filter((entry) => entry.fill && luminance(entry.fill) < 0.6).sort((a, b) => b.uses - a.uses || a.slide - b.slide)[0];
    const onIt = strong && found.candidates.filter((candidate) => candidate.panelFill === strong.fill && candidate.slide === strong.slide && contrast(candidate.color, strong.fill as string) >= 3).sort((a, b) => Number(b.bold) - Number(a.bold) || b.size - a.size || order(a, b))[0];
    if (strong && onIt) result.notePanel = { slide: strong.slide, shapeId: strong.shapeId, drawn: strong.drawn, fill: strong.fill as string, text: style(onIt) };
  }
  const rule = found.rules.slice().sort((a, b) => b.uses - a.uses || a.slide - b.slide)[0];
  if (rule) result.rule = { slide: rule.slide, shapeId: rule.shapeId, drawn: rule.drawn, ...(rule.fill ? { fill: rule.fill } : {}) };
  const figure = found.candidates.find((candidate) => candidate.slide === styles.figure.slide && candidate.shapeId === styles.figure.shapeId);
  result.figureOnPanel = !!result.panel && !!figure?.onPanel;
  if (found.picture) result.picture = found.picture;
  if (found.table) result.table = found.table;

  // What a studio says outright wins over what the rules chose.
  const stated = manifest.free ? manifest.free.styles : undefined;
  for (const [role, source] of Object.entries(stated ?? {})) {
    // A manifest written by hand may hold anything here; `checkManifest` says so.
    if (!source || typeof source !== 'object') continue;
    const candidate = found.candidates.find((entry) => entry.slide === source.slide && entry.shapeId === source.shape);
    if (candidate && (FREE_STYLE_ROLES as readonly string[]).includes(role)) result.styles[role as FreeStyleRole] = style(candidate);
    else if (role === 'panel') {
      const panel = found.areas.find((entry) => entry.slide === source.slide && entry.shapeId === source.shape);
      if (panel) result.panel = { slide: panel.slide, shapeId: panel.shapeId, drawn: panel.drawn, ...(panel.fill ? { fill: panel.fill } : {}) };
    }
  }
  return result;
}

/**
 * The room of a free slide across the slide: the content area, made wider at the side reading
 * ends where it leaves a hundredth of the slide's width more margin there than at the side
 * reading starts. See where it is called.
 */
export function widened(area: SlideBox, direction: Direction, staying: ShapeInfo[], title: Box | undefined): SlideBox {
  const rtl = direction === 'rtl';
  const start = rtl ? 1 - right(area) : area.x;
  const endNow = rtl ? area.x : 1 - right(area);
  if (!(endNow >= start + 0.01)) return area;
  // What stays above the content says how far the frame itself reaches: a rule, a logo, a page number.
  const head = title ? title.y : area.y;
  const above = staying.filter((shape) => shape.box.w > 0 && shape.box.h > 0 && bottom(shape.box) <= head + TOL);
  const frame = above.length > 0 ? Math.max(0, Math.min(...above.map((shape) => shape.box.x), ...above.map((shape) => 1 - right(shape.box)))) : 0;
  const margin = Math.max(start, frame, 0.02);
  if (!(margin < endNow - TOL)) return area;
  const strip: Box = rtl ? { x: margin, y: head, w: area.x - margin, h: bottom(area) - head } : { x: right(area), y: head, w: 1 - margin - right(area), h: bottom(area) - head };
  const blocked = staying.some((shape) => shape.box.x < right(strip) - TOL && right(shape.box) > strip.x + TOL && shape.box.y < bottom(strip) - TOL && bottom(shape.box) > strip.y + TOL);
  if (blocked) return area;
  return rtl ? { ...area, x: round(margin), w: round(right(area) - margin) } : { ...area, w: round(1 - margin - area.x) };
}

/** The manifest setting that stops free slides being offered when a deck is planned. */
export const freeOffered = (manifest: Pick<TemplateManifest, 'free' | 'patterns'>): boolean => manifest.free !== false;
