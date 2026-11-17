// Reading the colours of a template's slides, and changing the colours of a deck's slides.
// A colour in a pptx is a literal (`a:srgbClr`), a name from the theme (`a:schemeClr`, through the
// master's colour map), a system colour or a preset name. All of them are read; a colour that
// changes is written as a literal, with whatever lightens or darkens it kept on it.
import type { PptxPackage } from '../pptx/package.ts';
import { isRelType, readRels, resolveTarget } from '../pptx/package.ts';
import { readPresentation } from '../pptx/presentation.ts';
import { NS, childElement, childElements, descendants, isElement, parseXml, path, removeNode, rootElement, serializeXml } from '../pptx/xml.ts';
import type { Element, Node } from '../pptx/xml.ts';
import type { ChartSlotDef, TemplateManifest } from '../template/manifest.ts';
import { mapPngColors } from '../util/png.ts';
import { buildPalette } from './palette.ts';
import type { ColorMap, ColorUse, TemplatePalette } from './palette.ts';
import { contrast, hexToRgb, parseHex, readable, rgbToHex } from './space.ts';

const COLOR_ELEMENTS = new Set(['srgbClr', 'schemeClr', 'sysClr', 'prstClr']);

/** The preset colour names a deck is likely to hold. A name that is not here is left as it is. */
const PRESETS: Record<string, string> = {
  black: '000000',
  white: 'FFFFFF',
  red: 'FF0000',
  green: '008000',
  blue: '0000FF',
  yellow: 'FFFF00',
  orange: 'FFA500',
  purple: '800080',
  navy: '000080',
  teal: '008080',
  gray: '808080',
  grey: '808080',
  ltGray: 'D3D3D3',
  ltGrey: 'D3D3D3',
  dkGray: 'A9A9A9',
  dkGrey: 'A9A9A9',
  silver: 'C0C0C0',
  maroon: '800000',
  olive: '808000',
  lime: '00FF00',
  aqua: '00FFFF',
  cyan: '00FFFF',
  fuchsia: 'FF00FF',
  magenta: 'FF00FF',
};

/** What a theme colour name means on one slide, layout or master. */
export interface Scheme {
  /** The parts the colours are inherited through, where the package has them. */
  layout?: string;
  master?: string;
  theme?: string;
  /** The colour of a scheme name such as `accent1` or `tx1`, as six hex digits; `undefined` when the theme has none. */
  resolve(name: string): string | undefined;
}

const NO_SCHEME: Scheme = { resolve: () => undefined };

interface PackageColors {
  schemes: Map<string, Scheme>;
  themes: Map<string, Map<string, string>>;
  palettes: Map<string, TemplatePalette>;
}

const shared = new WeakMap<PptxPackage, PackageColors>();

function sharedFor(pkg: PptxPackage): PackageColors {
  let entry = shared.get(pkg);
  if (!entry) shared.set(pkg, (entry = { schemes: new Map(), themes: new Map(), palettes: new Map() }));
  return entry;
}

/** The twelve colours of a theme part, by name. */
function themeColors(pkg: PptxPackage, part: string): Map<string, string> {
  const cache = sharedFor(pkg).themes;
  const known = cache.get(part);
  if (known) return known;
  const colors = new Map<string, string>();
  if (pkg.has(part)) {
    const scheme = descendants(rootElement(parseXml(pkg.readText(part))), NS.a, 'clrScheme')[0];
    for (const entry of scheme ? childElements(scheme, NS.a) : []) {
      const color = childElements(entry, NS.a)[0];
      const hex = color ? literalOf(color, NO_SCHEME) : undefined;
      if (hex) colors.set(entry.localName ?? '', hex);
    }
  }
  cache.set(part, colors);
  return colors;
}

function mapOf(element: Element | undefined): Map<string, string> | undefined {
  if (!element) return undefined;
  const map = new Map<string, string>();
  for (let i = 0; i < element.attributes.length; i++) {
    const attr = element.attributes.item(i);
    if (attr && !attr.name.startsWith('xmlns')) map.set(attr.localName ?? attr.name, attr.value);
  }
  return map.size > 0 ? map : undefined;
}

const override = (root: Element): Map<string, string> | undefined => {
  const element = childElement(root, NS.p, 'clrMapOvr');
  return mapOf(element && childElement(element, NS.a, 'overrideClrMapping'));
};

/**
 * The theme colours as one part of a package sees them: a slide, a layout or a master. A slide
 * inherits through its layout and master; either may map the names anew.
 */
export function schemeOf(pkg: PptxPackage, part: string): Scheme {
  const cache = sharedFor(pkg).schemes;
  const known = cache.get(part);
  if (known) return known;
  if (!pkg.has(part)) return NO_SCHEME;

  const kind = /slideMasters\//.test(part) ? 'master' : /slideLayouts\//.test(part) ? 'layout' : 'slide';
  const related = (from: string, type: string): string | undefined => {
    const rel = readRels(pkg, from).find((candidate) => isRelType(candidate, type));
    const target = rel ? resolveTarget(from, rel.target) : undefined;
    return target && pkg.has(target) ? target : undefined;
  };
  const layout = kind === 'slide' ? related(part, 'slideLayout') : kind === 'layout' ? part : undefined;
  const master = kind === 'master' ? part : layout ? related(layout, 'slideMaster') : undefined;
  const theme = master ? related(master, 'theme') : undefined;
  const rootOf = (path: string | undefined) => (path ? rootElement(parseXml(pkg.readText(path))) : undefined);
  const masterRoot = rootOf(master);
  const names =
    (kind === 'slide' ? override(rootOf(part) as Element) : undefined) ??
    (layout && kind !== 'master' ? override(rootOf(layout) as Element) : undefined) ??
    mapOf(masterRoot && childElement(masterRoot, NS.p, 'clrMap')) ??
    new Map<string, string>();
  const colors = theme ? themeColors(pkg, theme) : new Map<string, string>();

  const scheme: Scheme = { resolve: (name) => colors.get(names.get(name) ?? name) };
  if (layout) scheme.layout = layout;
  if (master) scheme.master = master;
  if (theme) scheme.theme = theme;
  cache.set(part, scheme);
  return scheme;
}

/** The colour an element states, before anything lightens or darkens it; `undefined` when it cannot be told. */
export function literalOf(element: Element, scheme: Scheme): string | undefined {
  switch (element.localName) {
    case 'srgbClr':
      return parseHex(element.getAttribute('val'));
    case 'schemeClr':
      return scheme.resolve(element.getAttribute('val') ?? '');
    case 'sysClr': {
      const val = element.getAttribute('val');
      return parseHex(element.getAttribute('lastClr')) ?? (val === 'windowText' ? '000000' : val === 'window' ? 'FFFFFF' : undefined);
    }
    case 'prstClr':
      return PRESETS[element.getAttribute('val') ?? ''];
    default:
      return undefined;
  }
}

export const isColor = (element: Element): boolean => element.namespaceURI === NS.a && COLOR_ELEMENTS.has(element.localName ?? '');

function colorElements(root: Element): Element[] {
  return Array.from(root.getElementsByTagNameNS(NS.a, '*')).filter((element) => isColor(element as Element)) as Element[];
}

const ancestor = (element: Element, test: (node: Element) => boolean): Element | undefined => {
  for (let node: Node | null = element.parentNode; node && isElement(node); node = node.parentNode) if (test(node)) return node;
  return undefined;
};
const named = (ns: string, ...locals: string[]) => (node: Element) => node.namespaceURI === ns && locals.includes(node.localName ?? '');

/**
 * Changes the colours of everything under `root` (a slide, a layout, a master or a theme) with
 * `map`. A literal colour that changes has its value set. A colour named through the theme is
 * left alone where the theme of the output will show the right colour for it by itself:
 * `themed` says what that theme makes of a template colour (the deck's colours, when the deck
 * has any). Where it would not, the named colour becomes a literal, and what was inside it (a
 * tint, a shade, a transparency) stays inside. A placeholder colour (`phClr`) and the colour
 * effects of a picture are never touched. Returns how many elements were rewritten.
 */
export function recolorTree(root: Element, map: ColorMap, scheme: Scheme, themed?: ColorMap): number {
  let rewritten = 0;
  for (const element of colorElements(root)) {
    if (ancestor(element, named(NS.a, 'blip'))) continue;
    const from = literalOf(element, scheme);
    if (!from) continue;
    const to = map(from);
    if (element.localName === 'srgbClr') {
      if (to === from) continue;
      element.setAttribute('val', to);
      rewritten++;
      continue;
    }
    // Only a scheme colour follows the theme; a system or preset colour is what it says.
    const shown = element.localName === 'schemeClr' && themed ? themed(from) : from;
    if (to === shown) continue;
    const doc = element.ownerDocument;
    const parent = element.parentNode;
    if (!doc || !parent) continue;
    const literal = doc.createElementNS(NS.a, `${element.prefix ? `${element.prefix}:` : ''}srgbClr`);
    literal.setAttribute('val', to);
    for (const child of childElements(element)) literal.appendChild(child);
    parent.replaceChild(literal, element);
    rewritten++;
  }
  return rewritten;
}

interface Rect {
  x: number;
  y: number;
  w: number;
  h: number;
}

const FILL_NAMES = ['solidFill', 'gradFill', 'noFill', 'blipFill', 'pattFill', 'grpFill'];

/** Where every shape of a slide is, in slide units, with the scaling of the groups around it worked in. */
function shapeBoxes(tree: Element): Map<Element, Rect> {
  const boxes = new Map<Element, Rect>();
  const number = (element: Element | undefined, name: string) => Number(element?.getAttribute(name) ?? 0) || 0;
  const walk = (container: Element, sx: number, sy: number, tx: number, ty: number): void => {
    for (const shape of childElements(container, NS.p)) {
      const local = shape.localName ?? '';
      if (!SHAPES.includes(local)) continue;
      const xfrm = local === 'grpSp' ? path(shape, [NS.p, 'grpSpPr'], [NS.a, 'xfrm']) : local === 'graphicFrame' ? childElement(shape, NS.p, 'xfrm') : path(shape, [NS.p, 'spPr'], [NS.a, 'xfrm']);
      const off = xfrm && childElement(xfrm, NS.a, 'off');
      const ext = xfrm && childElement(xfrm, NS.a, 'ext');
      boxes.set(shape, { x: number(off, 'x') * sx + tx, y: number(off, 'y') * sy + ty, w: number(ext, 'cx') * sx, h: number(ext, 'cy') * sy });
      if (local !== 'grpSp') continue;
      const childOff = xfrm && childElement(xfrm, NS.a, 'chOff');
      const childExt = xfrm && childElement(xfrm, NS.a, 'chExt');
      const kx = number(childExt, 'cx') > 0 ? number(ext, 'cx') / number(childExt, 'cx') : 1;
      const ky = number(childExt, 'cy') > 0 ? number(ext, 'cy') / number(childExt, 'cy') : 1;
      walk(shape, sx * kx, sy * ky, tx + sx * (number(off, 'x') - number(childOff, 'x') * kx), ty + sy * (number(off, 'y') - number(childOff, 'y') * ky));
    }
  };
  walk(tree, 1, 1, 0, 0);
  return boxes;
}

/** The colour a fill shows: a solid fill's colour, or the first colour of a gradient. `undefined` for no fill, a picture, or none stated. */
export function fillColor(parent: Element | undefined, scheme: Scheme): string | undefined {
  const solid = parent && childElement(parent, NS.a, 'solidFill');
  const gradient = parent && childElement(parent, NS.a, 'gradFill');
  const holder = solid ?? (gradient && descendants(gradient, NS.a, 'gs')[0]);
  const color = holder && childElements(holder, NS.a).find(isColor);
  return color ? literalOf(color, scheme) : undefined;
}

/**
 * A fill as the guard of readable text sees it: the colours it shows, each with the share of the
 * fill it stands for. A solid fill is one colour. A gradient is its stops: between two stops the
 * colour runs evenly from one to the other, so a stop stands for half of the stretch on either
 * side of it and for all of what lies beyond the first and the last. `fillColor` keeps giving the
 * first stop: free slides choose their styles by it, and must not change with this.
 */
export type Surface = Array<{ hex: string; share: number }>;

function surfaceOf(parent: Element | undefined, scheme: Scheme): Surface | undefined {
  const gradient = parent && !childElement(parent, NS.a, 'solidFill') ? childElement(parent, NS.a, 'gradFill') : undefined;
  if (!gradient) {
    const hex = fillColor(parent, scheme);
    return hex ? [{ hex, share: 1 }] : undefined;
  }
  const stops: Array<{ hex: string; at: number }> = [];
  for (const stop of descendants(gradient, NS.a, 'gs')) {
    const color = childElements(stop, NS.a).find(isColor);
    const hex = color && literalOf(color, scheme);
    if (hex) stops.push({ hex, at: Math.min(1, Math.max(0, Number(stop.getAttribute('pos') ?? 0) / 100000 || 0)) });
  }
  stops.sort((a, b) => a.at - b.at);
  if (stops.length === 0) return undefined;
  return stops.map((stop, index) => {
    const before = stops[index - 1];
    const after = stops[index + 1];
    return { hex: stop.hex, share: (before ? (stop.at - before.at) / 2 : stop.at) + (after ? (after.at - stop.at) / 2 : 1 - stop.at) };
  });
}

/**
 * The one colour a surface comes to, with its colours as they are or as `map` changes them. Each
 * colour is changed and then they are mixed, not the mix changed, so that an exact replacement
 * of one stop counts.
 */
function shown(surface: Surface, map?: ColorMap): string {
  const first = (surface[0] as Surface[number]).hex;
  if (surface.length === 1) return map ? map(first) : first;
  const sum = [0, 0, 0];
  let total = 0;
  for (const { hex, share } of surface) {
    const rgb = hexToRgb(map ? map(hex) : hex);
    for (let channel = 0; channel < 3; channel++) sum[channel] = (sum[channel] as number) + (rgb[channel] as number) * share;
    total += share;
  }
  if (!(total > 0)) return map ? map(first) : first;
  return rgbToHex((sum[0] as number) / total, (sum[1] as number) / total, (sum[2] as number) / total);
}

/** The surface of a slide's own background, when it states one. */
function backgroundSurface(root: Element, scheme: Scheme): Surface | undefined {
  const background = path(root, [NS.p, 'cSld'], [NS.p, 'bg']);
  if (!background) return undefined;
  const own = ownBackground(root, scheme);
  return surfaceOf(childElement(background, NS.p, 'bgPr'), scheme) ?? (own ? [{ hex: own, share: 1 }] : undefined);
}

/** The colour of a slide's own background, when it states one. */
export function ownBackground(root: Element, scheme: Scheme): string | undefined {
  const background = path(root, [NS.p, 'cSld'], [NS.p, 'bg']);
  if (!background) return undefined;
  const reference = childElement(background, NS.p, 'bgRef');
  const color = reference && childElements(reference, NS.a).find(isColor);
  return fillColor(childElement(background, NS.p, 'bgPr'), scheme) ?? (color ? literalOf(color, scheme) : undefined);
}

/** Sets the colour of a run of text outright, where a run keeps it: after its outline and before everything else. */
function setRunColor(run: Element, hex: string): void {
  const doc = run.ownerDocument;
  if (!doc) return;
  const prefix = run.prefix ? `${run.prefix}:` : '';
  let props = childElement(run, NS.a, 'rPr');
  if (!props) {
    props = doc.createElementNS(NS.a, `${prefix}rPr`);
    run.insertBefore(props, run.firstChild);
  }
  for (const name of FILL_NAMES) {
    const old = childElement(props, NS.a, name);
    if (old) removeNode(old);
  }
  const fill = doc.createElementNS(NS.a, `${prefix}solidFill`);
  const color = doc.createElementNS(NS.a, `${prefix}srgbClr`);
  color.setAttribute('val', hex);
  fill.appendChild(color);
  const line = childElement(props, NS.a, 'ln');
  props.insertBefore(fill, line ? line.nextSibling : props.firstChild);
}

interface RunOnSurface {
  run: Element;
  /** The colour of the text and of what lies behind it, as the template has them. */
  text: string;
  behind: Surface;
}

/** Every run of text on a slide, and every field (a slide number, a date), with what lies behind it: its own box's fill, a cell's, a filled shape under it, or the background. */
function runsOnSurfaces(root: Element, scheme: Scheme, background: Surface): RunOnSurface[] {
  const tree = path(root, [NS.p, 'cSld'], [NS.p, 'spTree']);
  if (!tree) return [];
  const boxes = shapeBoxes(tree);
  const order = [...boxes.keys()];
  const fills = order.map((shape) => (shape.localName === 'sp' ? surfaceOf(childElement(shape, NS.p, 'spPr'), scheme) : undefined));
  const inherited = scheme.resolve('tx1') ?? '000000';
  const behindShape = (index: number): Surface => {
    const own = fills[index];
    if (own) return own;
    const box = boxes.get(order[index] as Element) as Rect;
    const cx = box.x + box.w / 2;
    const cy = box.y + box.h / 2;
    // Shapes are drawn in the order they are listed: the last filled one before this that covers its middle lies right under it.
    for (let under = index - 1; under >= 0; under--) {
      const fill = fills[under];
      const area = boxes.get(order[under] as Element) as Rect;
      if (fill && area.w > 0 && area.h > 0 && cx >= area.x && cx <= area.x + area.w && cy >= area.y && cy <= area.y + area.h) return fill;
    }
    return background;
  };
  const found: RunOnSurface[] = [];
  order.forEach((shape, index) => {
    if (shape.localName !== 'sp' && shape.localName !== 'graphicFrame') return;
    const behind = behindShape(index);
    // A field is text as a run is, and keeps its colour in the same place.
    for (const run of [...descendants(shape, NS.a, 'r'), ...descendants(shape, NS.a, 'fld')]) {
      const cell = ancestor(run, named(NS.a, 'tc'));
      const props = childElement(run, NS.a, 'rPr');
      const stated = props && childElement(props, NS.a, 'solidFill');
      const color = stated && childElements(stated, NS.a).find(isColor);
      const text = color ? literalOf(color, scheme) : inherited;
      if (text) found.push({ run, text, behind: (cell && surfaceOf(childElement(cell, NS.a, 'tcPr'), scheme)) ?? behind });
    }
  });
  return found;
}

/**
 * Changes the colours of one slide, and keeps its text readable. A new colour can leave a text
 * hard to read that read well before: white on a fill that became pale, or a heading that became
 * yellow on white. Such a text keeps its hue and is made darker or lighter until it reads again.
 * `background` is the colour behind the slide when it states none of its own, or the whole of a
 * gradient there (see {@link inheritedBackground}); `themed` is what the theme of the output makes
 * of a template colour (see {@link recolorTree}).
 *
 * A text on a gradient is judged against the gradient as a whole, every stop by the share of the
 * fill it stands for, and not against its first stop: a title in black on the light end of a
 * gradient whose other end turns dark stays black.
 */
export function recolorSlide(root: Element, map: ColorMap, scheme: Scheme, background: string | Surface = 'FFFFFF', themed?: ColorMap): void {
  const runs = runsOnSurfaces(root, scheme, backgroundSurface(root, scheme) ?? (typeof background === 'string' ? [{ hex: background, share: 1 }] : background));
  recolorTree(root, map, scheme, themed);
  for (const { run, text, behind } of runs) {
    const was = shown(behind);
    const before = contrast(text, was);
    const nowText = map(text);
    const nowBehind = shown(behind, map);
    if ((nowText === text && nowBehind === was) || before < 3 || contrast(nowText, nowBehind) >= Math.min(before, 3)) continue;
    setRunColor(run, readable(nowText, nowBehind, Math.min(before, 4.5)));
  }
}

/**
 * The background a slide inherits from its layout, or from its master when the layout has none:
 * the element, the colour it shows (of a gradient, the first), and the whole of it as the guard
 * of readable text sees it. `undefined` when neither states one.
 */
export function inheritedBackground(pkg: PptxPackage, slidePart: string): { element: Element; color: string | undefined; surface?: Surface } | undefined {
  const scheme = schemeOf(pkg, slidePart);
  for (const part of [scheme.layout, scheme.master]) {
    if (!part) continue;
    const root = rootElement(parseXml(pkg.readText(part)));
    const element = path(root, [NS.p, 'cSld'], [NS.p, 'bg']);
    if (element) {
      const surface = backgroundSurface(root, schemeOf(pkg, part));
      return { element, color: ownBackground(root, schemeOf(pkg, part)), ...(surface ? { surface } : {}) };
    }
  }
  return undefined;
}

/**
 * Gives a slide the background it inherits as its own, so that colours asked for this one slide
 * reach it. A background that is a picture stays where it is: a slide cannot point at a layout's
 * picture. Returns whether the slide has a background of its own afterwards.
 */
export function ownTheBackground(root: Element, inherited: Element | undefined): boolean {
  const slide = childElement(root, NS.p, 'cSld');
  if (!slide) return false;
  if (childElement(slide, NS.p, 'bg')) return true;
  const doc = root.ownerDocument;
  if (!inherited || !doc) return false;
  const pointsAway = [inherited, ...Array.from(inherited.getElementsByTagNameNS('*', '*'))].some((element) => {
    for (let i = 0; i < element.attributes.length; i++) if (element.attributes.item(i)?.namespaceURI === NS.r) return true;
    return false;
  });
  if (pointsAway) return false;
  slide.insertBefore(doc.importNode(inherited, true), slide.firstChild);
  return true;
}

/**
 * Changes the colours of what the given slide parts inherit, in `out` only: their layouts and
 * masters, the themes, and the table styles of the deck. What a slide does not state itself (a
 * background, a placeholder's text colour, the lines of a table style) comes from there. Every
 * other layout of those masters follows too: a person who adds a slide in PowerPoint may pick
 * any of them, and gets the deck's colours. Each part is read from the template and changed
 * once; one with nothing to change is not written, so that it keeps its bytes.
 */
export function recolorInherited(source: PptxPackage, out: PptxPackage, slideParts: Iterable<string>, map: ColorMap): void {
  const parts = new Set<string>();
  let master: string | undefined;
  for (const slide of slideParts) {
    const scheme = schemeOf(source, slide);
    for (const part of [scheme.layout, scheme.master, scheme.theme]) if (part) parts.add(part);
    master ??= scheme.master;
  }
  for (const part of [...parts]) {
    if (!/slideMasters\//.test(part)) continue;
    for (const rel of readRels(source, part)) {
      const layout = isRelType(rel, 'slideLayout') ? resolveTarget(part, rel.target) : undefined;
      if (layout && source.has(layout) && out.has(layout)) parts.add(layout);
    }
  }
  const presentation = readPresentation(source).part;
  const styles = readRels(source, presentation).find((rel) => isRelType(rel, 'tableStyles'));
  const stylesPart = styles ? resolveTarget(presentation, styles.target) : undefined;
  if (stylesPart && source.has(stylesPart)) parts.add(stylesPart);
  for (const part of [...parts].sort()) {
    const doc = parseXml(source.readText(part));
    // A theme names its own colours, so nothing is looked up there. A layout or a master looks
    // them up as its slides do, and the table styles as the first master does.
    const scheme = /theme\//.test(part) ? NO_SCHEME : part === stylesPart ? (master ? schemeOf(source, master) : NO_SCHEME) : schemeOf(source, part);
    if (recolorTree(rootElement(doc), map, scheme, map) > 0) out.write(part, serializeXml(doc));
  }
}

/**
 * Changes the colours of the text styles a presentation states for every text box that states
 * none of its own. `root` is the presentation part as it is about to be written.
 */
export function recolorDefaults(root: Element, map: ColorMap, source: PptxPackage, slideParts: Iterable<string>): void {
  const style = childElement(root, NS.p, 'defaultTextStyle');
  if (!style) return;
  let master: string | undefined;
  for (const slide of slideParts) master ??= schemeOf(source, slide).master;
  recolorTree(style, map, master ? schemeOf(source, master) : NO_SCHEME, map);
}

/**
 * A chart slot with its colours changed. A chart without colours of its own takes the palette's,
 * changed. Its text is text as a run of the slide is: where the new colour no longer reads on
 * the page (`background`, as the template has it) and the old one did, it keeps its hue and is
 * made darker or lighter until it reads, by the test `recolorSlide` uses.
 */
export function recolorChart(slot: ChartSlotDef, map: ColorMap, palette: TemplatePalette, background = 'FFFFFF'): ChartSlotDef {
  const given = slot.colors && slot.colors.length > 0 ? slot.colors : chartColors(palette);
  const next: ChartSlotDef = { ...slot };
  if (given.length > 0) next.colors = given.map((color) => map(parseHex(color) ?? color));
  if (slot.lineColor) next.lineColor = map(parseHex(slot.lineColor) ?? slot.lineColor);
  if (slot.text?.color) {
    const old = parseHex(slot.text.color);
    let color = map(old ?? slot.text.color);
    if (old) {
      const before = contrast(old, background);
      const page = map(background);
      if (before >= 3 && contrast(color, page) < Math.min(before, 3)) color = readable(color, page, Math.min(before, 4.5));
    }
    next.text = { ...slot.text, color };
  }
  return next;
}

/** Colours for the series of a chart, from a template's palette: its main colours first, then the rest it uses. */
export function chartColors(palette: TemplatePalette): string[] {
  const colors = [palette.accent, palette.accent2, palette.ink].filter((color): color is string => color !== undefined);
  for (const family of palette.families) {
    for (const member of family.members) {
      const [red, green, blue] = hexToRgb(member);
      // A tint as pale as the page would not show as a bar.
      if (!colors.includes(member) && (red + green + blue) / 3 < 215) colors.push(member);
    }
  }
  return colors.slice(0, 6);
}

/** A pictogram with its colours changed, or `undefined` when nothing changes or the picture is of a kind that is left alone. */
export function recolorIcon(data: Uint8Array, map: ColorMap): Uint8Array | undefined {
  const hex = (channel: number) => channel.toString(16).padStart(2, '0');
  return mapPngColors(data, (red, green, blue) => {
    const from = `${hex(red)}${hex(green)}${hex(blue)}`.toUpperCase();
    const to = map(from);
    return to === from ? [red, green, blue] : hexToRgb(to);
  });
}

// ---------------------------------------------------------------------------
// The palette of a template

const SHAPES = ['sp', 'cxnSp', 'pic', 'graphicFrame', 'grpSp'];
const FILLS = ['solidFill', 'gradFill', 'noFill', 'blipFill', 'pattFill', 'grpFill'];

/** How much a colour element counts for: a large fill for much, a thin line for little. */
function weightOf(element: Element, slideArea: number): number {
  if (ancestor(element, named(NS.a, 'blip'))) return 0;
  const shape = ancestor(element, named(NS.p, ...SHAPES));
  const style = ancestor(element, named(NS.p, 'style'));
  if (style) {
    // A style names the colour a shape falls back on. It shows only where the shape states no fill of its own.
    const ref = ancestor(element, named(NS.a, 'fillRef'));
    const props = shape && childElement(shape, NS.p, 'spPr');
    if (!ref || ref.getAttribute('idx') === '0' || !props || FILLS.some((fill) => childElement(props, NS.a, fill))) return 0;
  }
  // The guides of the editing grid and a bullet that is not drawn show nothing.
  if (ancestor(element, named(NS.p, 'extLst')) || ancestor(element, named(NS.a, 'extLst'))) return 0;
  const bullet = ancestor(element, named(NS.a, 'buClr'));
  if (bullet) return bullet.parentNode && isElement(bullet.parentNode) && childElement(bullet.parentNode, NS.a, 'buNone') ? 0 : 0.25;
  const run = ancestor(element, named(NS.a, 'rPr', 'endParaRPr', 'defRPr'));
  if (run) {
    if (run.localName !== 'rPr') return 0;
    const text = run.parentNode && isElement(run.parentNode) ? (childElement(run.parentNode, NS.a, 't')?.textContent ?? '') : '';
    return Math.min(10, Math.max(0.25, text.trim().length / 20));
  }
  if (ancestor(element, named(NS.a, 'ln', 'lnL', 'lnR', 'lnT', 'lnB', 'lnRef'))) return 0.5;
  if (ancestor(element, named(NS.p, 'bg'))) return 40;
  if (ancestor(element, named(NS.a, 'tcPr'))) return 1;
  if (ancestor(element, named(NS.a, 'effectLst', 'effectDag', 'scene3d', 'sp3d'))) return 0.1;
  if (!shape) return 0.25;
  const xfrm = descendants(shape, NS.a, 'xfrm')[0] ?? childElement(shape, NS.p, 'xfrm');
  const ext = xfrm && childElement(xfrm, NS.a, 'ext');
  const area = (Number(ext?.getAttribute('cx') ?? 0) * Number(ext?.getAttribute('cy') ?? 0)) / slideArea;
  const stops = ancestor(element, named(NS.a, 'gsLst'));
  const share = stops ? Math.max(1, childElements(stops, NS.a, 'gs').length) : 1;
  return Math.min(30, Math.max(0.5, area * 100)) / share;
}

/** The colours one slide uses, each with how much it counts for. */
export function colorUses(root: Element, scheme: Scheme, slideArea: number): ColorUse[] {
  const uses: ColorUse[] = [];
  for (const element of colorElements(root)) {
    const hex = literalOf(element, scheme);
    const weight = hex ? weightOf(element, slideArea) : 0;
    if (hex && weight > 0) uses.push({ hex, weight });
  }
  return uses;
}

const COMPOSED_SLIDE = /<(?:\w+:)?cSld\b[^>]*\bname="composed /;

/**
 * The colours of a template as its layouts use them: the main colour, the second, the ink, and
 * every family of tints and shades. Read from the slides the manifest's layouts stand on, with
 * no model, once for a template. `manifest.colors` states the main and second colour outright.
 */
export function paletteOf(pkg: PptxPackage, manifest: Pick<TemplateManifest, 'patterns' | 'colors'>): TemplatePalette {
  const pres = readPresentation(pkg);
  const numbers = [...new Set(manifest.patterns.map((pattern) => pattern.sourceSlide))].sort((a, b) => a - b);
  const key = JSON.stringify([numbers, manifest.colors ?? null]);
  const cache = sharedFor(pkg).palettes;
  const known = cache.get(key);
  if (known) return known;
  const uses: ColorUse[] = [];
  const area = pres.slideWidth * pres.slideHeight || 1;
  for (const number of numbers) {
    const slide = pres.slides[number - 1];
    if (!slide || !pkg.has(slide.part)) continue;
    const text = pkg.readText(slide.part);
    // A layout made for a deck copies the template's shapes: counting it again would only repeat them.
    if (COMPOSED_SLIDE.test(text)) continue;
    uses.push(...colorUses(rootElement(parseXml(text)), schemeOf(pkg, slide.part), area));
  }
  const palette = buildPalette(uses, manifest.colors || {});
  cache.set(key, palette);
  return palette;
}

/**
 * The colours of a loaded template as its layouts use them: its main colour (what `accent` in a
 * deck's colours takes the place of), its second colour, the dark colour of its text, and every
 * family of tints and shades. For showing a person what "the template's colour" is, or for
 * offering colours that suit it. Worked out once for a template; no model.
 */
export function templatePalette(template: { pkg: PptxPackage; manifest: Pick<TemplateManifest, 'patterns' | 'colors'> }): TemplatePalette {
  return paletteOf(template.pkg, template.manifest);
}
