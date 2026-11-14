// How the text of a template shape is set: its typefaces, sizes, spacing and the room its box
// gives. Read from the shape and, where the shape is silent, from what it inherits: its layout
// and master placeholders, the master's text styles, the presentation's defaults and the theme.
import { isRelType, readRels, resolveTarget } from '../pptx/package.ts';
import type { PptxPackage } from '../pptx/package.ts';
import { readPresentation } from '../pptx/presentation.ts';
import type { PresentationInfo } from '../pptx/presentation.ts';
import { NS, childElement, childElements, descendants, parseXml, path, rootElement } from '../pptx/xml.ts';
import type { Element } from '../pptx/xml.ts';

/** A typeface as a run states it, with the theme font it came from when it named one. */
export interface Typeface {
  /** The typeface name, or `''` when nothing states one. */
  name: string;
  /** Per-script typefaces of the theme font the run pointed at, for a script the name does not cover. */
  scripts?: ReadonlyMap<string, string>;
}

export interface RunStyle {
  /** Font size in points. */
  size: number;
  /** Whether the run itself states its size. A size that is only inherited is not rescaled by the renderer. */
  ownSize: boolean;
  bold: boolean;
  italic: boolean;
  latin: Typeface;
  ea: Typeface;
  cs: Typeface;
  /** The size in points from which pairs of letters are kerned; 0 when they never are. */
  kern: number;
  /** Extra space after every character, in points. */
  spacing: number;
  /** Text is shown in capitals. */
  caps: boolean;
  /** The language the run is marked with, e.g. `he-IL`; `''` when nothing states one. */
  lang: string;
}

/** A length in points, or a share of the line height. */
export type Spacing = { points: number } | { percent: number };

export interface ParagraphStyle {
  /** Space kept free at the start side of every line, in points. */
  marginStart: number;
  marginEnd: number;
  /** Shift of the first line against the others, in points; negative for a hanging bullet. */
  indent: number;
  bullet: boolean;
  lineSpacing: Spacing;
  before: Spacing;
  after: Spacing;
  run: RunStyle;
}

export interface TextFrame {
  /** Size of the shape in points. */
  width: number;
  height: number;
  insets: { left: number; right: number; top: number; bottom: number };
  /** Whether lines wrap at the box's edge. */
  wrap: boolean;
  /** `shape`: the box grows with its text when edited in PowerPoint. `shrink`: PowerPoint shrinks the text. */
  autofit: 'none' | 'shape' | 'shrink';
  /** Factors a shrinking box already applies to its text: on the font size, and off the line spacing. */
  fontScale: number;
  lineReduction: number;
  /** Text runs top to bottom, so the box's height is the length of a line. */
  vertical: boolean;
  columns: number;
  /** One entry for each paragraph of the sample text, styled as the renderer will style new text. */
  paragraphs: ParagraphStyle[];
  /** The sample text, paragraph by paragraph. */
  sample: string[];
  /** The shape's outline is one whose text area is not known exactly, so the room is close but not exact. */
  approximate: boolean;
}

const EMU = 12700;
const points = (value: string | null | undefined, fallback: number): number => {
  const number = Number(value);
  return value !== null && value !== undefined && value !== '' && Number.isFinite(number) ? number / EMU : fallback;
};

interface ThemeFonts {
  major: Typeface;
  minor: Typeface;
  majorEa: Typeface;
  minorEa: Typeface;
  majorCs: Typeface;
  minorCs: Typeface;
}

interface MasterPart {
  root: Element;
  theme: ThemeFonts;
}

interface SlidePart {
  root: Element;
  shapes: Map<string, Element>;
  layout?: Element;
  master?: MasterPart;
}

const NO_THEME: ThemeFonts = { major: { name: '' }, minor: { name: '' }, majorEa: { name: '' }, minorEa: { name: '' }, majorCs: { name: '' }, minorCs: { name: '' } };

/** Reads the text styling of a template's shapes. Parsed parts are kept, so asking again is cheap. */
export class TextStyles {
  private readonly pkg: PptxPackage;
  private readonly pres: PresentationInfo;
  private readonly defaults: Element | undefined;
  private readonly slides = new Map<number, SlidePart | undefined>();
  private readonly masters = new Map<string, MasterPart>();
  private readonly layouts = new Map<string, { root: Element; master?: MasterPart }>();
  private readonly frames = new Map<string, TextFrame | undefined>();

  constructor(pkg: PptxPackage) {
    this.pkg = pkg;
    this.pres = readPresentation(pkg);
    this.defaults = childElement(rootElement(parseXml(pkg.readText(this.pres.part))), NS.p, 'defaultTextStyle');
  }

  get slideWidth(): number {
    return this.pres.slideWidth / EMU;
  }

  get slideHeight(): number {
    return this.pres.slideHeight / EMU;
  }

  /**
   * The text frame of a shape on a template slide, or `undefined` when the shape is not there or
   * holds no text. `box` is the shape's place in points when the caller knows it better than the
   * shape's own transform does (a shape inside a scaled group).
   */
  frame(slideNumber: number, shapeId: string, styleRun: 'first' | 'last' = 'first', box?: { width: number; height: number }): TextFrame | undefined {
    const key = `${slideNumber}|${shapeId}|${styleRun}|${box ? `${box.width}x${box.height}` : ''}`;
    if (!this.frames.has(key)) this.frames.set(key, this.read(slideNumber, shapeId, styleRun, box));
    return this.frames.get(key);
  }

  private slide(number: number): SlidePart | undefined {
    if (this.slides.has(number)) return this.slides.get(number);
    const ref = this.pres.slides[number - 1];
    let part: SlidePart | undefined;
    if (ref && this.pkg.has(ref.part)) {
      const root = rootElement(parseXml(this.pkg.readText(ref.part)));
      const shapes = new Map<string, Element>();
      for (const props of descendants(root, NS.p, 'cNvPr')) {
        const shape = props.parentNode?.parentNode;
        const id = props.getAttribute('id');
        if (id && shape && shape.nodeType === 1) shapes.set(id, shape as Element);
      }
      part = { root, shapes };
      const layoutRel = readRels(this.pkg, ref.part).find((rel) => isRelType(rel, 'slideLayout'));
      const layout = layoutRel ? this.layout(resolveTarget(ref.part, layoutRel.target)) : undefined;
      if (layout) {
        part.layout = layout.root;
        if (layout.master) part.master = layout.master;
      }
    }
    this.slides.set(number, part);
    return part;
  }

  private layout(part: string): { root: Element; master?: MasterPart } | undefined {
    const known = this.layouts.get(part);
    if (known) return known;
    if (!this.pkg.has(part)) return undefined;
    const entry: { root: Element; master?: MasterPart } = { root: rootElement(parseXml(this.pkg.readText(part))) };
    const masterRel = readRels(this.pkg, part).find((rel) => isRelType(rel, 'slideMaster'));
    const master = masterRel ? this.master(resolveTarget(part, masterRel.target)) : undefined;
    if (master) entry.master = master;
    this.layouts.set(part, entry);
    return entry;
  }

  private master(part: string): MasterPart | undefined {
    const known = this.masters.get(part);
    if (known) return known;
    if (!this.pkg.has(part)) return undefined;
    const themeRel = readRels(this.pkg, part).find((rel) => isRelType(rel, 'theme'));
    const themePart = themeRel ? resolveTarget(part, themeRel.target) : '';
    const entry: MasterPart = {
      root: rootElement(parseXml(this.pkg.readText(part))),
      theme: themePart && this.pkg.has(themePart) ? readTheme(rootElement(parseXml(this.pkg.readText(themePart)))) : NO_THEME,
    };
    this.masters.set(part, entry);
    return entry;
  }

  private read(slideNumber: number, shapeId: string, styleRun: 'first' | 'last', box: { width: number; height: number } | undefined): TextFrame | undefined {
    const slide = this.slide(slideNumber);
    const shape = slide?.shapes.get(shapeId);
    if (!slide || !shape || shape.localName !== 'sp') return undefined;
    const body = childElement(shape, NS.p, 'txBody');
    if (!body) return undefined;
    const paragraphs = childElements(body, NS.a, 'p');
    if (paragraphs.length === 0) return undefined;

    // A placeholder takes what it does not state from the same placeholder on its layout and master.
    const placeholder = descendants(shape, NS.p, 'ph')[0];
    const layoutShape = placeholder && slide.layout ? findPlaceholder(slide.layout, placeholder, false) : undefined;
    const masterShape = placeholder && slide.master ? findPlaceholder(slide.master.root, layoutShape ? (descendants(layoutShape, NS.p, 'ph')[0] ?? placeholder) : placeholder, true) : undefined;
    const inherited = [layoutShape, masterShape].filter((entry): entry is Element => !!entry);

    const lists: Element[] = [];
    for (const source of [shape, ...inherited]) {
      const list = path(source, [NS.p, 'txBody'], [NS.a, 'lstStyle']);
      if (list) lists.push(list);
    }
    if (placeholder) {
      const type = placeholder.getAttribute('type') ?? 'body';
      const styles = slide.master ? childElement(slide.master.root, NS.p, 'txStyles') : undefined;
      const name = type === 'title' || type === 'ctrTitle' ? 'titleStyle' : type === 'body' || type === 'subTitle' || type === 'obj' ? 'bodyStyle' : 'otherStyle';
      const style = styles && childElement(styles, NS.p, name);
      if (style) lists.push(style);
    }
    if (this.defaults) lists.push(this.defaults);
    const theme = slide.master?.theme ?? NO_THEME;

    const bodies = [shape, ...inherited].map((source) => path(source, [NS.p, 'txBody'], [NS.a, 'bodyPr'])).filter((entry): entry is Element => !!entry);
    const bodyAttr = (name: string): string | null => {
      for (const entry of bodies) {
        const value = entry.getAttribute(name);
        if (value !== null && value !== '') return value;
      }
      return null;
    };
    const fit = bodies.map((entry) => childElements(entry, NS.a).find((child) => child.localName === 'normAutofit' || child.localName === 'spAutoFit' || child.localName === 'noAutofit')).find(Boolean);

    let size = box;
    if (!size) {
      const ext = [shape, ...inherited].map((source) => path(source, [NS.p, 'spPr'], [NS.a, 'xfrm'], [NS.a, 'ext'])).find(Boolean);
      size = { width: points(ext?.getAttribute('cx'), 0), height: points(ext?.getAttribute('cy'), 0) };
      // A shape inside groups is drawn at the groups' scale: each maps its children's space onto its own extent.
      for (let parent = shape.parentNode; parent && parent.nodeType === 1 && (parent as Element).localName === 'grpSp'; parent = parent.parentNode) {
        const transform = path(parent as Element, [NS.p, 'grpSpPr'], [NS.a, 'xfrm']);
        const outer = transform && childElement(transform, NS.a, 'ext');
        const inner = transform && childElement(transform, NS.a, 'chExt');
        const scaleX = Number(outer?.getAttribute('cx')) / Number(inner?.getAttribute('cx'));
        const scaleY = Number(outer?.getAttribute('cy')) / Number(inner?.getAttribute('cy'));
        if (Number.isFinite(scaleX) && scaleX > 0) size.width *= scaleX;
        if (Number.isFinite(scaleY) && scaleY > 0) size.height *= scaleY;
      }
    }

    const vert = bodyAttr('vert');
    // Text keeps clear of an outline that is not a rectangle: rounded corners, an ellipse.
    const outline = [shape, ...inherited].map((source) => path(source, [NS.p, 'spPr'], [NS.a, 'prstGeom'])).find(Boolean);
    const inside = textArea(outline, size.width, size.height);
    const frame: TextFrame = {
      width: size.width,
      height: size.height,
      insets: {
        left: points(bodyAttr('lIns'), 7.2) + inside.x,
        right: points(bodyAttr('rIns'), 7.2) + inside.x,
        top: points(bodyAttr('tIns'), 3.6) + inside.y,
        bottom: points(bodyAttr('bIns'), 3.6) + inside.y,
      },
      wrap: bodyAttr('wrap') !== 'none',
      autofit: fit?.localName === 'normAutofit' ? 'shrink' : fit?.localName === 'spAutoFit' ? 'shape' : 'none',
      fontScale: fit?.localName === 'normAutofit' ? Number(fit.getAttribute('fontScale') ?? 100000) / 100000 : 1,
      lineReduction: fit?.localName === 'normAutofit' ? Number(fit.getAttribute('lnSpcReduction') ?? 0) / 100000 : 0,
      vertical: vert !== null && vert !== 'horz',
      columns: Math.max(1, Number(bodyAttr('numCol') ?? 1) || 1),
      paragraphs: paragraphs.map((paragraph) => readParagraph(paragraph, styleRun, lists, theme)),
      sample: paragraphs.map((paragraph) =>
        descendants(paragraph, NS.a, 't')
          .map((text) => text.textContent ?? '')
          .join(''),
      ),
      approximate: !inside.exact,
    };
    // Vertical text and text in columns are laid out in the box turned or divided, which is close.
    if (frame.vertical || frame.columns > 1) frame.approximate = true;
    return frame;
  }
}

/** Shapes whose text area is the whole of their box. */
const RECTANGULAR = new Set(['rect', 'flowChartProcess', 'actionButtonBlank', 'wedgeRectCallout', 'bracketPair', 'bracePair', 'frame', 'line']);
/** One minus the cosine of 45 degrees: how far a quarter circle of radius 1 keeps a square corner out. */
const CORNER = 0.29289;

/**
 * How far a preset outline keeps text from the sides of the shape's box, in points on each side.
 * The common outlines are worked out as PowerPoint defines them; for the rest the whole box is
 * taken and the answer is marked as not exact.
 */
function textArea(outline: Element | undefined, width: number, height: number): { x: number; y: number; exact: boolean } {
  const preset = outline?.getAttribute('prst') ?? 'rect';
  if (RECTANGULAR.has(preset)) return { x: 0, y: 0, exact: true };
  const guides = new Map<string, number>();
  for (const guide of outline ? descendants(outline, NS.a, 'gd') : []) {
    const value = /^val\s+(-?\d+)/.exec(guide.getAttribute('fmla') ?? '');
    if (value) guides.set(guide.getAttribute('name') ?? '', Number(value[1]));
  }
  const short = Math.min(width, height);
  const adjust = (name: string, fallback: number, most = 50000) => (Math.min(most, Math.max(0, guides.get(name) ?? fallback)) / 100000) * short;
  switch (preset) {
    case 'roundRect': {
      const inset = adjust('adj', 16667) * CORNER;
      return { x: inset, y: inset, exact: true };
    }
    case 'round2SameRect':
    case 'round2DiagRect':
    case 'round1Rect':
    case 'snipRoundRect': {
      const inset = Math.max(adjust('adj1', 16667), adjust('adj2', 0)) * CORNER;
      return { x: inset, y: preset === 'round2SameRect' ? 0 : inset, exact: preset === 'round2DiagRect' };
    }
    case 'snip1Rect':
    case 'snip2SameRect':
    case 'snip2DiagRect': {
      const inset = Math.max(adjust('adj1', 16667), adjust('adj2', 0)) / 2;
      return { x: inset, y: preset === 'snip2DiagRect' ? inset : 0, exact: false };
    }
    case 'ellipse':
    case 'flowChartConnector':
      return { x: (width * CORNER) / 2, y: (height * CORNER) / 2, exact: true };
    case 'diamond':
    case 'flowChartDecision':
      return { x: width / 4, y: height / 4, exact: true };
    default:
      return { x: 0, y: 0, exact: false };
  }
}

/** The placeholder on a layout or master that a slide's placeholder stands for. */
function findPlaceholder(root: Element, wanted: Element, onMaster: boolean): Element | undefined {
  const type = wanted.getAttribute('type') ?? (wanted.getAttribute('idx') ? 'body' : 'obj');
  const index = wanted.getAttribute('idx');
  const tree = path(root, [NS.p, 'cSld'], [NS.p, 'spTree']);
  if (!tree) return undefined;
  const candidates = descendants(tree, NS.p, 'ph');
  const shapeOf = (ph: Element | undefined) => {
    const shape = ph?.parentNode?.parentNode?.parentNode;
    return shape && shape.nodeType === 1 ? (shape as Element) : undefined;
  };
  // A master has one placeholder of each kind: centred titles and subtitles stand for its title and body.
  const family = (name: string) => (name === 'ctrTitle' ? 'title' : name === 'subTitle' || name === 'obj' ? 'body' : name);
  if (!onMaster && index !== null) {
    const byIndex = candidates.find((ph) => ph.getAttribute('idx') === index);
    if (byIndex) return shapeOf(byIndex);
  }
  const byType = candidates.find((ph) => (ph.getAttribute('type') ?? 'obj') === type) ?? candidates.find((ph) => family(ph.getAttribute('type') ?? 'obj') === family(type));
  return shapeOf(byType);
}

function readTheme(root: Element): ThemeFonts {
  const scheme = descendants(root, NS.a, 'fontScheme')[0];
  const read = (group: 'majorFont' | 'minorFont', face: 'latin' | 'ea' | 'cs'): Typeface => {
    const fonts = scheme && childElement(scheme, NS.a, group);
    if (!fonts) return { name: '' };
    const scripts = new Map<string, string>();
    for (const font of childElements(fonts, NS.a, 'font')) {
      const script = font.getAttribute('script');
      const typeface = font.getAttribute('typeface');
      if (script && typeface) scripts.set(script, typeface);
    }
    return { name: childElement(fonts, NS.a, face)?.getAttribute('typeface') ?? '', scripts };
  };
  return {
    major: read('majorFont', 'latin'),
    minor: read('minorFont', 'latin'),
    majorEa: read('majorFont', 'ea'),
    minorEa: read('minorFont', 'ea'),
    majorCs: read('majorFont', 'cs'),
    minorCs: read('minorFont', 'cs'),
  };
}

function readSpacing(element: Element | undefined, fallback: Spacing): Spacing {
  const percent = element && childElement(element, NS.a, 'spcPct');
  if (percent) return { percent: Number(percent.getAttribute('val') ?? 100000) / 100000 };
  const fixed = element && childElement(element, NS.a, 'spcPts');
  if (fixed) return { points: Number(fixed.getAttribute('val') ?? 0) / 100 };
  return fallback;
}

function readParagraph(paragraph: Element, styleRun: 'first' | 'last', lists: Element[], theme: ThemeFonts): ParagraphStyle {
  const own = childElement(paragraph, NS.a, 'pPr');
  const level = Math.min(8, Math.max(0, Number(own?.getAttribute('lvl') ?? 0) || 0));
  const levels = lists.map((list) => childElement(list, NS.a, `lvl${level + 1}pPr`)).filter((entry): entry is Element => !!entry);
  const paragraphSources = own ? [own, ...levels] : levels;
  const attr = (name: string): string | null => {
    for (const source of paragraphSources) {
      const value = source.getAttribute(name);
      if (value !== null && value !== '') return value;
    }
    return null;
  };
  const child = (name: string): Element | undefined => {
    for (const source of paragraphSources) {
      const found = childElement(source, NS.a, name);
      if (found) return found;
    }
    return undefined;
  };
  const bullet = paragraphSources.map((source) => childElements(source, NS.a).find((entry) => /^bu(None|Char|AutoNum|Blip)$/.test(entry.localName ?? ''))).find(Boolean);

  // The renderer keeps one run of the template paragraph for its formatting; a paragraph that
  // was empty in the template is formatted by its end-of-paragraph properties.
  const runs = childElements(paragraph, NS.a, 'r');
  const run = styleRun === 'last' ? runs[runs.length - 1] : runs[0];
  const ownRun = run ? childElement(run, NS.a, 'rPr') : childElement(paragraph, NS.a, 'endParaRPr');
  const runSources = [ownRun, ...levels.map((entry) => childElement(entry, NS.a, 'defRPr'))].filter((entry): entry is Element => !!entry);
  const runAttr = (name: string): string | null => {
    for (const source of runSources) {
      const value = source.getAttribute(name);
      if (value !== null && value !== '') return value;
    }
    return null;
  };
  const typeface = (name: 'latin' | 'ea' | 'cs'): Typeface => {
    for (const source of runSources) {
      const stated = childElement(source, NS.a, name)?.getAttribute('typeface');
      if (stated === null || stated === undefined) continue;
      const themed = /^\+(mj|mn)-(lt|ea|cs)$/.exec(stated);
      if (!themed) return { name: stated };
      const major = themed[1] === 'mj';
      if (themed[2] === 'lt') return major ? theme.major : theme.minor;
      if (themed[2] === 'ea') return major ? theme.majorEa : theme.minorEa;
      return major ? theme.majorCs : theme.minorCs;
    }
    // Text that names no typeface anywhere is set in the theme's body font.
    return name === 'latin' ? theme.minor : name === 'ea' ? theme.minorEa : theme.minorCs;
  };

  const size = Number(runAttr('sz') ?? 1800) / 100;
  return {
    marginStart: points(attr('marL'), 0),
    marginEnd: points(attr('marR'), 0),
    indent: points(attr('indent'), 0),
    bullet: !!bullet && bullet.localName !== 'buNone',
    lineSpacing: readSpacing(child('lnSpc'), { percent: 1 }),
    before: readSpacing(child('spcBef'), { points: 0 }),
    after: readSpacing(child('spcAft'), { points: 0 }),
    run: {
      size: size > 0 ? size : 18,
      ownSize: !!ownRun?.getAttribute('sz'),
      bold: runAttr('b') === '1' || runAttr('b') === 'true',
      italic: runAttr('i') === '1' || runAttr('i') === 'true',
      latin: typeface('latin'),
      ea: typeface('ea'),
      cs: typeface('cs'),
      kern: Number(runAttr('kern') ?? 0) / 100,
      spacing: Number(runAttr('spc') ?? 0) / 100,
      caps: runAttr('cap') === 'all',
      lang: runAttr('lang') ?? lists.map((list) => path(list, [NS.a, 'defPPr'], [NS.a, 'defRPr'])?.getAttribute('lang')).find(Boolean) ?? '',
    },
  };
}
