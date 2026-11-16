// How a text sits in a slot of a template: the lines it breaks into at the slot's typeface and
// size, and whether they fit the box. Works from font files, with no PowerPoint and no browser.
import type { PptxPackage } from '../pptx/package.ts';
import type { Template } from '../render/render.ts';
import { detectDirection, marksRightToLeft } from '../render/text.ts';
import { inspectTemplate } from '../template/inspect.ts';
import type { ShapeInfo, TemplateInfo } from '../template/inspect.ts';
import { rowTextIds } from '../render/repeat.ts';
import { COPY_ID_STEP, effectiveSlots, itemCount, scalesWithText, widenFactor } from '../template/manifest.ts';
import type { PatternDef, TextSlotDef } from '../template/manifest.ts';
import { scriptClass, textWidth, themeScript } from './fonts.ts';
import type { FontFace, FontMetrics, FontStyle, ScriptClass } from './fonts.ts';
import { breakLines } from './layout.ts';
import { TextStyles } from './styles.ts';
import type { ParagraphStyle, RunStyle, Spacing, TextFrame, Typeface } from './styles.ts';

/** How far text may be reduced to make it fit. */
export interface ShrinkLimit {
  /** The smallest share of its size the text may be set at, e.g. 0.85. */
  to: number;
  /** A size in points under which text is never set, whatever `to` allows. */
  minSize?: number;
}

export interface FitOptions {
  /** The factor on body text the deck is rendered with: the template's times the deck's. Defaults to 1. */
  textScale?: number;
  /** Allow the text to be set smaller when it does not fit. */
  shrink?: ShrinkLimit;
  /** For a slot of a row of items: how many items the slide has, which can decide how wide an item is. */
  items?: number;
}

/** How a text sits in a slot. */
export interface TextFit {
  /** Whether the text stays inside its box. */
  fits: boolean;
  /** Lines the text takes, over all its paragraphs. */
  lines: number;
  /** Lines of this text the box holds. */
  capacity: number;
  /** Lines that do not fit; 0 when the text fits. */
  overflow: number;
  /** Height the text takes, and the height the box gives it, in points. */
  height: number;
  room: number;
  /** For a box that does not wrap: how much wider than the box the longest line is, as a factor. 1 or less fits. */
  widthRatio: number;
  /** Font size of the first paragraph as measured, in points. */
  size: number;
  /** The font size each paragraph is set at, in points: what the renderer writes when it shrinks the text. */
  sizes: number[];
  /** The factor the text was reduced by to make it fit; 1 when it was not reduced. */
  scale: number;
  /** About how many characters of text like this the box holds. */
  fitsChars: number;
  /** The lines of each paragraph, as they break. */
  paragraphs: string[][];
  /** The width of each of those lines in points, without the space that hangs at its end. */
  widths: number[][];
  /** The width a line of each paragraph may take, in points. */
  lineRoom: number[];
  /**
   * Lines the template's own sample text takes in this slot: how full the designer drew it.
   * 0 when the slot has no sample to compare with (a slot that moves its box or resets its size).
   */
  sampleLines: number;
  /** A typeface or a weight was stood in for by another, so the figures are close but not exact. */
  approximate: boolean;
}

/** A typeface in a weight and slant that the template's text is set in. */
export interface TemplateFont {
  typeface: string;
  bold: boolean;
  italic: boolean;
}

export interface TextMeasurer {
  /**
   * Measures a text in a slot of a pattern. `undefined` when it cannot be measured: the slot's
   * shape holds no text, or a typeface it needs was not supplied.
   */
  fit(pattern: string | PatternDef, slot: string | TextSlotDef, text: string | string[], options?: FitOptions): TextFit | undefined;
  /** The typefaces the template's text slots are set in for which no font was supplied. */
  missingFonts(): TemplateFont[];
}

/** A line that hangs under the box by less than this share of its height is counted as fitting. */
const SLACK = 0.3;
/** Width in points that a line of shrunk text must leave free: two steps of PowerPoint's layout grid. */
const SPARE = 0.25;
/** PowerPoint sets lines 1.2 times the font size apart, whatever the font's own metrics say. */
const LINE = 1.2;
/**
 * A slot's box of the shape's own size, to within this share of the slide, moves the shape
 * without resizing it. A manifest states boxes to four decimal places.
 */
const MOVED = 0.0005;

interface Shared {
  styles: TextStyles;
  info: TemplateInfo;
  shapes: Map<number, Map<string, ShapeInfo>>;
}

const shared = new WeakMap<PptxPackage, Shared>();

function sharedFor(pkg: PptxPackage): Shared {
  let entry = shared.get(pkg);
  if (!entry) {
    entry = { styles: new TextStyles(pkg), info: inspectTemplate(pkg), shapes: new Map() };
    shared.set(pkg, entry);
  }
  return entry;
}

class MissingFont extends Error {
  readonly font: TemplateFont;

  constructor(font: TemplateFont) {
    super(`no font was supplied for ${font.typeface}`);
    this.font = font;
  }
}

/** The theme's name for the script a language is written in, for the languages whose text takes a typeface of its own. */
const LANGUAGE_SCRIPTS: Record<string, string> = { he: 'Hebr', iw: 'Hebr', yi: 'Hebr', ar: 'Arab', fa: 'Arab', ur: 'Arab', ps: 'Arab', ug: 'Arab', sd: 'Arab', ckb: 'Arab', syr: 'Syrc', dv: 'Thaa', th: 'Thai', hi: 'Deva' };

function typefaceName(run: RunStyle, cls: ScriptClass, char: string): string {
  const face: Typeface = cls === 'cs' ? run.cs : cls === 'ea' ? run.ea : run.latin;
  if (face.name) return face.name;
  const script = cls === 'latin' ? undefined : themeScript(char);
  return (script && face.scripts?.get(script)) || run.latin.name;
}

const RTL_LANGUAGE = /^(he|iw|yi|ar|fa|ur|ps|ug|sd|ckb|syr|dv)(-|$)/i;
const DIGIT = /\p{Nd}/u;
const LAST_LETTER = /(\p{L})[^\p{L}]*$/u;

/**
 * Measures pieces of text in the typefaces of one run style, at one size. `rtl` says the run is
 * marked with a right-to-left language, which decides the typeface of what is not a letter:
 * PowerPoint then sets digits in the complex-script face and a space or a punctuation mark in
 * the face of the character before it. In any other run they are all set in the Latin face.
 */
function runMeasure(fonts: FontMetrics, run: RunStyle, size: number, rtl: boolean, flags: { approximate: boolean }): (text: string) => number {
  const style: FontStyle = { bold: run.bold, italic: run.italic };
  const kerning = run.kern > 0 && size >= run.kern;
  const faces = new Map<string, FontFace>();
  const faceFor = (cls: ScriptClass, char: string): FontFace => {
    const name = typefaceName(run, cls, char);
    let face = faces.get(name);
    if (!face) {
      face = name ? fonts.face(name, style) : undefined;
      if (!face) throw new MissingFont({ typeface: name || '(theme font)', ...style });
      faces.set(name, face);
    }
    if (face.substitute) flags.approximate = true;
    return face;
  };
  const widths = new Map<string, number>();

  return (given) => {
    const known = widths.get(given);
    if (known !== undefined) return known;
    const text = run.caps ? given.toUpperCase() : given;
    const chars = [...text];
    let current: ScriptClass = rtl ? 'cs' : 'latin';
    let width = 0;
    let piece = '';
    let pieceFace: FontFace | undefined;
    let pieceClass: ScriptClass = current;
    let pieceLetters = false;
    const flush = () => {
      // Kerning, as PowerPoint applies it: complex-script letters are always kerned with each
      // other, and never with the punctuation, digits or spaces next to them. Latin text is kerned
      // when the run turns kerning on for its size, except in a run marked right-to-left.
      const kerned = pieceClass === 'cs' ? pieceLetters : kerning && !rtl;
      if (piece !== '' && pieceFace) width += textWidth(pieceFace, piece, size, kerned);
      piece = '';
    };
    for (const char of chars) {
      const strong = scriptClass(char);
      if (strong) current = strong;
      else if (!rtl) current = 'latin';
      else if (DIGIT.test(char)) current = 'cs';
      let face = faceFor(current, char);
      const codePoint = char.codePointAt(0) as number;
      if (!face.has(codePoint) && codePoint > 0x20) {
        flags.approximate = true;
        face = fonts.fallback?.(codePoint, style) ?? face;
      }
      // Complex-script letters are shaped apart from what stands between them.
      const letters = strong !== undefined;
      if (face !== pieceFace || current !== pieceClass || (current === 'cs' && letters !== pieceLetters)) flush();
      pieceFace = face;
      pieceClass = current;
      pieceLetters = letters;
      piece += char;
    }
    flush();
    width += run.spacing * chars.length;
    widths.set(given, width);
    return width;
  };
}

const spacing = (value: Spacing, size: number): number => ('points' in value ? value.points : value.percent * LINE * size);

interface Laid {
  lines: string[][];
  widths: number[][];
  lineRoom: number[];
  height: number;
  /** Height of the gaps between paragraphs, which no shortening of the text removes. */
  gaps: number;
  lastPitch: number;
  widest: number;
  firstSize: number;
}

/**
 * Builds a measurer for a template from font metrics. Pass the metrics here, or load the
 * template with them (`loadTemplate(pptx, manifest, { fonts })`), which also makes validation,
 * planning and rendering measure text.
 */
export function textMeasurer(template: Template, fonts: FontMetrics | undefined = template.fonts): TextMeasurer {
  if (!fonts) throw new Error('textMeasurer needs font metrics: pass them, or load the template with fonts');
  const metrics = fonts;
  const { manifest, pkg } = template;
  const { styles, info, shapes } = sharedFor(pkg);
  const missing = new Map<string, TemplateFont>();
  const shapeOf = (slide: number, id: string): ShapeInfo | undefined => {
    let byId = shapes.get(slide);
    if (!byId) {
      byId = new Map((info.slides[slide - 1]?.shapes ?? []).map((shape) => [shape.id, shape]));
      shapes.set(slide, byId);
    }
    // A copied item's shapes exist only once the row is rendered; they are set like the shape they copy.
    return byId.get(id) ?? byId.get(String(Number(id) % COPY_ID_STEP));
  };

  /**
   * The frame a slot's text is measured in. This is the one place that decides the box: every
   * check (validation, shrinking, the design checks) comes through here. `items` is the number
   * of items on the slide: a row with `widen` makes its items wider when it holds fewer of them,
   * and the text of such an item is measured in the box the renderer gives it.
   */
  const frameOf = (pattern: PatternDef, slot: TextSlotDef, items?: number): TextFrame | undefined => {
    const shape = shapeOf(pattern.sourceSlide, slot.shapeId);
    if (!shape) return undefined;
    // The shape's own size is read exactly from the slide; a slot that moves its box states another.
    const box = slot.box;
    const size = box && box.w > 0 && box.h > 0 ? { width: box.w * styles.slideWidth, height: box.h * styles.slideHeight } : undefined;
    const repeat = pattern.repeat;
    if (!size && repeat && items !== undefined) {
      const count = itemCount(pattern, items);
      const texts = rowTextIds(pattern);
      if (texts?.has(slot.shapeId) && widenFactor(repeat, count) > 1) {
        const widened = styles.rowFrame(pattern.sourceSlide, slot.shapeId, slot.styleRun ?? 'first', { key: pattern.id, repeat, count, texts });
        if (widened) return widened;
      }
    }
    return styles.frame(pattern.sourceSlide, shape.id, slot.styleRun ?? 'first', size);
  };

  /** Lays text out in a frame. `size` gives the font size of a paragraph style; `gap` overrides the space after paragraphs. */
  const lay = (frame: TextFrame, paragraphs: string[], size: (style: ParagraphStyle) => number, gap: number | undefined, flags: { approximate: boolean }, spare = 0): Laid => {
    const inner = (frame.vertical ? frame.height - frame.insets.top - frame.insets.bottom : frame.width - frame.insets.left - frame.insets.right) / frame.columns - spare;
    const laid: Laid = { lines: [], widths: [], lineRoom: [], height: 0, gaps: 0, lastPitch: 0, widest: 0, firstSize: 0 };
    paragraphs.forEach((text, index) => {
      const style = frame.paragraphs[Math.min(index, frame.paragraphs.length - 1)] as ParagraphStyle;
      const points = size(style) * frame.fontScale;
      if (index === 0) laid.firstSize = points;
      // The renderer marks a paragraph that starts with a right-to-left letter, or is mostly
      // right-to-left letters, with such a language.
      const direction = detectDirection(text);
      const measure = runMeasure(metrics, style.run, points, marksRightToLeft(text) || RTL_LANGUAGE.test(style.run.lang), flags);
      const rest = inner - style.marginStart - style.marginEnd;
      const first = rest - (style.bullet && style.indent < 0 ? 0 : style.indent);
      // The space after a line hangs outside it, unless it sits inside text that runs against the paragraph.
      // A space between two words of one direction runs with them; between words of different
      // directions it runs with the paragraph.
      const against = (line: string, next: string): boolean => {
        const last = LAST_LETTER.exec(line)?.[1];
        if (last === undefined || direction === undefined) return false;
        const before = detectDirection(last);
        return before !== direction && (detectDirection(next) ?? before) === before;
      };
      const lines = frame.wrap ? breakLines(text, Math.max(first, 1), Math.max(rest, 1), measure, against) : [text];
      laid.lines.push(lines);
      laid.lineRoom.push(Math.round(rest * 100) / 100);
      const widths = lines.map((line) => measure(line.trimEnd()));
      laid.widths.push(widths.map((width) => Math.round(width * 100) / 100));
      for (const width of widths) laid.widest = Math.max(laid.widest, width / Math.max(rest, 1));
      const pitch = 'points' in style.lineSpacing ? style.lineSpacing.points : LINE * points * Math.max(0.1, style.lineSpacing.percent - frame.lineReduction);
      laid.height += lines.length * pitch;
      laid.lastPitch = pitch;
      // The space before the first paragraph and after the last is not drawn.
      const between = (index > 0 ? spacing(style.before, points) : 0) + (index < paragraphs.length - 1 ? (gap ?? spacing(style.after, points)) : 0);
      laid.height += between;
      laid.gaps += between;
    });
    return laid;
  };

  const note = (error: unknown): undefined => {
    if (!(error instanceof MissingFont)) throw error;
    missing.set(`${error.font.typeface}|${error.font.bold}|${error.font.italic}`, error.font);
    return undefined;
  };

  const fit: TextMeasurer['fit'] = (patternRef, slotRef, text, options = {}) => {
    const pattern = typeof patternRef === 'string' ? manifest.patterns.find((entry) => entry.id === patternRef) : patternRef;
    if (!pattern) return undefined;
    let slot: TextSlotDef | undefined;
    if (typeof slotRef === 'string') {
      // A slot of a copied item is not in the pattern's own list, so the widest row is searched.
      const found = effectiveSlots(pattern, pattern.repeat?.max).find((entry) => entry.name === slotRef) ?? pattern.slots.find((entry) => entry.name === slotRef);
      if (found?.kind === 'text') slot = found;
    } else {
      slot = slotRef;
    }
    if (!slot) return undefined;
    const frame = frameOf(pattern, slot, options.items);
    if (!frame) return undefined;
    // The box as the template draws it: the same frame unless the row widened this item.
    const drawn = options.items === undefined ? frame : (frameOf(pattern, slot) ?? frame);
    const paragraphs = (Array.isArray(text) ? text : text.split(/\r?\n/)).map((paragraph) => paragraph.trim()).filter((paragraph) => paragraph !== '');
    if (paragraphs.length === 0) return undefined;
    const chosen = slot;
    const flags = { approximate: frame.approximate };

    // The size a paragraph is rendered at, mirroring the renderer: a deck's text size applies to
    // body text, and only a run that states its size, or a slot that sets one, is rescaled.
    const bodyScale = scalesWithText(chosen) ? (options.textScale ?? 1) : 1;
    const sizeAt = (shrink: number) => (style: ParagraphStyle) => {
      const stated = chosen.fontSize ?? style.run.size;
      const factor = chosen.fontSize !== undefined || style.run.ownSize || shrink !== 1 ? bodyScale * shrink : 1;
      return factor !== 1 || chosen.fontSize !== undefined ? Math.round((stated * 100 * factor) / 50) / 2 : stated;
    };

    try {
      const boxHeight = (frame.vertical ? frame.width - frame.insets.left - frame.insets.right : frame.height - frame.insets.top - frame.insets.bottom) * frame.columns;
      let room = Math.max(0, boxHeight);
      let roomWidth = 1;
      let sampleLines = 0;
      // A box drawn around its sample text holds at least what the sample takes. That does not
      // carry over to a slot that resizes the box or resets the size: it uses the shape
      // differently. A box that only moves the shape, as a mirrored pattern's does, keeps it.
      const own = shapeOf(pattern.sourceSlide, chosen.shapeId)?.box;
      const sameSize = chosen.box !== undefined && own !== undefined && Math.abs(chosen.box.w - own.w) <= MOVED && Math.abs(chosen.box.h - own.h) <= MOVED;
      if ((chosen.box === undefined || sameSize) && chosen.fontSize === undefined) {
        const sample = frame.sample.map((paragraph) => paragraph.trim()).filter((paragraph) => paragraph !== '');
        if (sample.length > 0) {
          try {
            // An item that widens keeps its height, so the room below the first line, and how
            // full the designer drew the box, are those of the box as drawn.
            const laid = lay(drawn, sample, (style) => style.run.size, undefined, { approximate: false });
            room = Math.max(room, laid.height);
            sampleLines = laid.lines.reduce((sum, entry) => sum + entry.length, 0);
            const wide = frame === drawn ? laid : lay(frame, sample, (style) => style.run.size, undefined, { approximate: false });
            roomWidth = Math.max(roomWidth, wide.widest);
          } catch (error) {
            note(error);
          }
        }
      }

      const measured = (shrink: number): TextFit => {
        // A size chosen because the text just fits at it sits on the edge by construction, so a
        // shrunk text is asked to fit with a little width to spare.
        const laid = lay(frame, paragraphs, sizeAt(shrink), chosen.paragraphGap, flags, shrink === 1 ? 0 : SPARE);
        const lines = laid.lines.reduce((sum, entry) => sum + entry.length, 0);
        const slack = SLACK * laid.lastPitch;
        const over = laid.height - room - slack;
        const overflow = over > 0 ? Math.max(1, Math.ceil(over / laid.lastPitch - 1e-6)) : 0;
        const spare = over > 0 ? 0 : Math.floor(-over / laid.lastPitch + 1e-6);
        const widthRatio = frame.wrap ? 1 : laid.widest / roomWidth;
        const chars = paragraphs.reduce((sum, paragraph) => sum + paragraph.length, 0);
        const textHeight = Math.max(1, laid.height - laid.gaps);
        const share = frame.wrap ? Math.max(0, room - laid.gaps) / textHeight : 1 / Math.max(widthRatio, 0.01);
        return {
          fits: overflow === 0 && widthRatio <= 1.005,
          lines,
          capacity: Math.max(0, lines - overflow + spare),
          overflow,
          height: Math.round(laid.height * 100) / 100,
          room: Math.round(room * 100) / 100,
          widthRatio: Math.round(widthRatio * 1000) / 1000,
          size: laid.firstSize,
          sizes: paragraphs.map((_, index) => sizeAt(shrink)(frame.paragraphs[Math.min(index, frame.paragraphs.length - 1)] as ParagraphStyle)),
          scale: shrink,
          // Lines end at word boundaries, so a text cut to the box's share of its height needs a little margin.
          fitsChars: Math.max(1, Math.floor(chars * share * (share < 1 ? 0.95 : 1))),
          paragraphs: laid.lines,
          widths: laid.widths,
          lineRoom: laid.lineRoom,
          sampleLines,
          approximate: flags.approximate,
        };
      };

      let result = measured(1);
      if (!result.fits && options.shrink) {
        const first = sizeAt(1)(frame.paragraphs[0] as ParagraphStyle);
        const floor = Math.max(options.shrink.to, options.shrink.minSize !== undefined && first > 0 ? options.shrink.minSize / first : 0);
        // Sizes go in half points, so the scale steps down by half a point of the first paragraph at a time.
        for (let size = Math.floor(first * 2 - 1) / 2; size > 0 && size / first >= floor - 1e-9; size -= 0.5) {
          result = measured(size / first);
          if (result.fits) break;
        }
      }
      return result;
    } catch (error) {
      return note(error);
    }
  };

  return {
    fit,
    missingFonts() {
      const wanted = new Map<string, TemplateFont>(missing);
      for (const font of templateFonts(template)) {
        if (!metrics.face(font.typeface, font)) wanted.set(`${font.typeface}|${font.bold}|${font.italic}`, font);
      }
      return [...wanted.values()].sort((a, b) => a.typeface.localeCompare(b.typeface) || Number(a.bold) - Number(b.bold) || Number(a.italic) - Number(b.italic));
    },
  };
}

/** The measurer of a template that was loaded with fonts, or `undefined` for one that was not. */
export function measurerFor(template: Template): TextMeasurer | undefined {
  return template.fonts ? textMeasurer(template, template.fonts) : undefined;
}

/**
 * The typefaces a template's text slots are set in, each with the weights and slants used: the
 * font files to supply for measuring. `language` adds the typeface that text of that language is
 * set in when it differs from the Latin one (Hebrew in a template whose theme sends it to
 * another font); it defaults to the manifest's language.
 */
export function templateFonts(template: Template, language: string | undefined = template.manifest.language): TemplateFont[] {
  const { styles } = sharedFor(template.pkg);
  const script = LANGUAGE_SCRIPTS[(language ?? '').toLowerCase().split('-')[0] ?? ''];
  const found = new Map<string, TemplateFont>();
  const add = (typeface: string, run: RunStyle) => {
    if (typeface) found.set(`${typeface}|${run.bold}|${run.italic}`, { typeface, bold: run.bold, italic: run.italic });
  };
  for (const pattern of template.manifest.patterns) {
    for (const slot of pattern.slots) {
      if (slot.kind !== 'text' || slot.empty === 'drop' || slot.empty === 'keep') continue;
      const frame = styles.frame(pattern.sourceSlide, slot.shapeId, slot.styleRun ?? 'first');
      for (const paragraph of frame?.paragraphs ?? []) {
        add(paragraph.run.latin.name, paragraph.run);
        add(paragraph.run.cs.name || (script ? (paragraph.run.cs.scripts?.get(script) ?? '') : ''), paragraph.run);
      }
    }
  }
  return [...found.values()].sort((a, b) => a.typeface.localeCompare(b.typeface) || Number(a.bold) - Number(b.bold) || Number(a.italic) - Number(b.italic));
}
