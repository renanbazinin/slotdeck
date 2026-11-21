import { NS, childElement, childElements, removeNode } from '../pptx/xml.ts';
import type { Element } from '../pptx/xml.ts';
import type { Direction, TableTextStyle } from '../template/manifest.ts';
import { isRtlLanguage, leftToRightTag } from '../util/language.ts';
import { leadsInto } from '../deck/headings.ts';

export interface TextOptions {
  /** Swap left and right alignment (the deck reads the other way from the template). */
  flipAlign: boolean;
  /** Which run of the template paragraph supplies the formatting. Defaults to the first. */
  styleRun?: 'first' | 'last';
  /** Set the paragraph alignment outright, whatever the template has. */
  align?: 'l' | 'r' | 'ctr';
  /** Formatting to put on the text, for places where the template defines none. */
  style?: TableTextStyle;
  /**
   * The direction the deck reads. In a right-to-left deck a paragraph with any right-to-left
   * letter reads right to left and is marked with a right-to-left language (see
   * {@link paragraphDirection}). Without it, a paragraph's first letter decides.
   */
  direction?: Direction;
  /** The deck's language, which names the language of left-to-right text (see {@link filledLanguage}). */
  language?: string;
  /**
   * The direction the template reads. In a deck that reads left to right on a template that reads
   * right to left, figures without letters ("+12%", "4.6 / 5") are set left to right.
   */
  templateDirection?: Direction;
}

/** What decides the language and direction a filled text is written in. */
export type WritingOptions = Pick<TextOptions, 'direction' | 'language' | 'templateDirection'>;

/** A figure with no letters: "+12%", "-3.5%", "4.6 / 5", "2027 – 2029". */
const FIGURE = (text: string): boolean => /\p{Nd}/u.test(text) && !/\p{L}/u.test(text);

const RTL_LETTER = /[\p{Script=Hebrew}\p{Script=Arabic}\p{Script=Syriac}\p{Script=Thaana}]/u;

/** Base direction of a piece of text from its first letter; `undefined` when it has none (e.g. "78%"). */
export function detectDirection(text: string): Direction | undefined {
  const firstLetter = /\p{L}/u.exec(text)?.[0];
  if (!firstLetter) return undefined;
  return RTL_LETTER.test(firstLetter) ? 'rtl' : 'ltr';
}

/**
 * Whether the run holding a text is marked with a right-to-left language: when the text starts
 * with a right-to-left letter, or when most of its letters are right to left, as in
 * "AI בשירות הלקוחות". In a deck that reads right to left, any right-to-left letter does: a Hebrew
 * sentence that opens with a brand, "Microsoft Teams ככלי עבודה", is a Hebrew sentence. PowerPoint
 * checks the spelling of a run in its language, so Hebrew words in a run marked English are all
 * underlined as misspelled.
 */
export function marksRightToLeft(text: string, direction?: Direction): boolean {
  if (direction === 'rtl' && RTL_LETTER.test(text)) return true;
  if (detectDirection(text) === 'rtl') return true;
  let rtl = 0;
  let other = 0;
  for (const [letter] of text.matchAll(/\p{L}/gu)) {
    if (RTL_LETTER.test(letter)) rtl++;
    else other++;
  }
  return rtl > other;
}

/**
 * The direction a filled paragraph is written in: in a deck that reads right to left, right to
 * left when it holds any right-to-left letter; otherwise the direction of its first letter.
 * `undefined` for text without letters ("78%"), which keeps the direction the template gave it.
 */
export function paragraphDirection(text: string, direction?: Direction): Direction | undefined {
  if (direction === 'rtl' && RTL_LETTER.test(text)) return 'rtl';
  return detectDirection(text);
}

/**
 * The language a filled run states, from the language its template run states (`stated`, `''`
 * when it states none and inherits one). The renderer writes it, and the measurer picks the
 * typefaces of digits and punctuation by it, so the two agree.
 *
 * - A text that {@link marksRightToLeft} keeps a right-to-left language, or gets one from its
 *   letters (`he-IL` for Hebrew).
 * - A text with letters that is not marked right to left, in a run that states a right-to-left
 *   language, gets the left-to-right language of the deck ({@link leftToRightTag}). PowerPoint
 *   orders the words of a run by its language: English left in a run marked Hebrew reads
 *   "reached landing just ,74% 75% shy of our ambitioustarget".
 * - So does a figure without letters ("+12%", "4.6 / 5") in a deck that reads left to right: a
 *   run marked Hebrew puts the sign of "+12%" on its right.
 * - Anything else keeps what it states: other text without letters ("—"), and any figure in a
 *   deck that reads right to left, is set as the template sets it, and a run that states nothing
 *   keeps the language it inherits.
 */
export function filledLanguage(text: string, stated: string, options: WritingOptions = {}): string {
  if (marksRightToLeft(text, options.direction)) return isRtlLanguage(stated) ? stated : rightToLeftTag(text);
  if (isRtlLanguage(stated) && (/\p{L}/u.test(text) || (options.direction === 'ltr' && FIGURE(text)))) return leftToRightTag(options.language);
  return stated;
}

/** The language tag of the first right-to-left letter of a text. */
function rightToLeftTag(text: string): string {
  const letter = RTL_LETTER.exec(text)?.[0] ?? '';
  return /\p{Script=Hebrew}/u.test(letter) ? 'he-IL' : /\p{Script=Arabic}/u.test(letter) ? 'ar-SA' : /\p{Script=Syriac}/u.test(letter) ? 'syr-SY' : 'dv-MV';
}

/**
 * Replaces the text of a shape. Each new paragraph takes its formatting from the
 * template paragraph in the same position (the last one is reused when there are more).
 */
export function setShapeText(shape: Element, paragraphs: string[], options: TextOptions): void {
  const body = childElement(shape, NS.p, 'txBody');
  if (!body) throw new Error('shape has no text body to fill');
  setBodyText(body, paragraphs, options);
}

/** Replaces the paragraphs of a text body (`p:txBody` of a shape, `a:txBody` of a table cell). */
export function setBodyText(body: Element, paragraphs: string[], options: TextOptions): void {
  const existing = childElements(body, NS.a, 'p');
  if (existing.length === 0) throw new Error('text body has no paragraph to take formatting from');

  const fresh = paragraphs.map((text, index) => {
    const proto = existing[Math.min(index, existing.length - 1)] as Element;
    return buildParagraph(proto, text, options);
  });
  for (const paragraph of existing) removeNode(paragraph);
  for (const paragraph of fresh) body.appendChild(paragraph);
}

/**
 * Resizes the text of a shape after it was filled. `scale` multiplies each run's own size;
 * `size` sets every run to that size first (points). Runs that state no size are left alone
 * unless `size` is given. `paragraphGap` sets the space after each paragraph (points).
 */
export function adjustText(shape: Element, change: { scale?: number; size?: number; paragraphGap?: number }): void {
  const body = childElement(shape, NS.p, 'txBody');
  if (!body) return;
  const scale = change.scale ?? 1;
  for (const paragraph of childElements(body, NS.a, 'p')) {
    const props = [...childElements(paragraph, NS.a, 'r').map((run) => childElement(run, NS.a, 'rPr')), childElement(paragraph, NS.a, 'endParaRPr')];
    for (const runProps of props) {
      if (!runProps) continue;
      const base = change.size !== undefined ? change.size * 100 : Number(runProps.getAttribute('sz') ?? 0);
      if (base > 0 && (change.size !== undefined || scale !== 1)) runProps.setAttribute('sz', String(Math.round((base * scale) / 50) * 50));
    }
    if (change.paragraphGap !== undefined) setParagraphGap(paragraph, change.paragraphGap);
  }
}

/**
 * Sets the font size of each paragraph of a filled shape outright, in points: `sizes[n]` for
 * paragraph n, the last one for any paragraph beyond it. Unlike {@link adjustText} it also sizes
 * runs that state no size of their own, so text that inherits its size can be set smaller.
 */
export function setTextSizes(shape: Element, sizes: number[]): void {
  const body = childElement(shape, NS.p, 'txBody');
  if (!body || sizes.length === 0) return;
  childElements(body, NS.a, 'p').forEach((paragraph, index) => {
    const size = String(Math.round((sizes[Math.min(index, sizes.length - 1)] as number) * 100));
    const doc = paragraph.ownerDocument;
    const prefix = paragraph.prefix ? `${paragraph.prefix}:` : '';
    for (const run of childElements(paragraph, NS.a, 'r')) {
      let props = childElement(run, NS.a, 'rPr');
      if (!props && doc) {
        props = doc.createElementNS(NS.a, `${prefix}rPr`);
        run.insertBefore(props, run.firstChild);
      }
      props?.setAttribute('sz', size);
    }
    childElement(paragraph, NS.a, 'endParaRPr')?.setAttribute('sz', size);
  });
}

/**
 * How wide a character is counted where a heading may be set run in: wider than PowerPoint sets any
 * (0.41 to 0.47 of the size), and than the plans count one. A template that states how wide its
 * text is set (`TemplateManifest.characterWidth`) and is wider is counted so.
 */
const HEADING_WIDE = 0.55;
/** How much of a line is counted as filled: lines break at words. */
const HEADING_FULL = 0.85;
/** How much wider a heading set bold is counted than in the face of its text. */
const BOLD = 1.12;
/** The space between a heading set run in and the text it leads (an em space), counted as characters. */
const RUN_IN_GAP = 3;

/**
 * The characters a line of a column so many points wide holds for a heading set run in, at a size:
 * counted wide (`HEADING_WIDE`, or the template's own width where it is wider) and 85% full.
 */
export function headingRoom(column: number, size: number, characterWidth?: number): number {
  return Math.floor(column / (Math.max(HEADING_WIDE, characterWidth ?? 0) * size)) * HEADING_FULL;
}

/** The characters a heading of so many letters takes at the head of the paragraph it is set run in: its letters counted bold, and the em space. */
export const runInChars = (letters: number): number => letters * BOLD + RUN_IN_GAP;

/**
 * Whether a heading is set run in at the head of the paragraph under it, by their texts: the
 * paragraph is no item of a list or row of a table (`leadsInto`), keeps its direction with the
 * heading at its head (`readsAlike`), and the heading's bold letters and the em space take no more
 * than a line that holds `room` characters (`headingRoom`). Where the paragraph draws a bullet, or
 * is a heading itself, the heading is not set run in whatever this says (see `setHeadings`). The
 * plans of free slides ask the same, so that a text is counted as it is set.
 */
export function runsIn(heading: string, next: string, direction: Direction | undefined, room: number): boolean {
  return leadsInto(next) && readsAlike(heading, next, direction) && runInChars([...heading.trim()].length) <= room;
}

/**
 * Sets the headings of a text that was filled (`SlideSpec.headings`), by their positions among its
 * paragraphs (`headingPlaces`): each in bold. A heading over a paragraph of text is set run in, at
 * the head of that paragraph, an em space between them, where its bold letters and the space take
 * no more than a line of the text, counted wider than any plan counts one: the paragraph it joins
 * then takes no more lines than the two took apart, and no heading is left at the foot of a column
 * with its text in the next. The space after the heading goes before the paragraph. The em space
 * opens the paragraph's own text, in its language: PowerPoint places a space by the language of
 * its run, and in the heading's, an English heading over Hebrew text stands glued to it. A heading
 * over an item of a list (by its marker or by the bullet its paragraph draws), a row of a table,
 * another heading or a paragraph that reads the other way, or too long for a line, keeps a line of
 * its own, its bullet taken off, with its space after it moved before it. Run last, after the
 * text's size is set, so that a heading keeps the size and the language its own text was given.
 * `characterWidth` is the template's own (`TemplateManifest.characterWidth`), where it states one.
 */
export function setHeadings(shape: Element, places: readonly number[], paragraphs: readonly string[], direction?: Direction, characterWidth?: number): void {
  const body = childElement(shape, NS.p, 'txBody');
  if (!body || places.length === 0) return;
  const all = childElements(body, NS.a, 'p');
  if (all.length !== paragraphs.length) return;
  const marked = new Set(places);
  for (const at of [...places].sort((a, b) => b - a)) {
    const heading = all[at];
    const next = all[at + 1];
    if (!heading) continue;
    const room = next ? lineRoom(shape, body, heading, next, characterWidth) : 0;
    const runIn = next !== undefined && !marked.has(at + 1) && !drawsBullet(next, body) && runsIn(paragraphs[at] ?? '', paragraphs[at + 1] ?? '', direction, room);
    if (!(runIn && next && joinHeading(heading, next))) headingLine(heading);
  }
}

/**
 * The characters a line of a text holds for a heading set run in (`headingRoom`): the shape's width
 * less its insets, shared among its columns, less the margin of the paragraph under the heading, at
 * the heading's size. 0 where the shape or the size is not stated on the slide: the heading then
 * keeps its own line.
 */
function lineRoom(shape: Element, body: Element, heading: Element, next: Element, characterWidth?: number): number {
  const properties = childElement(shape, NS.p, 'spPr');
  const transform = properties ? childElement(properties, NS.a, 'xfrm') : undefined;
  const extent = transform ? childElement(transform, NS.a, 'ext') : undefined;
  const width = Number(extent?.getAttribute('cx') ?? 0);
  const run = childElements(heading, NS.a, 'r')[0];
  const size = Number((run ? childElement(run, NS.a, 'rPr') : undefined)?.getAttribute('sz') ?? 0) / 100;
  if (!(width > 0 && size > 0)) return 0;
  const box = childElement(body, NS.a, 'bodyPr');
  const emu = (name: string, fallback: number): number => Number(box?.getAttribute(name) ?? fallback);
  const columns = Math.max(1, emu('numCol', 1));
  const margin = Number(childElement(next, NS.a, 'pPr')?.getAttribute('marL') ?? 0);
  const column = (width - emu('lIns', 91440) - emu('rIns', 91440) - (columns - 1) * emu('spcCol', 0)) / columns - margin;
  return headingRoom(column / 12700, size, characterWidth);
}

/** Whether a paragraph draws a bullet: its own, or the one the first level of its shape's list style draws where it states none. */
function drawsBullet(paragraph: Element, body: Element): boolean {
  const BULLETS = ['buChar', 'buAutoNum', 'buBlip'];
  const own = childElement(paragraph, NS.a, 'pPr');
  if (own && childElement(own, NS.a, 'buNone')) return false;
  if (own && BULLETS.some((name) => childElement(own, NS.a, name))) return true;
  const level = Number(own?.getAttribute('lvl') ?? 0) || 0;
  const styles = childElement(body, NS.a, 'lstStyle');
  const listed = styles ? childElement(styles, NS.a, `lvl${level + 1}pPr`) : undefined;
  return !!listed && BULLETS.some((name) => childElement(listed, NS.a, name));
}

/**
 * Whether a paragraph keeps the direction it was written in with a heading at its head: by the rule
 * it was written by (`paragraphDirection`), in a deck that reads `direction`. A Hebrew heading at
 * the head of an English paragraph of a Hebrew deck would make it a Hebrew paragraph.
 */
const readsAlike = (heading: string, next: string, direction: Direction | undefined): boolean => paragraphDirection(`${heading} ${next}`, direction) === paragraphDirection(next, direction);

/** Sets a heading at the head of the paragraph under it (see `setHeadings`). False where it has no single run to move, or the paragraph no run to follow it. */
function joinHeading(heading: Element, next: Element): boolean {
  const runs = childElements(heading, NS.a, 'r');
  const run = runs[0];
  const first = childElements(next, NS.a, 'r')[0];
  const opening = first ? childElement(first, NS.a, 't') : undefined;
  if (!run || runs.length !== 1 || !childElement(run, NS.a, 't') || !first || !opening) return false;
  bold(run);
  opening.textContent = `\u2003${opening.textContent ?? ''}`;
  next.insertBefore(run, first);
  const before = spaceOf(heading, 'spcBef');
  const after = spaceOf(heading, 'spcAft');
  if (before !== undefined && after !== undefined && before + after > 0) setSpace(next, 'spcBef', before + after);
  removeNode(heading);
  return true;
}

/** Sets a heading on a line of its own (see `setHeadings`). */
function headingLine(heading: Element): void {
  for (const run of childElements(heading, NS.a, 'r')) bold(run);
  childElement(heading, NS.a, 'endParaRPr')?.setAttribute('b', '1');
  const props = childElement(heading, NS.a, 'pPr');
  if (props && props.ownerDocument) {
    const marker = ['buChar', 'buAutoNum', 'buBlip'].map((name) => childElement(props, NS.a, name)).find(Boolean);
    if (marker) {
      props.replaceChild(props.ownerDocument.createElementNS(NS.a, `${props.prefix ? `${props.prefix}:` : ''}buNone`), marker);
      if (Number(props.getAttribute('indent') ?? 0) < 0) props.setAttribute('indent', '0');
    }
  }
  const before = spaceOf(heading, 'spcBef');
  const after = spaceOf(heading, 'spcAft');
  if (before !== undefined && after !== undefined && after > 0) {
    setSpace(heading, 'spcBef', before + after);
    setSpace(heading, 'spcAft', 0);
  }
}

function bold(run: Element): void {
  let props = childElement(run, NS.a, 'rPr');
  if (!props && run.ownerDocument) {
    props = run.ownerDocument.createElementNS(NS.a, `${run.prefix ? `${run.prefix}:` : ''}rPr`);
    run.insertBefore(props, run.firstChild);
  }
  props?.setAttribute('b', '1');
}

/** The space a paragraph states before or after it, in points: 0 where it states none, `undefined` where it states it as a share of a line. */
function spaceOf(paragraph: Element, name: 'spcBef' | 'spcAft'): number | undefined {
  const props = childElement(paragraph, NS.a, 'pPr');
  const space = props ? childElement(props, NS.a, name) : undefined;
  if (!space) return 0;
  const points = childElement(space, NS.a, 'spcPts');
  return points ? Number(points.getAttribute('val') ?? 0) / 100 : undefined;
}

/** Sets the space before or after a paragraph, in points, in the order paragraph properties keep: line spacing, space before, space after, then the rest. */
function setSpace(paragraph: Element, name: 'spcBef' | 'spcAft', points: number): void {
  const doc = paragraph.ownerDocument;
  if (!doc) return;
  const prefix = paragraph.prefix ? `${paragraph.prefix}:` : '';
  let props = childElement(paragraph, NS.a, 'pPr');
  if (!props) {
    props = doc.createElementNS(NS.a, `${prefix}pPr`);
    paragraph.insertBefore(props, paragraph.firstChild);
  }
  const old = childElement(props, NS.a, name);
  if (old) removeNode(old);
  const space = doc.createElementNS(NS.a, `${prefix}${name}`);
  const value = doc.createElementNS(NS.a, `${prefix}spcPts`);
  value.setAttribute('val', String(Math.round(points * 100)));
  space.appendChild(value);
  const after = name === 'spcBef' ? (childElement(props, NS.a, 'lnSpc') ?? null) : (childElement(props, NS.a, 'spcBef') ?? childElement(props, NS.a, 'lnSpc') ?? null);
  props.insertBefore(space, after ? after.nextSibling : props.firstChild);
}

function setParagraphGap(paragraph: Element, points: number): void {
  const doc = paragraph.ownerDocument;
  if (!doc) return;
  const prefix = paragraph.prefix ? `${paragraph.prefix}:` : '';
  let props = childElement(paragraph, NS.a, 'pPr');
  if (!props) {
    props = doc.createElementNS(NS.a, `${prefix}pPr`);
    paragraph.insertBefore(props, paragraph.firstChild);
  }
  const old = childElement(props, NS.a, 'spcAft');
  if (old) removeNode(old);
  const gap = doc.createElementNS(NS.a, `${prefix}spcAft`);
  const value = doc.createElementNS(NS.a, `${prefix}spcPts`);
  value.setAttribute('val', String(Math.round(points * 100)));
  gap.appendChild(value);
  // Paragraph properties keep a fixed order: line spacing, space before, space after, then the rest.
  const before = childElement(props, NS.a, 'spcBef') ?? childElement(props, NS.a, 'lnSpc');
  props.insertBefore(gap, before ? before.nextSibling : props.firstChild);
}

function buildParagraph(proto: Element, text: string, options: TextOptions): Element {
  const doc = proto.ownerDocument;
  if (!doc) throw new Error('paragraph is detached from its document');
  const prefix = proto.prefix ? `${proto.prefix}:` : '';
  const paragraph = proto.cloneNode(true) as Element;

  // One run of the template paragraph carries the formatting we keep. A paragraph that was
  // empty in the template has no run, so its formatting comes from the end-of-paragraph properties.
  const runs = childElements(paragraph, NS.a, 'r');
  let run = options.styleRun === 'last' ? runs[runs.length - 1] : runs[0];
  const endProps = childElement(paragraph, NS.a, 'endParaRPr');
  if (!run) {
    run = doc.createElementNS(NS.a, `${prefix}r`);
    if (endProps) {
      const runProps = doc.createElementNS(NS.a, `${prefix}rPr`);
      for (let i = 0; i < endProps.attributes.length; i++) {
        const attr = endProps.attributes.item(i);
        if (attr) runProps.setAttribute(attr.name, attr.value);
      }
      for (const child of childElements(endProps)) runProps.appendChild(child.cloneNode(true));
      run.appendChild(runProps);
    }
  }

  let textNode = childElement(run, NS.a, 't');
  if (!textNode) {
    textNode = doc.createElementNS(NS.a, `${prefix}t`);
    run.appendChild(textNode);
  }
  textNode.textContent = text;
  forgetProofing(childElement(run, NS.a, 'rPr'));
  forgetProofing(endProps);
  if (options.style) applyStyle(run, options.style, prefix);

  // Drop every run, line break and field of the template paragraph, then put ours back.
  for (const child of childElements(paragraph)) {
    if (child.localName !== 'pPr' && child.localName !== 'endParaRPr') removeNode(child);
  }
  paragraph.insertBefore(run, endProps ?? null);

  applyDirection(paragraph, text, options, prefix);
  return paragraph;
}

/**
 * PowerPoint saves the result of its spelling check on each run: `err="1"` says a mistake was
 * found in it, `dirty="0"` says it was checked and needs no new check. Both describe the
 * template's sample text, not ours. Kept, they make PowerPoint underline the new text as
 * misspelled without checking it, until someone right-clicks it. Without them it checks the text.
 */
function forgetProofing(props: Element | null | undefined): void {
  props?.removeAttribute('err');
  props?.removeAttribute('dirty');
}

function applyStyle(run: Element, style: TableTextStyle, prefix: string): void {
  const doc = run.ownerDocument;
  if (!doc) return;
  let props = childElement(run, NS.a, 'rPr');
  if (!props) {
    props = doc.createElementNS(NS.a, `${prefix}rPr`);
    run.insertBefore(props, run.firstChild);
  }
  if (style.size !== undefined) props.setAttribute('sz', String(Math.round(style.size * 100)));
  if (style.bold !== undefined) props.setAttribute('b', style.bold ? '1' : '0');

  // Run properties keep a fixed child order: line, fill, effects, then the typefaces.
  if (style.color) {
    for (const fill of ['noFill', 'solidFill', 'gradFill', 'blipFill', 'pattFill', 'grpFill']) {
      const old = childElement(props, NS.a, fill);
      if (old) removeNode(old);
    }
    const fill = doc.createElementNS(NS.a, `${prefix}solidFill`);
    const color = doc.createElementNS(NS.a, `${prefix}srgbClr`);
    color.setAttribute('val', style.color.replace(/^#/, '').toUpperCase());
    fill.appendChild(color);
    const line = childElement(props, NS.a, 'ln');
    props.insertBefore(fill, line ? line.nextSibling : props.firstChild);
  }
  if (style.font) {
    const after = ['hlinkClick', 'hlinkMouseOver', 'rtl', 'extLst'].map((name) => childElement(props as Element, NS.a, name)).find(Boolean) ?? null;
    for (const script of ['latin', 'ea', 'cs', 'sym']) {
      const old = childElement(props, NS.a, script);
      if (old) removeNode(old);
      const face = doc.createElementNS(NS.a, `${prefix}${script}`);
      face.setAttribute('typeface', style.font);
      props.insertBefore(face, after);
    }
  }
}

function applyDirection(paragraph: Element, text: string, options: TextOptions, prefix: string): void {
  let props = childElement(paragraph, NS.a, 'pPr');
  // A figure without letters keeps the template's direction, except in a deck that reads left to
  // right where the template, or this paragraph, reads right to left: a right-to-left paragraph
  // shows "+12%" as "12%+" and "4.6 / 5" as "5 / 4.6".
  const turned = options.direction === 'ltr' && (options.templateDirection === 'rtl' || props?.getAttribute('rtl') === '1');
  const direction = paragraphDirection(text, options.direction) ?? (turned && FIGURE(text) ? 'ltr' : undefined);
  if (!props) {
    if (!direction && !options.flipAlign && !options.align) {
      writeLanguage(paragraph, text, options, prefix);
      return;
    }
    const doc = paragraph.ownerDocument;
    if (!doc) return;
    props = doc.createElementNS(NS.a, `${prefix}pPr`);
    paragraph.insertBefore(props, paragraph.firstChild);
  }
  if (direction) props.setAttribute('rtl', direction === 'rtl' ? '1' : '0');
  if (options.align) {
    props.setAttribute('algn', options.align);
  } else if (options.flipAlign) {
    // A paragraph that states no alignment is left-aligned.
    const align = props.getAttribute('algn') ?? 'l';
    if (align === 'r') props.setAttribute('algn', 'l');
    else if (align === 'l') props.setAttribute('algn', 'r');
  }
  writeLanguage(paragraph, text, options, prefix);
}

/**
 * Writes the language of the paragraph's runs (see {@link filledLanguage}). PowerPoint places
 * punctuation and spaces by the language of their run, not only by the paragraph's direction:
 * in a run marked English, "contact: name@example.org" after Hebrew words puts the colon on the
 * far side of the address, and in a run marked Hebrew, English text has its figures moved. It
 * also checks spelling in the run's language. The end of the paragraph is given a left-to-right
 * language too when it states a right-to-left one, as text typed at its end takes that language.
 */
function writeLanguage(paragraph: Element, text: string, options: WritingOptions, prefix: string): void {
  const doc = paragraph.ownerDocument;
  if (!doc) return;
  for (const run of childElements(paragraph, NS.a, 'r')) {
    let runProps = childElement(run, NS.a, 'rPr');
    const stated = runProps?.getAttribute('lang') ?? '';
    const language = filledLanguage(text, stated, options);
    if (language === stated) continue;
    if (!runProps) {
      runProps = doc.createElementNS(NS.a, `${prefix}rPr`);
      run.insertBefore(runProps, run.firstChild);
    }
    runProps.setAttribute('lang', language);
  }
  const end = childElement(paragraph, NS.a, 'endParaRPr');
  const ended = end?.getAttribute('lang') ?? '';
  if (end && isRtlLanguage(ended) && !isRtlLanguage(filledLanguage(text, ended, options))) end.setAttribute('lang', leftToRightTag(options.language));
}
