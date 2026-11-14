import { NS, childElement, childElements, removeNode } from '../pptx/xml.ts';
import type { Element } from '../pptx/xml.ts';
import type { Direction, TableTextStyle } from '../template/manifest.ts';

export interface TextOptions {
  /** Swap left and right alignment (the deck reads the other way from the template). */
  flipAlign: boolean;
  /** Which run of the template paragraph supplies the formatting. Defaults to the first. */
  styleRun?: 'first' | 'last';
  /** Set the paragraph alignment outright, whatever the template has. */
  align?: 'l' | 'r' | 'ctr';
  /** Formatting to put on the text, for places where the template defines none. */
  style?: TableTextStyle;
}

const RTL_LETTER = /[\p{Script=Hebrew}\p{Script=Arabic}\p{Script=Syriac}\p{Script=Thaana}]/u;

/** Base direction of a piece of text from its first letter; `undefined` when it has none (e.g. "78%"). */
export function detectDirection(text: string): Direction | undefined {
  const firstLetter = /\p{L}/u.exec(text)?.[0];
  if (!firstLetter) return undefined;
  return RTL_LETTER.test(firstLetter) ? 'rtl' : 'ltr';
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
  if (options.style) applyStyle(run, options.style, prefix);

  // Drop every run, line break and field of the template paragraph, then put ours back.
  for (const child of childElements(paragraph)) {
    if (child.localName !== 'pPr' && child.localName !== 'endParaRPr') removeNode(child);
  }
  paragraph.insertBefore(run, endProps ?? null);

  applyDirection(paragraph, text, options, prefix);
  return paragraph;
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
  const direction = detectDirection(text);
  let props = childElement(paragraph, NS.a, 'pPr');
  if (!props) {
    if (!direction && !options.flipAlign && !options.align) return;
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
  if (direction === 'rtl') markRightToLeftLanguage(paragraph, text, prefix);
}

const RTL_LANGUAGE = /^(he|iw|yi|ar|fa|ur|ps|ug|sd|ckb|syr|dv)\b/i;

/**
 * PowerPoint places punctuation and spaces by the language of their run, not only by the
 * paragraph's direction: in a run marked English, "contact: name@example.org" after Hebrew words
 * puts the colon on the far side of the address. A right-to-left text therefore gets a
 * right-to-left language on its run, unless the template already gave it one.
 */
function markRightToLeftLanguage(paragraph: Element, text: string, prefix: string): void {
  const doc = paragraph.ownerDocument;
  if (!doc) return;
  const letter = RTL_LETTER.exec(text)?.[0] ?? '';
  const language = /\p{Script=Hebrew}/u.test(letter) ? 'he-IL' : /\p{Script=Arabic}/u.test(letter) ? 'ar-SA' : /\p{Script=Syriac}/u.test(letter) ? 'syr-SY' : 'dv-MV';
  for (const run of childElements(paragraph, NS.a, 'r')) {
    let runProps = childElement(run, NS.a, 'rPr');
    if (!runProps) {
      runProps = doc.createElementNS(NS.a, `${prefix}rPr`);
      run.insertBefore(runProps, run.firstChild);
    }
    if (!RTL_LANGUAGE.test(runProps.getAttribute('lang') ?? '')) runProps.setAttribute('lang', language);
  }
}
