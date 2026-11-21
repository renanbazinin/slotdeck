// Makes the shapes of a free slide on a copy of the base slide: everything that is not frame is
// taken off, and the planned shapes are put on. Text shapes, panels and rules are copies of the
// template's own shapes, so the typeface, the colours and the corners are the template's; a
// picture frame and the place of a chart are made here. The renderer then fills them as it fills
// the slots of any slide.
import { NS, childElement, childElements, descendants, parseXml, path, removeNode, rootElement } from '../pptx/xml.ts';
import type { Document, Element } from '../pptx/xml.ts';
import { pictureFrameXml } from '../render/picture.ts';
import { keepCorners } from '../render/repeat.ts';
import { tableFrameXml } from '../render/table.ts';
import { setShapeText } from '../render/text.ts';
import type { WritingOptions } from '../render/text.ts';
import type { SlideBox } from '../template/manifest.ts';
import type { FreeKit } from './kit.ts';
import type { FreePlan, MadeShape } from './plan.ts';

/** Finds a shape of a template slide by its id: what a made shape is copied from. */
export type ShapeSource = (slide: number, shapeId: string) => Element | undefined;

const P = 'http://schemas.openxmlformats.org/presentationml/2006/main';
const A = 'http://schemas.openxmlformats.org/drawingml/2006/main';

/** Takes everything off a copy of the base slide that is not its frame, and puts the planned shapes on. */
export function buildFreeSlide(doc: Document, root: Element, kit: FreeKit, plan: FreePlan, source: ShapeSource, writing: WritingOptions): void {
  const rtl = writing.direction === 'rtl';
  const tree = path(root, [NS.p, 'cSld'], [NS.p, 'spTree']);
  if (!tree) throw new Error('the base slide of a free slide has no shapes to stand on');
  const stays = new Set(kit.stays);
  const idsOf = (element: Element) => descendants(element, NS.p, 'cNvPr').map((props) => props.getAttribute('id') ?? '');
  for (const child of childElements(tree)) {
    if (child.namespaceURI === NS.p && (child.localName === 'nvGrpSpPr' || child.localName === 'grpSpPr')) continue;
    if (!idsOf(child).some((id) => stays.has(id))) removeNode(child);
  }
  // Animations name shapes that are gone.
  for (const timing of childElements(root, NS.p, 'timing')) removeNode(timing);

  const area = (box: SlideBox) => ({ x: box.x * kit.slideWidth, y: box.y * kit.slideHeight, cx: box.w * kit.slideWidth, cy: box.h * kit.slideHeight });
  for (const shape of plan.shapes) {
    const at = area(shape.box);
    let made: Element | undefined;
    if (shape.kind === 'picture') {
      made = doc.importNode(rootElement(parseXml(pictureFrameXml(shape.id, shape.name, at))), true) as Element;
      const outline = shape.outline && source(shape.outline.slide, shape.outline.shapeId);
      const wanted = outline && path(outline, [NS.p, 'spPr'], [NS.a, 'prstGeom']);
      const own = path(made, [NS.p, 'spPr'], [NS.a, 'prstGeom']);
      // A photo takes the outline the template gives its photos: its rounded corners.
      if (wanted && own?.parentNode) own.parentNode.replaceChild(doc.importNode(wanted, true), own);
    } else if (shape.kind === 'anchor') {
      made = doc.importNode(rootElement(parseXml(anchorXml(shape, at))), true) as Element;
    } else if (shape.kind === 'key') {
      made = doc.importNode(rootElement(parseXml(keyXml(shape, at))), true) as Element;
    } else if (shape.kind === 'table') {
      const sample = shape.from && source(shape.from.slide, shape.from.shapeId);
      if (sample && sample.localName === 'graphicFrame') {
        made = doc.importNode(sample, true) as Element;
        rename(made, shape);
        place(made, at);
      } else {
        const accent = kit.palette.accent ?? kit.palette.ink ?? '404040';
        made = doc.importNode(rootElement(parseXml(tableFrameXml(shape.id, at, { header: accent, ...(shape.look ?? { band: 'F2F2F2', line: 'D9D9D9' }) }))), true) as Element;
        rename(made, shape);
      }
    } else {
      const sample = shape.from && source(shape.from.slide, shape.from.shapeId);
      if (!sample) throw new Error(`a free slide copies shape #${shape.from?.shapeId} of template slide ${shape.from?.slide}, which is not there`);
      made = doc.importNode(sample, true) as Element;
      rename(made, shape);
      if (shape.kind === 'text') {
        place(made, at);
        prepareText(made, shape, kit, writing);
      } else if (shape.kind === 'rule') {
        const extent = path(made, [NS.p, 'spPr'], [NS.a, 'xfrm'], [NS.a, 'ext']);
        const thick = Math.min(Number(extent?.getAttribute('cx') ?? 0), Number(extent?.getAttribute('cy') ?? 0));
        place(made, { x: at.x, y: at.y - thick / 2, cx: at.cx, cy: thick });
      } else {
        // A panel keeps the corners it was drawn with, whatever its new size, up to an eighth of its shorter side.
        const drawn = kit.panel && shape.from?.shapeId === kit.panel.shapeId ? kit.panel.drawn : kit.notePanel?.drawn;
        const was = drawn ? Math.min(drawn.w * kit.slideWidth, drawn.h * kit.slideHeight) : 0;
        place(made, at);
        const now = Math.min(at.cx, at.cy);
        if (was > 0 && now > 0) keepCorners(made, was, now, Number.POSITIVE_INFINITY);
        capCorners(made);
        const body = childElement(made, NS.p, 'txBody');
        // A panel that carried a sample text of its own is drawn empty.
        if (body && descendants(body, NS.a, 't').some((node) => (node.textContent ?? '').trim() !== '')) setShapeText(made, [''], { flipAlign: false });
      }
    }
    tree.appendChild(made);
  }
}

function anchorXml(shape: MadeShape, at: { x: number; y: number; cx: number; cy: number }): string {
  return (
    `<p:sp xmlns:p="${P}" xmlns:a="${A}"><p:nvSpPr><p:cNvPr id="${shape.id}" name="${shape.name}"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>` +
    `<p:spPr><a:xfrm><a:off x="${Math.round(at.x)}" y="${Math.round(at.y)}"/><a:ext cx="${Math.round(at.cx)}" cy="${Math.round(at.cy)}"/></a:xfrm>` +
    `<a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/></p:spPr></p:sp>`
  );
}

/** The key of a legend: a small square filled in the colour of its series, with no outline. */
function keyXml(shape: MadeShape, at: { x: number; y: number; cx: number; cy: number }): string {
  const color = (shape.color ?? '808080').replace(/[^0-9A-Fa-f]/g, '').toUpperCase().padStart(6, '0').slice(0, 6);
  return (
    `<p:sp xmlns:p="${P}" xmlns:a="${A}"><p:nvSpPr><p:cNvPr id="${shape.id}" name="${shape.name}"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>` +
    `<p:spPr><a:xfrm><a:off x="${Math.round(at.x)}" y="${Math.round(at.y)}"/><a:ext cx="${Math.round(at.cx)}" cy="${Math.round(at.cy)}"/></a:xfrm>` +
    `<a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:solidFill><a:srgbClr val="${color}"/></a:solidFill><a:ln><a:noFill/></a:ln></p:spPr></p:sp>`
  );
}

/** A copy is a shape of its own: a new id and name, and nothing that named the shape it was copied from. */
function rename(copy: Element, shape: MadeShape): void {
  const props = descendants(copy, NS.p, 'cNvPr')[0];
  if (props) {
    props.setAttribute('id', shape.id);
    props.setAttribute('name', shape.name);
    props.removeAttribute('descr');
    props.removeAttribute('title');
    for (const list of childElements(props, NS.a, 'extLst')) removeNode(list);
  }
  // A line that was glued to two shapes of its old slide is glued to nothing here.
  for (const glue of [...descendants(copy, NS.a, 'stCxn'), ...descendants(copy, NS.a, 'endCxn')]) removeNode(glue);
}

/** Puts a copy at its place, upright: it is not turned or mirrored as its source may have been. */
function place(copy: Element, at: { x: number; y: number; cx: number; cy: number }): void {
  const doc = copy.ownerDocument;
  if (!doc) return;
  let xfrm: Element | undefined;
  if (copy.localName === 'graphicFrame') {
    xfrm = childElement(copy, NS.p, 'xfrm');
  } else {
    const props = childElement(copy, NS.p, 'spPr');
    if (!props) return;
    xfrm = childElement(props, NS.a, 'xfrm');
    if (!xfrm) {
      xfrm = doc.createElementNS(NS.a, 'a:xfrm');
      props.insertBefore(xfrm, props.firstChild);
    }
  }
  if (!xfrm) return;
  for (const name of ['rot', 'flipH', 'flipV']) xfrm.removeAttribute(name);
  const child = (name: string): Element => {
    let element = childElement(xfrm as Element, NS.a, name);
    if (!element) {
      element = doc.createElementNS(NS.a, `a:${name}`);
      (xfrm as Element).appendChild(element);
    }
    return element;
  };
  const off = child('off');
  const ext = child('ext');
  off.setAttribute('x', String(Math.round(at.x)));
  off.setAttribute('y', String(Math.round(at.y)));
  ext.setAttribute('cx', String(Math.max(0, Math.round(at.cx))));
  ext.setAttribute('cy', String(Math.max(0, Math.round(at.cy))));
}

/** A rounded corner is at most an eighth of the shape's shorter side: a low tile must not become a pill. */
function capCorners(copy: Element): void {
  const geometry = path(copy, [NS.p, 'spPr'], [NS.a, 'prstGeom']);
  if (!geometry || !/^round\w*Rect$/.test(geometry.getAttribute('prst') ?? '')) return;
  for (const guide of descendants(geometry, NS.a, 'gd')) {
    const value = /^val (-?\d+)$/.exec(guide.getAttribute('fmla') ?? '');
    if (value && Number(value[1]) > 12500) guide.setAttribute('fmla', 'val 12500');
  }
}

/**
 * Makes a copied text shape ready to be filled: one sample paragraph that states its size, its
 * weight, its alignment and its direction; a box that keeps the size it was given and does not
 * shrink or grow with its text; the insets the plan counted with.
 */
function prepareText(copy: Element, shape: MadeShape, kit: FreeKit, writing: WritingOptions): void {
  const rtl = writing.direction === 'rtl';
  const doc = copy.ownerDocument;
  const body = childElement(copy, NS.p, 'txBody');
  if (!doc || !body) throw new Error(`a free slide copies shape #${shape.from?.shapeId} of template slide ${shape.from?.slide} as a text, and it holds none`);
  const size = shape.size ?? 12;
  const paragraphs = childElements(body, NS.a, 'p');
  const first = paragraphs[0];
  if (!first) throw new Error('a text shape of the template has no paragraph to take its look from');
  for (const paragraph of paragraphs.slice(1)) removeNode(paragraph);
  const prefix = first.prefix ? `${first.prefix}:` : 'a:';
  const create = (name: string) => doc.createElementNS(NS.a, `${prefix}${name}`);

  // Run properties: on the first run, which gives every new paragraph its look, and at the paragraph's end.
  let end = childElement(first, NS.a, 'endParaRPr');
  let run = childElements(first, NS.a, 'r')[0];
  if (!run) {
    run = create('r');
    const props = create('rPr');
    if (end) {
      for (let i = 0; i < end.attributes.length; i++) {
        const attr = end.attributes.item(i);
        if (attr) props.setAttribute(attr.name, attr.value);
      }
      for (const child of childElements(end)) props.appendChild(child.cloneNode(true));
    }
    run.appendChild(props);
    run.appendChild(create('t'));
    first.insertBefore(run, end ?? null);
  }
  let props = childElement(run, NS.a, 'rPr');
  if (!props) {
    props = create('rPr');
    run.insertBefore(props, run.firstChild);
  }
  if (!end) {
    end = create('endParaRPr');
    first.appendChild(end);
  }
  for (const target of [props, end]) {
    target.setAttribute('sz', String(Math.round(size * 100)));
    if (shape.bold !== undefined) target.setAttribute('b', shape.bold ? '1' : '0');
    if (shape.italic !== undefined) target.setAttribute('i', shape.italic ? '1' : '0');
  }
  if (shape.color) {
    for (const name of ['noFill', 'solidFill', 'gradFill', 'blipFill', 'pattFill', 'grpFill']) {
      const old = childElement(props, NS.a, name);
      if (old) removeNode(old);
    }
    const fill = create('solidFill');
    const color = create('srgbClr');
    color.setAttribute('val', shape.color);
    fill.appendChild(color);
    const line = childElement(props, NS.a, 'ln');
    props.insertBefore(fill, line ? line.nextSibling : props.firstChild);
  }

  // Paragraph properties: the deck's direction, an alignment that starts where reading starts, and bullets set for the size.
  let para = childElement(first, NS.a, 'pPr');
  if (!para) {
    para = create('pPr');
    first.insertBefore(para, first.firstChild);
  }
  para.setAttribute('rtl', rtl ? '1' : '0');
  para.setAttribute('algn', shape.align === 'ctr' ? 'ctr' : rtl ? 'r' : 'l');
  // Space before and after a paragraph is the plan's to give (a slot's `paragraphGap`). The copy
  // states none, whatever its source and the list style it inherits say: PowerPoint draws the
  // space before a first paragraph too, and the plan counted neither. The line spacing stays the
  // source's: the plan counts its lines at it.
  for (const name of ['spcBef', 'spcAft']) {
    const old = childElement(para, NS.a, name);
    if (old) removeNode(old);
  }
  let lineSpacing = childElement(para, NS.a, 'lnSpc');
  if (shape.lineSpacing) {
    // Lines the plan closed up (`MadeShape.lineSpacing`) are drawn as it counted them, whatever the
    // source states or inherits. Line spacing comes first among a paragraph's properties.
    if (lineSpacing) removeNode(lineSpacing);
    lineSpacing = create('lnSpc');
    const drawn = shape.lineSpacing;
    const value = create('points' in drawn ? 'spcPts' : 'spcPct');
    value.setAttribute('val', String('points' in drawn ? Math.round(drawn.points * 100) : Math.round(drawn.percent * 100000)));
    lineSpacing.appendChild(value);
    para.insertBefore(lineSpacing, para.firstChild);
  } else if (lineSpacing && shape.lines === 'several') {
    // A text that can take several lines has them at least as far apart as its size asks. A
    // template may set the one line of a heading closer than that (70% is in use), which lifts a
    // single line a little and does no harm; two lines set so run into each other. A copy that
    // holds one line by its making keeps the spacing of its source, and with it its place.
    const percent = childElement(lineSpacing, NS.a, 'spcPct');
    const points = childElement(lineSpacing, NS.a, 'spcPts');
    if (percent && Number(percent.getAttribute('val') ?? 100000) < 100000) percent.setAttribute('val', '100000');
    else if (points && Number(points.getAttribute('val') ?? 0) < 120 * size) points.setAttribute('val', String(Math.round(120 * size)));
  }
  const afterSpacing = lineSpacing ? lineSpacing.nextSibling : para.firstChild;
  for (const name of ['spcBef', 'spcAft']) {
    const space = create(name);
    const none = create('spcPts');
    none.setAttribute('val', '0');
    space.appendChild(none);
    para.insertBefore(space, afterSpacing);
  }
  const BULLETS = ['buNone', 'buAutoNum', 'buChar', 'buBlip'];
  const hang = Math.round(1.1 * size * 12700);
  if (shape.bullets === 'add') {
    for (const name of BULLETS) {
      const old = childElement(para, NS.a, name);
      if (old) removeNode(old);
    }
    // Paragraph properties keep an order: spacing, bullet colour, bullet size, bullet font, the bullet, then the rest.
    const after = ['tabLst', 'defRPr', 'extLst'].map((name) => childElement(para as Element, NS.a, name)).find(Boolean) ?? null;
    if (!childElement(para, NS.a, 'buFont') && !childElement(para, NS.a, 'buFontTx')) {
      const font = create('buFont');
      font.setAttribute('typeface', 'Arial');
      para.insertBefore(font, after);
    }
    const bullet = create('buChar');
    bullet.setAttribute('char', '•');
    para.insertBefore(bullet, after);
    para.setAttribute('marL', String(hang));
    para.setAttribute('indent', String(-hang));
  } else if (shape.bullets === 'none') {
    // A stand-in for running text that was taken from a bulleted list: the copy has no bullet
    // and no hanging indent, whatever the source and the list style it inherits say.
    for (const name of BULLETS) {
      const old = childElement(para, NS.a, name);
      if (old) removeNode(old);
    }
    const after = ['tabLst', 'defRPr', 'extLst'].map((name) => childElement(para as Element, NS.a, name)).find(Boolean) ?? null;
    para.insertBefore(create('buNone'), after);
    para.setAttribute('marL', '0');
    para.setAttribute('indent', '0');
  } else if (shape.bullets === 'own') {
    // The source's bullet hung as far as its own size asked; at the new size it hangs to match.
    para.setAttribute('marL', String(hang));
    para.setAttribute('indent', String(-hang));
  }

  // The box: it wraps or it does not, keeps its size, and has the insets the plan counted with.
  let box = childElement(body, NS.a, 'bodyPr');
  if (!box) {
    box = create('bodyPr');
    body.insertBefore(box, body.firstChild);
  }
  const inset = String(Math.round((shape.inset ?? 0) * 12700));
  box.setAttribute('wrap', shape.wrap === false ? 'none' : 'square');
  box.setAttribute('lIns', inset);
  box.setAttribute('rIns', inset);
  box.setAttribute('tIns', '0');
  box.setAttribute('bIns', '0');
  box.setAttribute('anchor', shape.anchor === 'ctr' ? 'ctr' : shape.anchor === 'b' ? 'b' : 't');
  // Nor does it draw space before its first paragraph or after its last, which the plan does not count.
  for (const name of ['vert', 'numCol', 'spcCol', 'rtlCol', 'rot', 'spcFirstLastPara']) box.removeAttribute(name);
  if (shape.columns && shape.columns > 1) {
    box.setAttribute('numCol', String(shape.columns));
    box.setAttribute('spcCol', String(Math.round(kit.gap * kit.slideWidth)));
    if (rtl) box.setAttribute('rtlCol', '1');
  }
  for (const name of ['noAutofit', 'normAutofit', 'spAutoFit']) {
    const old = childElement(box, NS.a, name);
    if (old) removeNode(old);
  }
  const warp = childElement(box, NS.a, 'prstTxWarp');
  box.insertBefore(create('noAutofit'), warp ? warp.nextSibling : box.firstChild);

  // A label of the slide (a legend's series name) is written as the deck's own text is.
  if (shape.text !== undefined) setShapeText(copy, [shape.text], { flipAlign: false, ...writing });
}
