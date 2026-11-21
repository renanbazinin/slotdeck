import { NS, childElement, childElements, descendants, escapeXmlAttr, isElement, path, removeNode } from '../pptx/xml.ts';
import type { Element, Node } from '../pptx/xml.ts';

export interface PictureFill {
  /** Relationship id of the new image on the slide. */
  relId: string;
  /** Pixel size of the new image. */
  width: number;
  height: number;
  /** `cover` crops the image to the frame; `contain` shrinks the frame to the image. */
  fit: 'cover' | 'contain';
  alt?: string;
  /**
   * The picture is decoration, which a screen reader passes over, where it has no alternative text:
   * a pictogram beside a heading says nothing the heading does not. Office's own mark for it.
   */
  decorative?: boolean;
  /** Area the picture may use, in slide units (EMU), replacing the sample picture's own frame. */
  area?: { x: number; y: number; cx: number; cy: number };
}

/** Swaps the image of a `p:pic`, keeping its frame, geometry (e.g. rounded corners) and effects. */
export function setPicture(pic: Element, fill: PictureFill): void {
  if (pic.namespaceURI !== NS.p || pic.localName !== 'pic') {
    throw new Error('image slots must point at a picture shape');
  }
  const blipFill = childElement(pic, NS.p, 'blipFill');
  const blip = blipFill && childElement(blipFill, NS.a, 'blip');
  const xfrm = path(pic, [NS.p, 'spPr'], [NS.a, 'xfrm']);
  const off = xfrm && childElement(xfrm, NS.a, 'off');
  const ext = xfrm && childElement(xfrm, NS.a, 'ext');
  if (!blipFill || !blip || !off || !ext) throw new Error('picture shape is missing its image or its frame');
  const doc = pic.ownerDocument;
  if (!doc) throw new Error('picture is detached from its document');

  const relPrefix = pic.lookupPrefix(NS.r) ?? 'r';
  blip.setAttributeNS(NS.r, `${relPrefix}:embed`, fill.relId);
  blip.removeAttributeNS(NS.r, 'link');

  // An SVG variant stored next to the old image would win over the new bitmap.
  for (const extension of descendants(blip, NS.a, 'ext')) {
    if (childElements(extension).some((child) => child.localName === 'svgBlip')) removeNode(extension);
  }
  const extensionList = childElement(blip, NS.a, 'extLst');
  if (extensionList && childElements(extensionList).length === 0) removeNode(extensionList);

  if (fill.area) {
    const parent = pic.parentNode as Element | null;
    if (!parent || parent.localName !== 'spTree') throw new Error('a picture area can only be set for a picture that is not inside a group');
    off.setAttribute('x', String(Math.round(fill.area.x)));
    off.setAttribute('y', String(Math.round(fill.area.y)));
    ext.setAttribute('cx', String(Math.round(fill.area.cx)));
    ext.setAttribute('cy', String(Math.round(fill.area.cy)));
  }

  const frame = {
    x: Number(off.getAttribute('x') ?? 0),
    y: Number(off.getAttribute('y') ?? 0),
    cx: Number(ext.getAttribute('cx') ?? 0),
    cy: Number(ext.getAttribute('cy') ?? 0),
  };
  const scale = ancestorScale(pic);
  const frameAspect = (frame.cx * scale.sx) / (frame.cy * scale.sy);
  const imageAspect = fill.width / fill.height;
  const usable = Number.isFinite(frameAspect) && frameAspect > 0 && Number.isFinite(imageAspect) && imageAspect > 0;

  const oldCrop = childElement(blipFill, NS.a, 'srcRect');
  if (oldCrop) removeNode(oldCrop);

  if (usable && fill.fit === 'cover') {
    // Crop the image symmetrically until its aspect ratio matches the frame.
    const crop = doc.createElementNS(NS.a, `${blip.prefix ? `${blip.prefix}:` : ''}srcRect`);
    if (imageAspect > frameAspect) {
      const side = Math.round(((1 - frameAspect / imageAspect) / 2) * 100000);
      if (side > 0) {
        crop.setAttribute('l', String(side));
        crop.setAttribute('r', String(side));
      }
    } else {
      const side = Math.round(((1 - imageAspect / frameAspect) / 2) * 100000);
      if (side > 0) {
        crop.setAttribute('t', String(side));
        crop.setAttribute('b', String(side));
      }
    }
    if (crop.attributes.length > 0) blipFill.insertBefore(crop, blip.nextSibling);
  } else if (usable) {
    // Shrink the frame around its centre so the whole image shows undistorted.
    if (imageAspect > frameAspect) {
      const cy = Math.round((frame.cx * scale.sx) / imageAspect / scale.sy);
      off.setAttribute('y', String(Math.round(frame.y + (frame.cy - cy) / 2)));
      ext.setAttribute('cy', String(cy));
    } else {
      const cx = Math.round((frame.cy * scale.sy * imageAspect) / scale.sx);
      off.setAttribute('x', String(Math.round(frame.x + (frame.cx - cx) / 2)));
      ext.setAttribute('cx', String(cx));
    }
  }

  if (!childElement(blipFill, NS.a, 'stretch') && !childElement(blipFill, NS.a, 'tile')) {
    const prefix = blip.prefix ? `${blip.prefix}:` : '';
    const stretch = doc.createElementNS(NS.a, `${prefix}stretch`);
    stretch.appendChild(doc.createElementNS(NS.a, `${prefix}fillRect`));
    blipFill.appendChild(stretch);
  }

  const props = path(pic, [NS.p, 'nvPicPr'], [NS.p, 'cNvPr']);
  if (props) {
    props.removeAttribute('title');
    props.setAttribute('descr', fill.alt ?? '');
    setDecorative(props, fill.decorative === true && !fill.alt);
  }
}

/** The extension by which Office marks a picture as decoration (Alt Text > Mark as decorative). */
const DECORATIVE_URI = '{C183D7F6-B498-43B3-948B-1728B52AA6E4}';
const NS_DECORATIVE = 'http://schemas.microsoft.com/office/drawing/2017/decorative';

/** Marks a shape's properties as decoration, or takes the mark away: a picture with words to say is no decoration. */
function setDecorative(props: Element, decorative: boolean): void {
  const doc = props.ownerDocument;
  const prefix = props.lookupPrefix(NS.a) ?? 'a';
  let list: Element | undefined = childElement(props, NS.a, 'extLst') ?? undefined;
  for (const extension of list ? childElements(list, NS.a, 'ext') : []) if (extension.getAttribute('uri') === DECORATIVE_URI) removeNode(extension);
  if (list && childElements(list).length === 0) {
    removeNode(list);
    list = undefined;
  }
  if (!decorative || !doc) return;
  if (!list) {
    list = doc.createElementNS(NS.a, `${prefix}:extLst`);
    props.appendChild(list);
  }
  const extension = doc.createElementNS(NS.a, `${prefix}:ext`);
  extension.setAttribute('uri', DECORATIVE_URI);
  const mark = doc.createElementNS(NS_DECORATIVE, 'adec:decorative');
  mark.setAttribute('val', '1');
  extension.appendChild(mark);
  list.appendChild(extension);
}

/**
 * An empty picture frame at an area of the slide (EMU), for a slide that the software lays out
 * itself: {@link setPicture} then puts a picture in it. It names no image yet.
 */
export function pictureFrameXml(shapeId: string, name: string, area: { x: number; y: number; cx: number; cy: number }): string {
  return (
    `<p:pic xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">` +
    `<p:nvPicPr><p:cNvPr id="${escapeXmlAttr(shapeId)}" name="${escapeXmlAttr(name)}" descr=""/><p:cNvPicPr><a:picLocks noChangeAspect="1"/></p:cNvPicPr><p:nvPr/></p:nvPicPr>` +
    `<p:blipFill><a:blip r:embed=""/><a:stretch><a:fillRect/></a:stretch></p:blipFill>` +
    `<p:spPr><a:xfrm><a:off x="${Math.round(area.x)}" y="${Math.round(area.y)}"/><a:ext cx="${Math.round(area.cx)}" cy="${Math.round(area.cy)}"/></a:xfrm>` +
    `<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr></p:pic>`
  );
}

/** Combined scale of every group the element sits in (child units to slide units). */
function ancestorScale(element: Element): { sx: number; sy: number } {
  let sx = 1;
  let sy = 1;
  for (let node: Node | null = element.parentNode; node && isElement(node); node = node.parentNode) {
    if (node.namespaceURI !== NS.p || node.localName !== 'grpSp') continue;
    const xfrm = path(node, [NS.p, 'grpSpPr'], [NS.a, 'xfrm']);
    const ext = xfrm && childElement(xfrm, NS.a, 'ext');
    const childExt = xfrm && childElement(xfrm, NS.a, 'chExt');
    const cx = Number(childExt?.getAttribute('cx') ?? 0);
    const cy = Number(childExt?.getAttribute('cy') ?? 0);
    if (cx > 0) sx *= Number(ext?.getAttribute('cx') ?? cx) / cx;
    if (cy > 0) sy *= Number(ext?.getAttribute('cy') ?? cy) / cy;
  }
  return { sx, sy };
}
