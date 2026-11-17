import { DOMParser, XMLSerializer } from '@xmldom/xmldom';
import type { Document, Element, Node } from '@xmldom/xmldom';
import { writable } from '../util/chars.ts';

export type { Document, Element, Node };

export const NS = {
  a: 'http://schemas.openxmlformats.org/drawingml/2006/main',
  p: 'http://schemas.openxmlformats.org/presentationml/2006/main',
  r: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
  rels: 'http://schemas.openxmlformats.org/package/2006/relationships',
  contentTypes: 'http://schemas.openxmlformats.org/package/2006/content-types',
} as const;

export const XML_HEADER = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n';

export function parseXml(text: string): Document {
  return new DOMParser().parseFromString(text, 'text/xml');
}

/** A document as the text of a part. Characters that XML cannot hold are left out: a part with one could not be opened. */
export function serializeXml(doc: Document): string {
  return writable(new XMLSerializer().serializeToString(doc));
}

/** The text sizes PowerPoint opens a file with, in hundredths of a point: 1 to 4000 pt. One size outside them and the file does not open. */
const OPENABLE_SIZE = { least: 100, most: 400_000 };

/**
 * The text of a part with every text size inside the range PowerPoint opens. A size the renderer
 * writes is the template's size times the manifest's `textScale` times the deck's, and a deck's
 * factor alone cannot keep that product inside it: 300 on a template whose body text is 14 pt at
 * a factor of 1.3 is past 4000 pt. Such a size is set to the nearest that opens; every other is
 * left as written, so a part with none outside the range is the same text as before.
 */
export function openableSizes(xml: string): string {
  return xml.replace(/(<(?:[A-Za-z_][\w.-]*:)?(?:rPr|defRPr|endParaRPr)\b[^>]*?\ssz=")([^"]*)(")/g, (whole, head: string, value: string, tail: string) => {
    const size = Number(value);
    if (!Number.isFinite(size) || (size >= OPENABLE_SIZE.least && size <= OPENABLE_SIZE.most)) return whole;
    return `${head}${size < OPENABLE_SIZE.least ? OPENABLE_SIZE.least : OPENABLE_SIZE.most}${tail}`;
  });
}

export function rootElement(doc: Document): Element {
  const root = doc.documentElement;
  if (!root) throw new Error('XML document has no root element');
  return root;
}

const ELEMENT_NODE = 1;

export function isElement(node: Node | null | undefined): node is Element {
  return !!node && node.nodeType === ELEMENT_NODE;
}

/** Direct child elements, optionally filtered by namespace and local name. */
export function childElements(parent: Node, ns?: string, local?: string): Element[] {
  const out: Element[] = [];
  for (let node = parent.firstChild; node; node = node.nextSibling) {
    if (!isElement(node)) continue;
    if (ns !== undefined && node.namespaceURI !== ns) continue;
    if (local !== undefined && node.localName !== local) continue;
    out.push(node);
  }
  return out;
}

export function childElement(parent: Node, ns: string, local: string): Element | undefined {
  return childElements(parent, ns, local)[0];
}

/** Follows a chain of child element names, e.g. `path(sp, [NS.p, 'spPr'], [NS.a, 'xfrm'])`. */
export function path(parent: Node, ...steps: Array<readonly [string, string]>): Element | undefined {
  let current: Node | undefined = parent;
  for (const [ns, local] of steps) {
    if (!current) return undefined;
    current = childElement(current, ns, local);
  }
  return current as Element | undefined;
}

/** All descendant elements with the given name, as a static array (safe to mutate the tree while iterating). */
export function descendants(parent: Element | Document, ns: string, local: string): Element[] {
  return Array.from(parent.getElementsByTagNameNS(ns, local));
}

export function removeNode(node: Node): void {
  node.parentNode?.removeChild(node);
}

export function escapeXmlAttr(value: string): string {
  return writable(value)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;');
}
