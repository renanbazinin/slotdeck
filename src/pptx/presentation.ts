import { isRelType, readRels, resolveTarget } from './package.ts';
import type { PptxPackage } from './package.ts';
import { NS, childElement, childElements, parseXml, path, rootElement } from './xml.ts';

export interface SlideRef {
  /** `p:sldId/@id` */
  id: string;
  relId: string;
  /** Package path, e.g. `ppt/slides/slide3.xml` */
  part: string;
}

export interface PresentationInfo {
  /** Package path of the main presentation part. */
  part: string;
  /** Slides in presentation order. */
  slides: SlideRef[];
  /** Slide size in EMU. */
  slideWidth: number;
  slideHeight: number;
  /** Size in points of text that sets no size of its own, when the presentation states one. */
  defaultFontSize?: number;
}

export function readPresentation(pkg: PptxPackage): PresentationInfo {
  const main = readRels(pkg, '').find((rel) => isRelType(rel, 'officeDocument'));
  if (!main) throw new Error('Package has no main document relationship');
  const part = resolveTarget('', main.target);

  const root = rootElement(parseXml(pkg.readText(part)));
  const targets = new Map(readRels(pkg, part).map((rel) => [rel.id, resolveTarget(part, rel.target)]));

  const list = childElement(root, NS.p, 'sldIdLst');
  const slides: SlideRef[] = [];
  for (const el of list ? childElements(list, NS.p, 'sldId') : []) {
    const relId = el.getAttributeNS(NS.r, 'id') ?? '';
    const target = targets.get(relId);
    if (!target) throw new Error(`Slide relationship ${relId} is missing from ${part}`);
    slides.push({ id: el.getAttribute('id') ?? '', relId, part: target });
  }

  const size = childElement(root, NS.p, 'sldSz');
  const info: PresentationInfo = {
    part,
    slides,
    slideWidth: Number(size?.getAttribute('cx') ?? 0),
    slideHeight: Number(size?.getAttribute('cy') ?? 0),
  };
  const defaultSize = Number(path(root, [NS.p, 'defaultTextStyle'], [NS.a, 'lvl1pPr'], [NS.a, 'defRPr'])?.getAttribute('sz') ?? 0);
  if (defaultSize > 0) info.defaultFontSize = defaultSize / 100;
  return info;
}
