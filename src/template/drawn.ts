// What a slide layout and a slide master draw of their own: the shapes behind every slide that
// stands on them, which no slide can move or take off. A decoration down one side, a banner
// under the title, a strip along the foot.
import { isRelType, readRels, resolveTarget } from '../pptx/package.ts';
import type { PptxPackage } from '../pptx/package.ts';
import { NS, childElement, childElements, parseXml, path, rootElement } from '../pptx/xml.ts';
import type { Element } from '../pptx/xml.ts';
import type { Box } from './inspect.ts';

/** A place in EMU, and how the space of a group's children maps onto the group's own place. */
type Place = { x: number; y: number; cx: number; cy: number };
type Scale = (place: Place) => Place;

const DRAWN = ['sp', 'pic', 'grpSp', 'cxnSp', 'graphicFrame'];

/** The shapes of a tree that are no placeholders, each at its place on the slide: the shapes of a group, not the group. */
function drawnIn(tree: Element, onSlide: Scale, out: Place[]): void {
  for (const child of childElements(tree)) {
    if (!DRAWN.includes(child.localName ?? '')) continue;
    const nonVisual = childElements(child).find((entry) => /^nv\w+Pr$/.test(entry.localName ?? ''));
    const application = nonVisual && childElement(nonVisual, NS.p, 'nvPr');
    if (application && childElement(application, NS.p, 'ph')) continue;
    const group = child.localName === 'grpSp';
    const properties = childElement(child, NS.p, group ? 'grpSpPr' : 'spPr');
    const transform = properties ? childElement(properties, NS.a, 'xfrm') : child.localName === 'graphicFrame' ? childElement(child, NS.p, 'xfrm') : undefined;
    const read = (name: string, first: string, second: string): [number, number] | undefined => {
      const element = transform && childElement(transform, NS.a, name);
      return element ? [Number(element.getAttribute(first) ?? 0), Number(element.getAttribute(second) ?? 0)] : undefined;
    };
    const off = read('off', 'x', 'y');
    const ext = read('ext', 'cx', 'cy');
    if (!off || !ext) continue;
    if (!group) {
      out.push(onSlide({ x: off[0], y: off[1], cx: ext[0], cy: ext[1] }));
      continue;
    }
    // A group draws its children in a space of their own, mapped onto the place the group has. One
    // group may span the slide and hold a decoration at each side: its children say where it draws.
    const inner = read('chOff', 'x', 'y') ?? off;
    const size = read('chExt', 'cx', 'cy') ?? ext;
    const sx = size[0] > 0 ? ext[0] / size[0] : 1;
    const sy = size[1] > 0 ? ext[1] / size[1] : 1;
    drawnIn(child, (place) => onSlide({ x: off[0] + (place.x - inner[0]) * sx, y: off[1] + (place.y - inner[1]) * sy, cx: place.cx * sx, cy: place.cy * sy }), out);
  }
}

/**
 * Where a slide layout, or a master, draws something of its own: the boxes of its pictures and
 * shapes that are no placeholders, those inside groups among them, as fractions of the slide.
 * `root` is the part's root element.
 */
export function drawnBoxes(root: Element, slideWidth: number, slideHeight: number): Box[] {
  const tree = path(root, [NS.p, 'cSld'], [NS.p, 'spTree']);
  if (!tree) return [];
  const places: Place[] = [];
  drawnIn(tree, (place) => place, places);
  return places.map((place) => ({ x: place.x / slideWidth, y: place.y / slideHeight, w: place.cx / slideWidth, h: place.cy / slideHeight })).filter((box) => box.w > 0 && box.h > 0);
}

/**
 * What stands behind a slide whatever the slide holds: what its layout draws and, where the
 * layout shows the master's shapes, what its master draws. As fractions of the slide.
 */
export function drawnBehind(pkg: PptxPackage, slidePart: string, slideWidth: number, slideHeight: number): Box[] {
  const related = (from: string, type: string): string | undefined => {
    if (!pkg.has(from)) return undefined;
    const rel = readRels(pkg, from).find((candidate) => isRelType(candidate, type));
    const target = rel ? resolveTarget(from, rel.target) : undefined;
    return target && pkg.has(target) ? target : undefined;
  };
  const layout = related(slidePart, 'slideLayout');
  if (!layout) return [];
  const layoutRoot = rootElement(parseXml(pkg.readText(layout)));
  const boxes = drawnBoxes(layoutRoot, slideWidth, slideHeight);
  const master = layoutRoot.getAttribute('showMasterSp') === '0' ? undefined : related(layout, 'slideMaster');
  if (master) boxes.push(...drawnBoxes(rootElement(parseXml(pkg.readText(master))), slideWidth, slideHeight));
  return boxes;
}
