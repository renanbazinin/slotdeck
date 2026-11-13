import { isRelType, readRels, resolveTarget } from '../pptx/package.ts';
import type { PptxPackage } from '../pptx/package.ts';
import { readPresentation } from '../pptx/presentation.ts';
import { NS, childElement, childElements, descendants, parseXml, path, rootElement } from '../pptx/xml.ts';
import type { Element } from '../pptx/xml.ts';
import { readImageInfo } from '../util/image.ts';

export type ShapeKind = 'text' | 'picture' | 'shape' | 'line' | 'group' | 'table' | 'other';

/** Position and size as fractions of the slide (0..1), after group transforms. */
export interface Box {
  x: number;
  y: number;
  w: number;
  h: number;
}

export interface ShapeInfo {
  /** `p:cNvPr/@id`, unique within the slide. This is what a manifest slot points at. */
  id: string;
  name: string;
  kind: ShapeKind;
  box: Box;
  /** Id of the enclosing group, when nested. */
  groupId?: string;
  /** Placeholder type when the shape is a layout placeholder (e.g. `sldNum`). */
  placeholder?: string;
  /** Paragraph texts, for shapes that carry text. */
  paragraphs?: string[];
  /** Distinct font sizes in points, largest first. */
  fontSizes?: number[];
  image?: {
    part: string;
    width?: number;
    height?: number;
    cropped: boolean;
    geometry: string;
  };
  /** Size of a table, with the header row counted. */
  table?: { columns: number; rows: number };
}

export interface SlideInfo {
  /** 1-based position in the presentation. */
  number: number;
  part: string;
  layout: string;
  shapes: ShapeInfo[];
}

export interface TemplateInfo {
  /** Slide size in EMU. */
  slideWidth: number;
  slideHeight: number;
  /** Size in points of text that sets no size of its own. */
  defaultFontSize: number;
  slides: SlideInfo[];
}

interface Transform {
  sx: number;
  sy: number;
  tx: number;
  ty: number;
}

const IDENTITY: Transform = { sx: 1, sy: 1, tx: 0, ty: 0 };

/** Lists every shape on every slide: the raw material for writing a template manifest. */
export function inspectTemplate(pkg: PptxPackage): TemplateInfo {
  const pres = readPresentation(pkg);
  const slides = pres.slides.map((ref, index): SlideInfo => {
    const rels = readRels(pkg, ref.part);
    const targets = new Map(rels.map((rel) => [rel.id, resolveTarget(ref.part, rel.target)]));
    const layoutRel = rels.find((rel) => isRelType(rel, 'slideLayout'));
    const root = rootElement(parseXml(pkg.readText(ref.part)));
    const tree = path(root, [NS.p, 'cSld'], [NS.p, 'spTree']);

    const shapes: ShapeInfo[] = [];
    if (tree) walk(tree, IDENTITY, undefined);

    function walk(container: Element, transform: Transform, groupId: string | undefined): void {
      for (const el of childElements(container, NS.p)) {
        const local = el.localName;
        if (local !== 'sp' && local !== 'pic' && local !== 'grpSp' && local !== 'cxnSp' && local !== 'graphicFrame') {
          continue;
        }
        const props = descendants(el, NS.p, 'cNvPr')[0];
        const id = props?.getAttribute('id') ?? '';
        const xfrm =
          local === 'grpSp'
            ? path(el, [NS.p, 'grpSpPr'], [NS.a, 'xfrm'])
            : local === 'graphicFrame'
              ? childElement(el, NS.p, 'xfrm')
              : path(el, [NS.p, 'spPr'], [NS.a, 'xfrm']);
        const info: ShapeInfo = {
          id,
          name: props?.getAttribute('name') ?? '',
          kind: 'other',
          box: toBox(xfrm, transform, pres.slideWidth, pres.slideHeight),
        };
        if (groupId !== undefined) info.groupId = groupId;

        const placeholder = descendants(el, NS.p, 'ph')[0];
        if (placeholder && local !== 'grpSp') info.placeholder = placeholder.getAttribute('type') ?? 'body';

        if (local === 'grpSp') {
          info.kind = 'group';
          shapes.push(info);
          walk(el, compose(transform, xfrm), id);
          continue;
        }

        if (local === 'pic') {
          info.kind = 'picture';
          const blipFill = childElement(el, NS.p, 'blipFill');
          const blip = blipFill && childElement(blipFill, NS.a, 'blip');
          const part = targets.get(blip?.getAttributeNS(NS.r, 'embed') ?? '');
          const spPr = childElement(el, NS.p, 'spPr');
          const preset = spPr && childElement(spPr, NS.a, 'prstGeom');
          if (part) {
            info.image = {
              part,
              cropped: !!(blipFill && hasCrop(childElement(blipFill, NS.a, 'srcRect'))),
              geometry: preset?.getAttribute('prst') ?? 'custom',
            };
            if (pkg.has(part)) {
              try {
                const size = readImageInfo(pkg.read(part));
                info.image.width = size.width;
                info.image.height = size.height;
              } catch {
                // Unknown image format: leave the size out.
              }
            }
          }
        } else if (local === 'cxnSp') {
          info.kind = 'line';
        } else if (local === 'graphicFrame') {
          const table = descendants(el, NS.a, 'tbl')[0];
          info.kind = table ? 'table' : 'other';
          if (table) {
            info.table = { columns: descendants(table, NS.a, 'gridCol').length, rows: descendants(table, NS.a, 'tr').length };
            const sizes = new Set(descendants(table, NS.a, 'rPr').map((run) => Number(run.getAttribute('sz') ?? 0) / 100).filter((size) => size > 0));
            if (sizes.size > 0) info.fontSizes = [...sizes].sort((a, b) => b - a);
          }
        } else {
          info.kind = 'shape';
        }

        if (local === 'sp') {
          const body = childElement(el, NS.p, 'txBody');
          const paragraphs = body
            ? childElements(body, NS.a, 'p').map((p) =>
                descendants(p, NS.a, 't')
                  .map((t) => t.textContent ?? '')
                  .join(''),
              )
            : [];
          if (paragraphs.some((text) => text.trim() !== '')) {
            info.kind = 'text';
            info.paragraphs = paragraphs;
            const sizes = new Set<number>();
            for (const run of descendants(el, NS.a, 'rPr')) {
              const size = run.getAttribute('sz');
              if (size) sizes.add(Number(size) / 100);
            }
            info.fontSizes = [...sizes].sort((a, b) => b - a);
          }
        }
        shapes.push(info);
      }
    }

    return {
      number: index + 1,
      part: ref.part,
      layout: layoutRel ? resolveTarget(ref.part, layoutRel.target) : '',
      shapes,
    };
  });

  // 18 pt is what PowerPoint uses when a presentation states no default.
  return { slideWidth: pres.slideWidth, slideHeight: pres.slideHeight, defaultFontSize: pres.defaultFontSize ?? 18, slides };
}

function hasCrop(srcRect: Element | undefined): boolean {
  if (!srcRect) return false;
  return ['l', 't', 'r', 'b'].some((side) => Number(srcRect.getAttribute(side) ?? 0) !== 0);
}

function point(el: Element | undefined, xName: string, yName: string): { x: number; y: number } {
  return { x: Number(el?.getAttribute(xName) ?? 0), y: Number(el?.getAttribute(yName) ?? 0) };
}

/** Maps a group's child coordinate space into its parent's space. */
function compose(parent: Transform, groupXfrm: Element | undefined): Transform {
  if (!groupXfrm) return parent;
  const off = point(childElement(groupXfrm, NS.a, 'off'), 'x', 'y');
  const ext = point(childElement(groupXfrm, NS.a, 'ext'), 'cx', 'cy');
  const chOff = point(childElement(groupXfrm, NS.a, 'chOff'), 'x', 'y');
  const chExt = point(childElement(groupXfrm, NS.a, 'chExt'), 'cx', 'cy');
  const sx = chExt.x ? ext.x / chExt.x : 1;
  const sy = chExt.y ? ext.y / chExt.y : 1;
  return {
    sx: parent.sx * sx,
    sy: parent.sy * sy,
    tx: parent.tx + parent.sx * (off.x - chOff.x * sx),
    ty: parent.ty + parent.sy * (off.y - chOff.y * sy),
  };
}

function toBox(xfrm: Element | undefined, transform: Transform, slideWidth: number, slideHeight: number): Box {
  const off = point(xfrm && childElement(xfrm, NS.a, 'off'), 'x', 'y');
  const ext = point(xfrm && childElement(xfrm, NS.a, 'ext'), 'cx', 'cy');
  const round = (value: number) => Math.round(value * 10000) / 10000;
  return {
    x: round((transform.tx + transform.sx * off.x) / slideWidth),
    y: round((transform.ty + transform.sy * off.y) / slideHeight),
    w: round((transform.sx * ext.x) / slideWidth),
    h: round((transform.sy * ext.y) / slideHeight),
  };
}
