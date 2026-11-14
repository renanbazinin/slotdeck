import { DeckValidationError, fitOptions, resolveSlot, validateDeck } from '../deck/validate.ts';
import type { Issue, LayoutMode } from '../deck/validate.ts';
import type { ChartValue, DeckSpec } from '../deck/spec.ts';
import {
  ContentTypes,
  PptxPackage,
  isRelType,
  nextRelId,
  pruneUnreachable,
  readRels,
  relativeTarget,
  relsPath,
  writeRels,
} from '../pptx/package.ts';
import type { Relationship } from '../pptx/package.ts';
import { readPresentation } from '../pptx/presentation.ts';
import { NS, childElement, childElements, descendants, parseXml, removeNode, rootElement, serializeXml } from '../pptx/xml.ts';
import type { Element } from '../pptx/xml.ts';
import { bodyTextScale, chartAtScale, defaultFit, effectiveSlots, itemCount, scalesWithText, slotSources, tableAtScale } from '../template/manifest.ts';
import type { ChartSlotDef, PatternDef, TemplateManifest } from '../template/manifest.ts';
import { readImageInfo } from '../util/image.ts';
import type { ImageInfo } from '../util/image.ts';
import { CHART_CONTENT_TYPE, CHART_REL_TYPE, WORKBOOK_CONTENT_TYPE, WORKBOOK_REL_TYPE, buildChartWorkbook, buildChartXml, chartFrameXml } from './chart.ts';
import { setPicture } from './picture.ts';
import type { PictureFill } from './picture.ts';
import { setTable } from './table.ts';
import { applyRepeat, placeShape } from './repeat.ts';
import type { FontMetrics } from '../text/fonts.ts';
import { measurerFor } from '../text/measure.ts';
import { adjustText, setShapeText, setTextSizes } from './text.ts';

const SLIDE_CONTENT_TYPE = 'application/vnd.openxmlformats-officedocument.presentationml.slide+xml';
const IMAGE_REL_TYPE = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image';
const SLIDE_REL_TYPE = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide';
const P14 = 'http://schemas.microsoft.com/office/powerpoint/2010/main';
const FIRST_SLIDE_ID = 256;

/** A template pptx, unzipped once, together with its manifest. Reuse it across renders. */
export interface Template {
  pkg: PptxPackage;
  manifest: TemplateManifest;
  /**
   * Metrics of the template's fonts. With them, text is measured as it will be set: validation
   * reports overflow in lines, the planner repairs by it, and text can be shrunk to fit.
   */
  fonts?: FontMetrics;
}

export interface TemplateOptions {
  /** Metrics of the template's fonts, e.g. from `fontkitMetrics`. Without them budgets are character counts. */
  fonts?: FontMetrics;
}

export function loadTemplate(pptx: Uint8Array, manifest: TemplateManifest, options: TemplateOptions = {}): Template {
  const template: Template = { pkg: PptxPackage.load(pptx), manifest };
  if (options.fonts) template.fonts = options.fonts;
  return template;
}

export interface RenderOptions {
  /** Image bytes by asset id, for the image references in the deck. */
  assets?: Record<string, Uint8Array>;
  /**
   * `strict` (default): every slot must be filled; nothing on a slide moves or disappears.
   * `adaptive`: optional slots may be left out, and their shapes are removed.
   */
  layout?: LayoutMode;
}

export interface RenderResult {
  pptx: Uint8Array;
  /** Non-blocking findings, e.g. text longer than a slot's budget. */
  warnings: Issue[];
  /** Texts that were set smaller to fit their boxes, when the manifest's `shrinkText` allows it. */
  shrunk: Array<{ slide: number; slot: string; size: number; scale: number }>;
}

/**
 * Builds a pptx from a deck spec by cloning the template slide of each pattern
 * and filling its slots. Deterministic: no model calls, no network.
 */
export function renderDeck(template: Template, deck: DeckSpec, options: RenderOptions = {}): RenderResult {
  const { manifest } = template;
  const layout = options.layout ?? 'strict';
  const assets = options.assets ?? {};

  // A template loaded with fonts measures its text: warnings are in lines, and text may be shrunk to fit.
  const measure = measurerFor(template);
  const issues = validateDeck(manifest, deck, measure ? { layout, assets: Object.keys(assets), measure } : { layout, assets: Object.keys(assets) });
  const errors = issues.filter((issue) => issue.severity === 'error');
  if (errors.length > 0) throw new DeckValidationError(errors);
  const shrunk: RenderResult['shrunk'] = [];

  const source = template.pkg;
  const out = source.clone();
  const pres = readPresentation(source);
  const types = ContentTypes.read(source);
  const patterns = new Map<string, PatternDef>(manifest.patterns.map((pattern) => [pattern.id, pattern]));
  const directionChanged = (deck.direction ?? manifest.direction) !== manifest.direction;

  const firstSlide = pres.slides[0];
  if (!firstSlide) throw new Error('The template has no slides');
  const slideDir = firstSlide.part.slice(0, firstSlide.part.lastIndexOf('/'));
  const mediaDir = `${pres.part.slice(0, pres.part.lastIndexOf('/') + 1)}media`;

  // The template's own slides are replaced wholesale by the deck's slides.
  for (const slide of pres.slides) {
    out.delete(slide.part);
    out.delete(relsPath(slide.part));
    types.removeOverride(slide.part);
  }

  const media = new Map<string, { part: string; info: ImageInfo }>();
  const addMedia = (assetId: string): { part: string; info: ImageInfo } => {
    const known = media.get(assetId);
    if (known) return known;
    const data = assets[assetId] as Uint8Array;
    const info = readImageInfo(data);
    let index = media.size + 1;
    let part = `${mediaDir}/slotdeck-${index}.${info.extension}`;
    while (out.has(part)) part = `${mediaDir}/slotdeck-${++index}.${info.extension}`;
    out.write(part, data);
    types.ensureDefault(info.extension, info.contentType);
    const entry = { part, info };
    media.set(assetId, entry);
    return entry;
  };

  // A native chart is two new parts: the chart itself and the workbook that holds its data.
  const partsDir = pres.part.slice(0, pres.part.lastIndexOf('/') + 1);
  let chartCount = 0;
  const addChart = (chart: ChartValue, slot: ChartSlotDef): string => {
    let chartPart: string;
    let workbookPart: string;
    do {
      chartCount++;
      chartPart = `${partsDir}charts/chart${chartCount}.xml`;
      workbookPart = `${partsDir}embeddings/Microsoft_Excel_Sheet${chartCount}.xlsx`;
    } while (out.has(chartPart) || out.has(workbookPart));
    out.write(workbookPart, buildChartWorkbook(chart));
    out.write(chartPart, buildChartXml(chart, slot, 'rId1', deck.direction ?? manifest.direction));
    writeRels(out, chartPart, [{ id: 'rId1', type: WORKBOOK_REL_TYPE, target: relativeTarget(chartPart, workbookPart), external: false }]);
    types.setOverride(chartPart, CHART_CONTENT_TYPE);
    types.ensureDefault('xlsx', WORKBOOK_CONTENT_TYPE);
    return chartPart;
  };

  const newSlideParts: string[] = [];
  deck.slides.forEach((slide, index) => {
    const number = index + 1;
    const pattern = patterns.get(slide.pattern) as PatternDef;
    const sourceSlide = pres.slides[pattern.sourceSlide - 1];
    if (!sourceSlide) {
      throw new Error(`Pattern "${pattern.id}" points at slide ${pattern.sourceSlide}, but the template has ${pres.slides.length} slides`);
    }

    const doc = parseXml(source.readText(sourceSlide.part));
    const root = rootElement(doc);
    // Notes belong to the template's sample content, so they are not carried over.
    let rels = readRels(source, sourceSlide.part).filter((rel) => !isRelType(rel, 'notesSlide'));
    const part = `${slideDir}/slide${number}.xml`;
    // A row of items is first brought to the number the slide asks for; the shapes it copies get new ids.
    if (pattern.repeat) applyRepeat(indexShapes(root), pattern.repeat, itemCount(pattern, slide.items), pres.slideWidth);
    const shapes = indexShapes(root);
    const sources = slotSources(pattern, directionChanged, slide.items);
    const textScale = bodyTextScale(template.manifest, deck.textScale);

    for (const slot of effectiveSlots(pattern, slide.items)) {
      const shape = shapes.get(slot.shapeId);
      if (!shape) {
        throw new Error(`Pattern "${pattern.id}": slot "${slot.name}" points at shape #${slot.shapeId}, which is not on template slide ${pattern.sourceSlide}`);
      }
      const resolved = resolveSlot(slot, { ...slide, slots: slide.slots ?? {} }, deck, layout, sources.get(slot.name));
      if ('error' in resolved) throw new Error(`slide ${number} / ${slot.name}: ${resolved.error}`);
      const action = resolved.action;

      if (action.type === 'remove') {
        removeNode(shape);
      } else if (action.type === 'text' && slot.kind === 'text') {
        const align = slot.alignOnDirectionChange ?? manifest.alignOnDirectionChange ?? 'flip';
        const text = { flipAlign: directionChanged && align === 'flip' };
        setShapeText(shape, action.paragraphs, slot.styleRun ? { ...text, styleRun: slot.styleRun } : text);
        if (slot.box) placeShape(shape, slot.box, pres.slideWidth, pres.slideHeight);
        const change: Parameters<typeof adjustText>[1] = {};
        if (slot.fontSize !== undefined) change.size = slot.fontSize;
        if (textScale !== 1 && scalesWithText(slot)) change.scale = textScale;
        if (slot.paragraphGap !== undefined) change.paragraphGap = slot.paragraphGap;
        if (Object.keys(change).length > 0) adjustText(shape, change);
        // Text that does not fit is set smaller, as far as the manifest allows; the measurer says how far.
        if (measure && manifest.shrinkText) {
          const fit = measure.fit(pattern, slot, action.paragraphs, fitOptions(manifest, deck, slide));
          if (fit && fit.scale < 1) {
            setTextSizes(shape, fit.sizes);
            shrunk.push({ slide: number, slot: sources.get(slot.name) ?? slot.name, size: fit.size, scale: Math.round(fit.scale * 1000) / 1000 });
          }
        }
      } else if (action.type === 'image' && slot.kind === 'image') {
        const image = addMedia(action.image.asset);
        const relId = nextRelId(rels);
        rels.push({ id: relId, type: IMAGE_REL_TYPE, target: relativeTarget(part, image.part), external: false });
        const fill: PictureFill = { relId, width: image.info.width, height: image.info.height, fit: defaultFit(slot) };
        if (action.image.alt !== undefined) fill.alt = action.image.alt;
        if (slot.box) {
          fill.area = {
            x: slot.box.x * pres.slideWidth,
            y: slot.box.y * pres.slideHeight,
            cx: slot.box.w * pres.slideWidth,
            cy: slot.box.h * pres.slideHeight,
          };
        }
        setPicture(shape, fill);
      } else if (action.type === 'table' && slot.kind === 'table') {
        setTable(shape, action.table, tableAtScale(slot, textScale), deck.direction ?? manifest.direction);
      } else if (action.type === 'chart' && slot.kind === 'chart') {
        // The chart takes the place of the sample picture, in the same stacking position.
        const parent = shape.parentNode as Element | null;
        if (!parent || parent.localName !== 'spTree') throw new Error(`slide ${number} / ${slot.name}: a chart slot must point at a shape that is not inside a group`);
        const xfrm = descendants(shape, NS.a, 'xfrm')[0] ?? childElement(shape, NS.p, 'xfrm');
        const off = xfrm && childElement(xfrm, NS.a, 'off');
        const ext = xfrm && childElement(xfrm, NS.a, 'ext');
        const area = slot.box
          ? { x: slot.box.x * pres.slideWidth, y: slot.box.y * pres.slideHeight, cx: slot.box.w * pres.slideWidth, cy: slot.box.h * pres.slideHeight }
          : { x: Number(off?.getAttribute('x') ?? 0), y: Number(off?.getAttribute('y') ?? 0), cx: Number(ext?.getAttribute('cx') ?? 0), cy: Number(ext?.getAttribute('cy') ?? 0) };
        if (!(area.cx > 0 && area.cy > 0)) throw new Error(`slide ${number} / ${slot.name}: the chart slot has no area`);

        const chartPart = addChart(action.chart, chartAtScale(slot, textScale));
        const relId = nextRelId(rels);
        rels.push({ id: relId, type: CHART_REL_TYPE, target: relativeTarget(part, chartPart), external: false });
        const alt = `${action.chart.type} chart: ${action.chart.series.map((series) => series.name).filter(Boolean).join(', ')}`;
        const frame = doc.importNode(rootElement(parseXml(chartFrameXml(slot.shapeId, relId, area, alt))), true);
        parent.replaceChild(frame, shape);
      }
    }

    setSlideNumber(root, number);

    // Images that were replaced leave relationships nobody uses; drop them so their media can go too.
    const used = usedRelationshipIds(root);
    rels = rels.filter((rel) => !isRelType(rel, 'image') || used.has(rel.id));

    out.write(part, serializeXml(doc));
    writeRels(out, part, rels);
    types.setOverride(part, SLIDE_CONTENT_TYPE);
    newSlideParts.push(part);
  });

  // Point the presentation at the new slides.
  const sourcePresRels = readRels(source, pres.part);
  const slideRelType = sourcePresRels.find((rel) => isRelType(rel, 'slide'))?.type ?? SLIDE_REL_TYPE;
  const presRels: Relationship[] = sourcePresRels.filter(
    // Revision history describes the template's slides, not the new ones.
    (rel) => !isRelType(rel, 'slide') && !isRelType(rel, 'changesInfo') && !isRelType(rel, 'revisionInfo'),
  );
  const presDoc = parseXml(source.readText(pres.part));
  const presRoot = rootElement(presDoc);
  const slideList = childElement(presRoot, NS.p, 'sldIdLst');
  if (!slideList) throw new Error('The template presentation has no slide list');
  for (const entry of childElements(slideList)) removeNode(entry);

  const prefix = presRoot.prefix ? `${presRoot.prefix}:` : '';
  const relPrefix = presRoot.lookupPrefix(NS.r) ?? 'r';
  newSlideParts.forEach((part, index) => {
    const relId = nextRelId(presRels);
    presRels.push({ id: relId, type: slideRelType, target: relativeTarget(pres.part, part), external: false });
    const entry = presDoc.createElementNS(NS.p, `${prefix}sldId`);
    entry.setAttribute('id', String(FIRST_SLIDE_ID + index));
    entry.setAttributeNS(NS.r, `${relPrefix}:id`, relId);
    slideList.appendChild(entry);
  });

  // Named sections list the template's slide ids, which no longer exist.
  for (const sections of descendants(presRoot, P14, 'sectionLst')) {
    if (sections.parentNode) removeNode(sections.parentNode);
  }

  out.write(pres.part, serializeXml(presDoc));
  writeRels(out, pres.part, presRels);

  pruneUnreachable(out, types);
  types.save(out);

  return { pptx: out.save(), warnings: issues.filter((issue) => issue.severity === 'warning'), shrunk };
}

/** Maps `p:cNvPr/@id` to the shape element that owns it. */
export function indexShapes(root: Element): Map<string, Element> {
  const shapes = new Map<string, Element>();
  for (const props of descendants(root, NS.p, 'cNvPr')) {
    const shape = props.parentNode?.parentNode;
    const id = props.getAttribute('id');
    if (id && shape && shape.nodeType === 1) shapes.set(id, shape as Element);
  }
  return shapes;
}

/** PowerPoint recomputes slide-number fields, but other viewers show the cached text. */
function setSlideNumber(root: Element, number: number): void {
  for (const field of descendants(root, NS.a, 'fld')) {
    if (field.getAttribute('type') !== 'slidenum') continue;
    const text = childElement(field, NS.a, 't');
    if (text) text.textContent = String(number);
  }
}

function usedRelationshipIds(root: Element): Set<string> {
  const used = new Set<string>();
  for (const element of [root, ...Array.from(root.getElementsByTagNameNS('*', '*'))]) {
    for (let i = 0; i < element.attributes.length; i++) {
      const attr = element.attributes.item(i);
      if (attr && attr.namespaceURI === NS.r) used.add(attr.value);
    }
  }
  return used;
}
