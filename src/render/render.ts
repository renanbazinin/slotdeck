import { DeckValidationError, checkDeck, deckShapeIssues, fitOptions, freeTextFloor, resolveSlot } from '../deck/validate.ts';
import type { Issue, LayoutMode } from '../deck/validate.ts';
import { isFreeSlide } from '../deck/spec.ts';
import type { ChartValue, DeckSpec, SlideSpec } from '../deck/spec.ts';
import { buildFreeSlide } from '../free/build.ts';
import { freeFor } from '../free/resolve.ts';
import type { FreeResolved } from '../free/resolve.ts';
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
import { NS, childElement, childElements, descendants, openableSizes, parseXml, removeNode, rootElement, serializeXml } from '../pptx/xml.ts';
import type { Element } from '../pptx/xml.ts';
import { bodyTextScale, chartAtScale, chartsReadRightToLeft, defaultFit, effectiveSlots, itemCount, scalesWithText, slotSources, tableAtScale, templateTextSetting } from '../template/manifest.ts';
import type { ChartSlotDef, PatternDef, SlideBox, TemplateManifest } from '../template/manifest.ts';
import { withLayouts } from '../template/compose.ts';
import { readImageInfo } from '../util/image.ts';
import type { ImageInfo } from '../util/image.ts';
import { colorMap, readColors } from '../color/palette.ts';
import type { ColorMap } from '../color/palette.ts';
import { said } from '../messages/en.ts';
import { inheritedBackground, ownBackground, ownTheBackground, paletteOf, recolorChart, recolorDefaults, recolorIcon, recolorInherited, recolorSlide, schemeOf } from '../color/recolor.ts';
import { contrast } from '../color/space.ts';
import { CHART_CONTENT_TYPE, CHART_REL_TYPE, WORKBOOK_CONTENT_TYPE, WORKBOOK_REL_TYPE, buildChartWorkbook, buildChartXml, chartFrameXml } from './chart.ts';
import { chartAlt } from './alt.ts';
import { writeProperties } from './properties.ts';
import type { DeckProperties } from './properties.ts';
import { setPicture } from './picture.ts';
import type { PictureFill } from './picture.ts';
import { fitTable, numberAligns, setTable, tableArea } from './table.ts';
import { applyRepeat, moveShape, placeShape, rowTextIds } from './repeat.ts';
import { dropLeftovers } from './leftovers.ts';
import type { FontMetrics } from '../text/fonts.ts';
import { measurerFor } from '../text/measure.ts';
import { adjustText, setHeadings, setShapeText, setTextSizes } from './text.ts';
import { headingPlaces } from '../deck/headings.ts';
import type { TextOptions, WritingOptions } from './text.ts';

const SLIDE_CONTENT_TYPE = 'application/vnd.openxmlformats-officedocument.presentationml.slide+xml';
const IMAGE_REL_TYPE = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image';
const SLIDE_REL_TYPE = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide';
const P14 = 'http://schemas.microsoft.com/office/powerpoint/2010/main';
const FIRST_SLIDE_ID = 256;
/** The colours of a slide that keeps the template's own in a deck that has others. */
const TEMPLATE_COLORS: { map: ColorMap; key: string } = { map: (hex) => hex, key: '' };

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
  /**
   * Writes what the file says of itself, as a deck and not as its template: its title, language,
   * author and date, its number of slides, and the version of slotdeck that made it; the
   * template's picture of its cover, which Outlook and Explorer show, is taken out. `embedDeck`
   * keeps the deck in the file (see `readEmbeddedDeck`). Left out, the file keeps the template's
   * properties, as before 0.7.0, and is the same file. `generateDeck` and the chat tools set it.
   */
  properties?: DeckProperties;
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
  // A deck of the wrong shape is told so before anything of it is read.
  const shape = deckShapeIssues(deck);
  if (shape.length > 0) throw new DeckValidationError(shape);
  // Layouts made for the deck are built onto the template first. A free slide stands on the
  // template as it was loaded, so that it is the same slide whatever layouts the deck brings.
  const loaded = template;
  template = withLayouts(template, deck.layouts);
  const { manifest } = template;
  const layout = options.layout ?? 'strict';
  const assets = options.assets ?? {};
  const laidOut = new Map<SlideSpec, FreeResolved | string>();
  const freeOf = (slide: SlideSpec): FreeResolved | string => {
    let free = laidOut.get(slide);
    if (free === undefined) laidOut.set(slide, (free = freeFor(loaded, deck, slide)));
    return free;
  };

  // A template loaded with fonts measures its text: warnings are in lines, and text may be shrunk to fit.
  const measure = measurerFor(template);
  const issues = checkDeck(manifest, deck, measure ? { layout, assets: Object.keys(assets), measure } : { layout, assets: Object.keys(assets) }, freeOf);
  const errors = issues.filter((issue) => issue.severity === 'error');
  if (errors.length > 0) throw new DeckValidationError(errors);
  const shrunk: RenderResult['shrunk'] = [];

  const source = template.pkg;
  const out = source.clone();
  const pres = readPresentation(source);
  const types = ContentTypes.read(source);
  const patterns = new Map<string, PatternDef>(manifest.patterns.map((pattern) => [pattern.id, pattern]));
  const directionChanged = (deck.direction ?? manifest.direction) !== manifest.direction;
  // The direction and language filled text is written in: see filledLanguage in text.ts.
  const language = deck.language ?? manifest.language;
  const writing: WritingOptions = { direction: deck.direction ?? manifest.direction, templateDirection: manifest.direction, ...(language ? { language } : {}) };
  // Charts read right to left only where the deck or the template says so.
  const chartsMirrored = chartsReadRightToLeft(manifest, deck);

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

  // A deck may ask for other colours than the template's, for all of its slides or for one. Nothing
  // below runs for a deck that asks for none, so that such a deck gives the bytes it always gave.
  // A template may keep its own colours in every deck; validation says so when a deck asks for others.
  const recoloured = manifest.colors !== false && (deck.colors !== undefined || deck.slides.some((slide) => slide.colors !== undefined));
  const palette = recoloured ? paletteOf(source, manifest) : undefined;
  const colorsOf = (colors: DeckSpec['colors']): { map: ColorMap; key: string } | undefined => {
    const map = palette && colors ? colorMap(palette, colors) : undefined;
    return map ? { map, key: JSON.stringify(readColors(colors).colors) } : undefined;
  };
  const deckColors = colorsOf(deck.colors);
  // A second colour asked of a template that has none changes nothing: that is said, for the deck
  // and for a slide, without a slide number in the issue so that no slide is written again for it.
  if (palette && !palette.accent2) {
    const lacking = (given: unknown, slide?: number): void => {
      if (readColors(given).colors?.accent2) issues.push({ severity: 'warning', ...said('accent2-missing', slide !== undefined ? { slide } : {}) });
    };
    lacking(deck.colors);
    deck.slides.forEach((slide, index) => lacking(slide.colors, index + 1));
  }
  /** Colours asked for that are close to the page of a slide they are used on, where the page does not change with them. */
  const nearPage = new Set<string>();

  const media = new Map<string, { part: string; info: ImageInfo }>();
  const addMedia = (assetId: string, colors?: { map: ColorMap; key: string }): { part: string; info: ImageInfo } => {
    // A pictogram in other colours is another picture; one that has no colour to change is the same one.
    const changed = colors ? recolorIcon(assets[assetId] as Uint8Array, colors.map) : undefined;
    const name = changed ? `${assetId}
${colors?.key}` : assetId;
    const known = media.get(name);
    if (known) return known;
    const data = changed ?? (assets[assetId] as Uint8Array);
    const info = readImageInfo(data);
    let index = media.size + 1;
    let part = `${mediaDir}/slotdeck-${index}.${info.extension}`;
    while (out.has(part)) part = `${mediaDir}/slotdeck-${++index}.${info.extension}`;
    out.write(part, data);
    types.ensureDefault(info.extension, info.contentType);
    const entry = { part, info };
    media.set(name, entry);
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
    out.write(workbookPart, buildChartWorkbook(chart, { rightToLeft: chartsMirrored }));
    out.write(chartPart, openableSizes(buildChartXml(chart, slot, 'rId1', writing.direction, { mirrored: chartsMirrored, ...(language ? { language } : {}) })));
    writeRels(out, chartPart, [{ id: 'rId1', type: WORKBOOK_REL_TYPE, target: relativeTarget(chartPart, workbookPart), external: false }]);
    types.setOverride(chartPart, CHART_CONTENT_TYPE);
    types.ensureDefault('xlsx', WORKBOOK_CONTENT_TYPE);
    return chartPart;
  };

  // Shapes of template slides, for a free slide to copy: each slide is read once.
  const sourceShapes = new Map<number, Map<string, Element>>();
  const shapeFrom = (slideNumber: number, shapeId: string): Element | undefined => {
    let shapes = sourceShapes.get(slideNumber);
    if (!shapes) {
      const ref = pres.slides[slideNumber - 1];
      shapes = ref && source.has(ref.part) ? indexShapes(rootElement(parseXml(source.readText(ref.part)))) : new Map();
      sourceShapes.set(slideNumber, shapes);
    }
    return shapes.get(shapeId);
  };
  const usedSlideParts = new Set<string>();
  /** Where each picture of the deck stands, by slide and shape: what a deck kept in its file names its pictures by. */
  const placed: Array<{ asset: string; slide: number; shape: string }> = [];

  const newSlideParts: string[] = [];
  deck.slides.forEach((slide, index) => {
    const number = index + 1;
    // A slide the software lays out itself has a pattern of its own, worked out for it.
    const free = isFreeSlide(slide) ? (freeOf(slide) as FreeResolved) : undefined;
    const pattern = free ? free.plan.pattern : (patterns.get(slide.pattern) as PatternDef);
    const slideLayout: LayoutMode = free ? 'adaptive' : layout;
    const sourceSlide = pres.slides[pattern.sourceSlide - 1];
    if (!sourceSlide) {
      throw new Error(`Pattern "${pattern.id}" points at slide ${pattern.sourceSlide}, but the template has ${pres.slides.length} slides`);
    }

    usedSlideParts.add(sourceSlide.part);
    const doc = parseXml(source.readText(sourceSlide.part));
    const root = rootElement(doc);
    if (free) buildFreeSlide(doc, root, free.kit, free.plan, shapeFrom, writing);
    // Notes belong to the template's sample content, so they are not carried over.
    let rels = readRels(source, sourceSlide.part).filter((rel) => !isRelType(rel, 'notesSlide'));
    const part = `${slideDir}/slide${number}.xml`;
    // A row of items is first brought to the number the slide asks for; the shapes it copies get new ids.
    if (pattern.repeat) applyRepeat(indexShapes(root), pattern.repeat, itemCount(pattern, slide.items), pres.slideWidth, rowTextIds(pattern), directionChanged);
    const shapes = indexShapes(root);
    const sources = slotSources(pattern, directionChanged, slide.items);
    const textScale = bodyTextScale(template.manifest, deck.textScale);
    // A slide's own colours take the place of the deck's on that slide. A slide that asks for the
    // template's own (`{}`, or colours that change nothing), in a deck in other colours, is kept
    // as the template has it. Colours that are no colours at all (null, a value that is no
    // object, or settings of which none is usable, which the deck's check reports as left out)
    // ask for nothing: such a slide follows the deck.
    const read = slide.colors === undefined || slide.colors === null ? undefined : readColors(slide.colors);
    const ownColors = !!read && (read.colors !== undefined || read.problems.length === 0);
    const colors = !recoloured ? undefined : ownColors ? (colorsOf(slide.colors) ?? (deckColors ? TEMPLATE_COLORS : undefined)) : deckColors;
    // The page of the slide as the template has it: what a chart's text has to read on.
    const pageOf = (): string | undefined => ownBackground(root, schemeOf(source, sourceSlide.part)) ?? inheritedBackground(source, sourceSlide.part)?.color;

    for (const slot of effectiveSlots(pattern, slide.items)) {
      const shape = shapes.get(slot.shapeId);
      if (!shape) {
        throw new Error(`Pattern "${pattern.id}": slot "${slot.name}" points at shape #${slot.shapeId}, which is not on template slide ${pattern.sourceSlide}`);
      }
      const resolved = resolveSlot(slot, { ...slide, slots: slide.slots ?? {} }, deck, slideLayout, sources.get(slot.name));
      if ('error' in resolved) throw new Error(`slide ${number} / ${slot.name}: ${resolved.error}`);
      const action = resolved.action;

      if (action.type === 'remove') {
        removeNode(shape);
      } else if (action.type === 'keep') {
        // A part of the design that a derived pattern shows in another place.
        if ((slot.kind === 'text' || slot.kind === 'image') && slot.box) moveShape(shape, slot.box, pres.slideWidth, pres.slideHeight);
      } else if (action.type === 'text' && slot.kind === 'text') {
        const align = slot.alignOnDirectionChange ?? manifest.alignOnDirectionChange ?? 'flip';
        const text: TextOptions = { flipAlign: directionChanged && align === 'flip', ...writing };
        setShapeText(shape, action.paragraphs, slot.styleRun ? { ...text, styleRun: slot.styleRun } : text);
        if (slot.box) placeShape(shape, slot.box, pres.slideWidth, pres.slideHeight);
        const change: Parameters<typeof adjustText>[1] = {};
        if (slot.fontSize !== undefined) change.size = slot.fontSize;
        if (textScale !== 1 && scalesWithText(slot)) change.scale = textScale;
        if (slot.paragraphGap !== undefined) change.paragraphGap = slot.paragraphGap;
        if (Object.keys(change).length > 0) adjustText(shape, change);
        // Text that does not fit is set smaller, as far as the manifest allows; the measurer says how far.
        if (measure && manifest.shrinkText) {
          // On a free slide of a deck that asks for larger text, no further than the size the text has without that wish.
          const fit = measure.fit(pattern, slot, action.paragraphs, fitOptions(manifest, deck, slide, freeTextFloor(free, slot, action.paragraphs, manifest, slide, measure, deck.direction), action.paragraphs));
          if (fit && fit.scale < 1) {
            setTextSizes(shape, fit.sizes);
            shrunk.push({ slide: number, slot: sources.get(slot.name) ?? slot.name, size: fit.size, scale: Math.round(fit.scale * 1000) / 1000 });
          }
        }
        // The headings of a person's document in the text, last, once its size is set.
        const places = headingPlaces(action.paragraphs, slide.headings);
        if (places.length > 0) setHeadings(shape, places, action.paragraphs, writing.direction, templateTextSetting(manifest, 'characterWidth'));
      } else if (action.type === 'image' && slot.kind === 'image') {
        let image: { part: string; info: ImageInfo };
        try {
          image = addMedia(action.image.asset, slot.imageKind === 'icon' ? colors : undefined);
        } catch {
          // Bytes that are no picture, or no bytes at all: the deck's check knows assets by their ids only.
          throw new DeckValidationError([{ severity: 'error', slide: number, slot: sources.get(slot.name) ?? slot.name, message: `asset "${action.image.asset}" is not a PNG, JPEG or GIF picture` }]);
        }
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
        // A pictogram in other colours is not the picture the deck names.
        if (media.get(action.image.asset) === image) placed.push({ asset: action.image.asset, slide: number, shape: slot.shapeId });
      } else if (action.type === 'table' && slot.kind === 'table') {
        // The colours of its text are written onto the slide, and change with the slide below.
        const box = (slot as { box?: SlideBox }).box;
        // A table of a free slide states its sizes itself and takes the box it was given.
        const planned = slot as { rowHeights?: number[]; columnWidths?: number[]; columnAligns?: Array<'l' | 'r' | undefined> };
        // A table that fills its frame takes the area its sample takes, and sets its columns of numbers as a free slide's are set.
        const frameArea = !box && slot.fill ? tableArea(shape) : undefined;
        const sampleRows = frameArea ? descendants(shape, NS.a, 'tr').length : 0;
        const aligns = box ? planned.columnAligns : frameArea ? numberAligns(action.table, writing.direction ?? 'ltr') : undefined;
        setTable(shape, action.table, tableAtScale(slot, textScale), deck.direction ?? manifest.direction, aligns, writing);
        // Its rows share the frame as the sample's rows do: a heading drawn lower than the rows stays lower.
        if (frameArea && action.table.rows.length + 1 < sampleRows) fitTable(shape, frameArea, descendants(shape, NS.a, 'tr').map((row) => Number(row.getAttribute('h')) || 1));
        // Rows whose cells take more lines than their neighbours' are higher, and columns are as wide as what is written in them asks, as the plan worked out.
        if (box) fitTable(shape, { x: box.x * pres.slideWidth, y: box.y * pres.slideHeight, cx: box.w * pres.slideWidth, cy: box.h * pres.slideHeight }, planned.rowHeights, planned.columnWidths);
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

        const sized = free ? slot : chartAtScale(slot, textScale);
        const chartPart = addChart(action.chart, colors && palette ? recolorChart(sized, colors.map, palette, pageOf()) : sized);
        const relId = nextRelId(rels);
        rels.push({ id: relId, type: CHART_REL_TYPE, target: relativeTarget(part, chartPart), external: false });
        const alt = chartAlt(action.chart, language, writing.direction);
        const frame = doc.importNode(rootElement(parseXml(chartFrameXml(slot.shapeId, relId, area, alt))), true);
        parent.replaceChild(frame, shape);
      }
    }

    setSlideNumber(root, number);

    // Every slot is filled: what the slide shows now is what changes colour.
    if (colors) {
      const inherited = inheritedBackground(source, sourceSlide.part);
      // A colour that is the page's own, or close to it, leaves fills, bars and icons unseen on a
      // page that does not change with the colours. Nothing is refused: the deck is told.
      const page = ownBackground(root, schemeOf(source, sourceSlide.part)) ?? inherited?.color ?? 'FFFFFF';
      const asked = readColors(ownColors ? slide.colors : deck.colors).colors;
      if (colors.map(page) === page) {
        for (const hex of [palette?.accent ? asked?.accent : undefined, palette?.accent2 ? asked?.accent2 : undefined]) if (hex && contrast(hex, page) < 1.3) nearPage.add(hex);
      }
      // A slide in colours of its own takes the background it inherits along; the deck's colours reach it where it is.
      if (ownColors) ownTheBackground(root, inherited?.element);
      recolorSlide(root, colors.map, schemeOf(source, sourceSlide.part), inherited?.surface ?? inherited?.color, deckColors?.map);
    }

    // What the template's sample brought along goes: its add-ins' data, animations of shapes that
    // are gone. Then the slide keeps its layout and what it still points at: pictures that were
    // replaced, the sample's links, its comments and its add-ins' tags are left behind, and so are
    // their parts.
    dropLeftovers(root);
    const used = usedRelationshipIds(root);
    rels = rels.filter((rel) => isRelType(rel, 'slideLayout') || used.has(rel.id));

    out.write(part, openableSizes(serializeXml(doc)));
    writeRels(out, part, rels);
    types.setOverride(part, SLIDE_CONTENT_TYPE);
    newSlideParts.push(part);
  });
  for (const hex of nearPage) issues.push({ severity: 'warning', ...said('color-near-page', { hex }) });

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

  // What the slides inherit (a background, the colours of the theme, the text styles of the
  // presentation) follows the deck's colours too.
  if (deckColors) {
    recolorDefaults(presRoot, deckColors.map, source, usedSlideParts);
    recolorInherited(source, out, usedSlideParts, deckColors.map);
  }

  out.write(pres.part, serializeXml(presDoc));
  writeRels(out, pres.part, presRels);

  if (options.properties) writeProperties(out, types, pres.part, deck, newSlideParts.length, options.properties, { language: manifest.language }, placed);

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
