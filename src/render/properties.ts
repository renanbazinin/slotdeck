// The file properties of a rendered deck: what Outlook, Explorer and File > Info show of it without
// opening it (its title, its number of slides, its picture), and the deck itself kept in the file
// where it is asked for.
import type { DeckSpec } from '../deck/spec.ts';
import { isTextValue, toParagraphs } from '../deck/spec.ts';
import { ContentTypes, PptxPackage, isRelType, nextRelId, readRels, relativeTarget, resolveTarget, writeRels } from '../pptx/package.ts';
import { readPresentation } from '../pptx/presentation.ts';
import { NS, XML_HEADER, childElements, descendants, escapeXmlAttr, parseXml, path, removeNode, rootElement, serializeXml } from '../pptx/xml.ts';
import type { Element } from '../pptx/xml.ts';
import { SLOTDECK_VERSION } from '../version.ts';

/** What a rendered deck's file says of itself. Every field may be left out. */
export interface DeckProperties {
  /** The title its file properties give it. Left out, the deck's `title`, and else the title of its first slide that has one. */
  title?: string;
  /** Its language, as a BCP 47 tag. Left out, the deck's language, and else the template's. */
  language?: string;
  /** Who wrote it. Left out, no one is named: the template's author is not the deck's. */
  author?: string;
  /**
   * When it was made, as a `Date` or a date the `Date` constructor reads (`2026-11-21T09:30:00Z`).
   * Left out, no date is written, so that the same deck always gives the same file.
   */
  date?: Date | string;
  /**
   * Keeps the deck, as its JSON, inside the file, with the names of its pictures in it, so that a
   * file that comes back by mail can be rendered again or looked into: `readEmbeddedDeck` takes it
   * out. The deck's `brief` and `document` are left out of it: they are what a person asked and
   * gave, and are not needed to render it. Off unless set.
   */
  embedDeck?: boolean;
}

const CORE_REL = 'http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties';
const APP_REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties';
const CUSTOM_REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/custom-properties';
const THUMBNAIL_REL = 'http://schemas.openxmlformats.org/package/2006/relationships/metadata/thumbnail';
const CUSTOM_XML_REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/customXml';
const CUSTOM_XML_PROPS_REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/customXmlProps';
const CORE_TYPE = 'application/vnd.openxmlformats-package.core-properties+xml';
const APP_TYPE = 'application/vnd.openxmlformats-officedocument.extended-properties+xml';
const CUSTOM_TYPE = 'application/vnd.openxmlformats-officedocument.custom-properties+xml';
const CUSTOM_XML_PROPS_TYPE = 'application/vnd.openxmlformats-officedocument.customXmlProperties+xml';
const NS_CUSTOM = 'http://schemas.openxmlformats.org/officeDocument/2006/custom-properties';
const NS_VT = 'http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes';
const NS_APP = 'http://schemas.openxmlformats.org/officeDocument/2006/extended-properties';
/** The format id every custom property of an Office file has. */
const CUSTOM_FMTID = '{D5CDD505-2E9C-101B-9397-08002B2CF9AE}';
/** The namespace of the part that keeps a deck in its file. */
export const EMBEDDED_DECK_NS = 'urn:slotdeck:deck';
/** The id of that part among a file's custom XML parts: the same in every file, so that the same deck gives the same file. */
const EMBEDDED_DECK_ID = '{6B1D3F0A-5C2E-4E8B-9A7D-3F1C2B4D5E6F}';
/** The name of the custom property that says which version of slotdeck made the file. */
export const VERSION_PROPERTY = 'slotdeck';

const text = (value: string) => escapeXmlAttr(value);

/** A date as the file writes it: to the second, in UTC. */
function w3cdtf(value: Date | string): string {
  const date = value instanceof Date ? value : new Date(value);
  if (Number.isNaN(date.getTime())) throw new RangeError(`properties.date is ${JSON.stringify(String(value))}; it is a Date, or a date such as "2026-11-21T09:30:00Z"`);
  return date.toISOString().replace(/\.\d{3}Z$/, 'Z');
}

/** The title of the first slide that has one, as one line. */
function firstTitle(deck: DeckSpec): string | undefined {
  for (const slide of deck.slides) {
    const value = slide.slots?.title;
    if (!isTextValue(value)) continue;
    const line = toParagraphs(value).filter(Boolean).join(' ').trim();
    if (line) return line;
  }
  return undefined;
}

/** The part a relationship of the package root points at, and that relationship. */
function rootPart(pkg: PptxPackage, type: string): string | undefined {
  const rel = readRels(pkg, '').find((entry) => entry.type === type && !entry.external);
  return rel ? resolveTarget('', rel.target) : undefined;
}

/** Points the package root at a part of one kind, in place of any it pointed at before. */
function setRootPart(pkg: PptxPackage, types: ContentTypes, type: string, part: string, contentType: string, xml: string): void {
  const rels = readRels(pkg, '');
  const known = rels.find((entry) => entry.type === type && !entry.external);
  if (known) known.target = part;
  else rels.push({ id: nextRelId(rels), type, target: part, external: false });
  writeRels(pkg, '', rels);
  pkg.write(part, xml);
  types.setOverride(part, contentType);
}

/**
 * Writes what a deck's file says of itself: a title, a language, an author and a date where they
 * are known, the number of slides it has, and the version of slotdeck that made it (a custom
 * property, beside any the template has, such as a label of how confidential it is). The
 * template's own picture of its cover, its sample titles and its dates are taken out. With
 * `embedDeck`, the deck is kept in the file. `pictures` says where each asset stands, by slide and shape.
 */
export function writeProperties(pkg: PptxPackage, types: ContentTypes, presentationPart: string, deck: DeckSpec, slides: number, properties: DeckProperties, defaults: { language?: string | undefined }, pictures: ReadonlyArray<PlacedPicture>): void {
  // The core properties: fresh, so that nothing of the template's (its title, its author, its dates) is left.
  const title = properties.title ?? deck.title ?? firstTitle(deck);
  const language = properties.language ?? deck.language ?? defaults.language;
  const date = properties.date === undefined ? undefined : w3cdtf(properties.date);
  const core =
    `${XML_HEADER}<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:dcmitype="http://purl.org/dc/dcmitype/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">` +
    (title ? `<dc:title>${text(title)}</dc:title>` : '') +
    (properties.author ? `<dc:creator>${text(properties.author)}</dc:creator><cp:lastModifiedBy>${text(properties.author)}</cp:lastModifiedBy>` : '') +
    (language ? `<dc:language>${text(language)}</dc:language>` : '') +
    '<cp:revision>1</cp:revision>' +
    (date ? `<dcterms:created xsi:type="dcterms:W3CDTF">${date}</dcterms:created><dcterms:modified xsi:type="dcterms:W3CDTF">${date}</dcterms:modified>` : '') +
    '</cp:coreProperties>';
  setRootPart(pkg, types, CORE_REL, rootPart(pkg, CORE_REL) ?? 'docProps/core.xml', CORE_TYPE, core);

  // The extended properties: the template's (its company, its format), with the deck's number of
  // slides, and without the lists of the template's fonts and slide titles, or its counts of words.
  const appPart = rootPart(pkg, APP_REL) ?? 'docProps/app.xml';
  const appDoc = parseXml(pkg.has(appPart) ? pkg.readText(appPart) : `<Properties xmlns="${NS_APP}" xmlns:vt="${NS_VT}"><Application>Microsoft Office PowerPoint</Application></Properties>`);
  const app = rootElement(appDoc);
  const set = (name: string, value: string) => {
    let element = childElements(app).find((child) => child.localName === name);
    if (!element) {
      element = appDoc.createElementNS(app.namespaceURI ?? NS_APP, app.prefix ? `${app.prefix}:${name}` : name);
      app.appendChild(element);
    }
    element.textContent = value;
  };
  for (const child of childElements(app)) if (['HeadingPairs', 'TitlesOfParts', 'Words', 'Paragraphs', 'TotalTime', 'MMClips'].includes(child.localName ?? '')) removeNode(child);
  set('Slides', String(slides));
  set('Notes', '0');
  set('HiddenSlides', '0');
  setRootPart(pkg, types, APP_REL, appPart, APP_TYPE, serializeXml(appDoc));

  // The version that made it, among the custom properties the template has.
  const customPart = rootPart(pkg, CUSTOM_REL) ?? 'docProps/custom.xml';
  const customDoc = parseXml(pkg.has(customPart) ? pkg.readText(customPart) : `<Properties xmlns="${NS_CUSTOM}" xmlns:vt="${NS_VT}"/>`);
  const custom = rootElement(customDoc);
  const existing = childElements(custom).filter((child) => child.localName === 'property');
  for (const property of existing) if (property.getAttribute('name') === VERSION_PROPERTY) removeNode(property);
  const pid = Math.max(1, ...existing.filter((property) => property.parentNode).map((property) => Number(property.getAttribute('pid')) || 1)) + 1;
  const prefix = custom.prefix ? `${custom.prefix}:` : '';
  const property = customDoc.createElementNS(custom.namespaceURI ?? NS_CUSTOM, `${prefix}property`);
  property.setAttribute('fmtid', CUSTOM_FMTID);
  property.setAttribute('pid', String(pid));
  property.setAttribute('name', VERSION_PROPERTY);
  const value = customDoc.createElementNS(NS_VT, `${custom.lookupPrefix(NS_VT) ?? 'vt'}:lpwstr`);
  value.textContent = SLOTDECK_VERSION;
  property.appendChild(value);
  custom.appendChild(property);
  setRootPart(pkg, types, CUSTOM_REL, customPart, CUSTOM_TYPE, serializeXml(customDoc));

  // The picture of the template's cover, which Explorer and Outlook would show for every deck.
  writeRels(pkg, '', readRels(pkg, '').filter((rel) => rel.type !== THUMBNAIL_REL));

  if (properties.embedDeck) embedDeck(pkg, types, presentationPart, deck, pictures);
}

/** A picture of the deck in its file: the asset, the slide it is on (from 1) and the id of its shape. */
export interface PlacedPicture {
  asset: string;
  slide: number;
  shape: string;
}

/**
 * Keeps the deck in a custom XML part of the presentation, which PowerPoint carries along when the
 * file is saved again. Its pictures are named by slide and shape, which PowerPoint keeps, and not
 * by the names of their parts, which it changes.
 */
function embedDeck(pkg: PptxPackage, types: ContentTypes, presentationPart: string, deck: DeckSpec, pictures: ReadonlyArray<PlacedPicture>): void {
  const { brief: _brief, document: _document, ...kept } = deck;
  const assets: Record<string, { slide: number; shape: string }> = {};
  for (const picture of pictures) assets[picture.asset] ??= { slide: picture.slide, shape: picture.shape };
  const payload = JSON.stringify({ slotdeck: SLOTDECK_VERSION, deck: kept, assets });
  let index = 1;
  while (pkg.has(`customXml/item${index}.xml`) || pkg.has(`customXml/itemProps${index}.xml`)) index++;
  const item = `customXml/item${index}.xml`;
  const props = `customXml/itemProps${index}.xml`;
  pkg.write(item, `${XML_HEADER}<deck xmlns="${EMBEDDED_DECK_NS}">${text(payload)}</deck>`);
  pkg.write(
    props,
    `${XML_HEADER}<ds:datastoreItem ds:itemID="${EMBEDDED_DECK_ID}" xmlns:ds="http://schemas.openxmlformats.org/officeDocument/2006/customXml"><ds:schemaRefs><ds:schemaRef ds:uri="${EMBEDDED_DECK_NS}"/></ds:schemaRefs></ds:datastoreItem>`,
  );
  writeRels(pkg, item, [{ id: 'rId1', type: CUSTOM_XML_PROPS_REL, target: relativeTarget(item, props), external: false }]);
  types.ensureDefault('xml', 'application/xml');
  types.setOverride(props, CUSTOM_XML_PROPS_TYPE);
  const rels = readRels(pkg, presentationPart);
  rels.push({ id: nextRelId(rels), type: CUSTOM_XML_REL, target: relativeTarget(presentationPart, item), external: false });
  writeRels(pkg, presentationPart, rels);
}

/** A deck kept in its file (see `DeckProperties.embedDeck`): the deck, the version that made the file, and the bytes of its pictures by asset id. */
export interface EmbeddedDeck {
  deck: DeckSpec;
  version: string;
  assets: Record<string, Uint8Array>;
}

/**
 * Takes out of a pptx the deck that `renderDeck` kept in it with `properties: { embedDeck: true }`,
 * with the pictures it uses, so that it can be rendered again: `renderDeck(template, found.deck,
 * { assets: found.assets })`. `undefined` for a file that keeps no deck. Throws for bytes that are
 * no pptx.
 */
export function readEmbeddedDeck(pptx: Uint8Array): EmbeddedDeck | undefined {
  const pkg = PptxPackage.load(pptx);
  const office = readRels(pkg, '').find((rel) => isRelType(rel, 'officeDocument'));
  if (!office) return undefined;
  const presentation = resolveTarget('', office.target);
  for (const rel of readRels(pkg, presentation)) {
    if (rel.type !== CUSTOM_XML_REL) continue;
    const part = resolveTarget(presentation, rel.target);
    if (!pkg.has(part)) continue;
    const root: Element = rootElement(parseXml(pkg.readText(part)));
    if (root.namespaceURI !== EMBEDDED_DECK_NS) continue;
    const found = JSON.parse(root.textContent ?? '') as { slotdeck: string; deck: DeckSpec; assets?: Record<string, { slide: number; shape: string }> };
    const slides = readPresentation(pkg).slides;
    const assets: Record<string, Uint8Array> = {};
    for (const [asset, where] of Object.entries(found.assets ?? {})) {
      const slide = slides[where.slide - 1];
      if (!slide || !pkg.has(slide.part)) continue;
      const pic = descendants(rootElement(parseXml(pkg.readText(slide.part))), NS.p, 'pic').find((element) => path(element, [NS.p, 'nvPicPr'], [NS.p, 'cNvPr'])?.getAttribute('id') === where.shape);
      const embed = pic ? descendants(pic, NS.a, 'blip')[0]?.getAttributeNS(NS.r, 'embed') : undefined;
      const rel = embed ? readRels(pkg, slide.part).find((entry) => entry.id === embed && !entry.external) : undefined;
      const media = rel ? resolveTarget(slide.part, rel.target) : undefined;
      if (media && pkg.has(media)) assets[asset] = pkg.read(media);
    }
    return { deck: found.deck, version: found.slotdeck, assets };
  }
  return undefined;
}

