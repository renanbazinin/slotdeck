// Other layouts for a slide that already has its content. The content is read out of the slide
// by what each slot is for (a title, a body, items with a heading and a text) and put into every
// other pattern that has a place for all of it. Code only: no model is asked and nothing is
// rewritten, so an alternative says exactly what the slide said.
import { effectiveSlots, itemRange } from '../template/manifest.ts';
import type { ImageKind, ImageSlotDef, PatternDef, SlotDef, TemplateManifest, TextSlotDef } from '../template/manifest.ts';
import { isImageRef, toParagraphs } from './spec.ts';
import type { DeckSpec, ImageRef, SlideSpec, SlotValue } from './spec.ts';
import { validateDeck } from './validate.ts';
import type { Issue, LayoutMode } from './validate.ts';

export interface SlideAlternative {
  /** Id of a pattern that has a place for everything the slide says. */
  pattern: string;
  /** What that pattern is for, from the manifest. */
  description: string;
  /** The slide written for that pattern. Put it in the deck in place of the original. */
  slide: SlideSpec;
  /** What was done to the content where it was not copied slot for slot, a sentence each. */
  changes: string[];
  /** Slots of the original slide whose photo or icon the pattern has no place for. */
  dropped: string[];
  /** What `validateDeck` says about the new slide, such as a text over its budget. Empty when everything fits. */
  warnings: Issue[];
}

export interface AlternativesOptions {
  /** `strict` (default): every slot of the new pattern must be filled. `adaptive`: its optional slots may stay empty. */
  layout?: LayoutMode;
  /** Consider only these patterns. */
  patterns?: string[];
}

type TextPart = 'heading' | 'text' | 'figure';
type PicturePart = 'icon' | 'picture';
type Part = TextPart | PicturePart;

/** What the slots of a pattern are for, as far as their names and limits show. */
interface Anatomy {
  title?: TextSlotDef;
  subtitle?: TextSlotDef;
  /** A large number that stands on its own. */
  figure?: TextSlotDef;
  /** The running text or list. */
  body?: TextSlotDef;
  /** The one photo or graphic that is not part of an item. */
  picture?: ImageSlotDef;
  /** Every other slot. It can only be filled from, or put into, a slot of the same name. */
  others: SlotDef[];
  /** The parts of one item of the row. `clear` is false when the item's slots cannot be told apart. */
  row?: { parts: Map<Part, { template: string; slot: SlotDef }>; clear: boolean };
}

interface Item {
  heading?: string[];
  text?: string[];
  figure?: string[];
  icon?: ImageRef;
  picture?: ImageRef;
}

interface Content {
  title?: string[];
  subtitle?: string[];
  figure?: string[];
  body?: string[];
  /** Whether the body's slot is a list of points, as far as its paragraph limit shows. */
  bodyIsList: boolean;
  picture?: { ref: ImageRef; kind: ImageKind; from: string };
  others: Map<string, { value: SlotValue; slot: SlotDef }>;
  items: Item[];
  /** What one item is called, for the sentences in `changes`. */
  itemName: string;
  /** The icons and pictures of the items, with the deck slot each came from. */
  itemPictures: Array<{ from: string; part: PicturePart; kind: ImageKind }>;
}

const fillable = (slot: SlotDef) => slot.empty !== 'keep' && slot.empty !== 'drop' && !(slot.kind === 'text' && slot.deckField);
const sizeOf = (slot: TextSlotDef) => slot.fontSize ?? slot.size ?? 0;
/** A big, short text is a figure: "87%", "3.5x". */
const isFigure = (slot: TextSlotDef) => (slot.maxChars ?? Infinity) <= 10 && sizeOf(slot) >= 24;
const largest = (slots: TextSlotDef[]) => slots.slice().sort((a, b) => (b.maxChars ?? 0) - (a.maxChars ?? 0))[0];
const plural = (count: number, word: string) => `${count} ${word}${count === 1 ? '' : 's'}`;
const empty = (value: SlotValue | undefined): value is null | undefined => value === null || value === undefined || (typeof value === 'string' && value.trim() === '');
/** A photo or an icon illustrates and may be left out; a supplied graphic is something the slide says. */
const illustrates = (kind: ImageKind) => kind !== 'graphic';
const asValue = (text: string[]): SlotValue => (text.length === 1 ? (text[0] as string) : text);
const isList = (slot: TextSlotDef | undefined) => (slot?.maxParagraphs ?? 1) > 1;

function textOf(value: SlotValue | undefined): string[] | undefined {
  if (typeof value !== 'string' && !Array.isArray(value)) return undefined;
  const paragraphs = toParagraphs(value).filter((paragraph) => paragraph !== '');
  return paragraphs.length > 0 ? paragraphs : undefined;
}

function anatomy(pattern: PatternDef): Anatomy {
  const repeat = pattern.repeat;
  const inRow = new Set(repeat ? repeat.shapes.flatMap((_, index) => repeat.slots.map((name) => name.replace('{n}', String(index + 1)))) : []);
  const own = pattern.slots.filter((slot) => fillable(slot) && !inRow.has(slot.name));
  const texts = own.filter((slot): slot is TextSlotDef => slot.kind === 'text');
  const take = (slot: TextSlotDef | undefined): TextSlotDef | undefined => {
    if (slot) texts.splice(texts.indexOf(slot), 1);
    return slot;
  };
  const shape: Anatomy = { others: [] };
  const title = take(texts.find((slot) => /title$/i.test(slot.name) && !/sub/i.test(slot.name)));
  const subtitle = take(texts.find((slot) => /sub(title|heading)$/i.test(slot.name)));
  const figure = take(texts.find(isFigure));
  // Of what is left, the text with the most room is the body.
  const body = take(largest(texts));
  if (title) shape.title = title;
  if (subtitle) shape.subtitle = subtitle;
  if (figure) shape.figure = figure;
  if (body) shape.body = body;
  const pictures = own.filter((slot): slot is ImageSlotDef => slot.kind === 'image' && slot.imageKind !== 'icon');
  if (pictures.length === 1 && pictures[0]) shape.picture = pictures[0];
  shape.others = [...texts, ...own.filter((slot) => slot.kind !== 'text' && slot !== shape.picture)];

  if (repeat) {
    const parts = new Map<Part, { template: string; slot: SlotDef }>();
    let clear = true;
    const itemTexts: Array<{ template: string; slot: TextSlotDef }> = [];
    for (const template of repeat.slots) {
      const slot = pattern.slots.find((candidate) => candidate.name === template.replace('{n}', '1'));
      if (!slot || !fillable(slot)) continue;
      if (slot.kind === 'text') itemTexts.push({ template, slot });
      else if (slot.kind === 'image') {
        const part: PicturePart = slot.imageKind === 'icon' ? 'icon' : 'picture';
        if (parts.has(part)) clear = false;
        parts.set(part, { template, slot });
      } else clear = false;
    }
    const number = itemTexts.length > 1 ? itemTexts.find((entry) => isFigure(entry.slot)) : undefined;
    if (number) parts.set('figure', number);
    const rest = itemTexts.filter((entry) => entry !== number).sort((a, b) => (a.slot.maxChars ?? 0) - (b.slot.maxChars ?? 0));
    // Of two texts, the one with less room is the heading. More than two cannot be told apart.
    if (rest.length > 2) clear = false;
    if (rest.length === 2 && rest[0]) parts.set('heading', rest[0]);
    const main = rest.at(-1);
    if (main) parts.set('text', main);
    shape.row = { parts, clear };
  }
  return shape;
}

function contentOf(pattern: PatternDef, shape: Anatomy, slide: SlideSpec): Content {
  const content: Content = { others: new Map(), items: [], itemName: pattern.repeat?.item ?? 'item', itemPictures: [], bodyIsList: isList(shape.body) };
  const read = (slot: TextSlotDef | undefined) => (slot ? textOf(slide.slots[slot.name]) : undefined);
  const title = read(shape.title);
  const subtitle = read(shape.subtitle);
  const figure = read(shape.figure);
  const body = read(shape.body);
  if (title) content.title = title;
  if (subtitle) content.subtitle = subtitle;
  if (figure) content.figure = figure;
  if (body) content.body = body;
  const picture = shape.picture && slide.slots[shape.picture.name];
  if (shape.picture && isImageRef(picture)) content.picture = { ref: picture, kind: shape.picture.imageKind, from: shape.picture.name };
  for (const slot of shape.others) {
    const value = slide.slots[slot.name];
    if (!empty(value)) content.others.set(slot.name, { value, slot });
  }
  const range = itemRange(pattern);
  if (shape.row && range) {
    for (let n = 1; n <= (slide.items ?? range.drawn); n++) {
      const item: Item = {};
      for (const [part, { template, slot }] of shape.row.parts) {
        const name = template.replace('{n}', String(n));
        const value = slide.slots[name];
        if (part === 'icon' || part === 'picture') {
          if (isImageRef(value) && slot.kind === 'image') {
            item[part] = value;
            content.itemPictures.push({ from: name, part, kind: slot.imageKind });
          }
        } else {
          const text = textOf(value);
          if (text) item[part] = text;
        }
      }
      content.items.push(item);
    }
  }
  return content;
}

class NoPlace extends Error {}
const no = (): never => {
  throw new NoPlace();
};

/** A heading and a text as one line: "Heading: text". */
function joined(lead: string[] | undefined, text: string[] | undefined): string {
  const head = (lead ?? []).join(' ');
  const rest = (text ?? []).join(' ');
  if (!head || !rest) return head || rest;
  return /[:.!?…]$/.test(head) ? `${head} ${rest}` : `${head}: ${rest}`;
}

/** The two halves of "Heading: text", or nothing when the line is not written that way. */
function split(line: string): [string, string] | undefined {
  const match = /^(.{1,80}?)(?::\s+|\s+[–—-]\s+)(\S.*)$/.exec(line);
  return match ? [match[1] as string, match[2] as string] : undefined;
}

interface Draft {
  slots: Record<string, SlotValue>;
  items?: number;
  changes: string[];
  dropped: string[];
}

/** Copies every value to the slot of the same name: the way between variants of one slide. */
function byName(source: PatternDef, target: PatternDef, slide: SlideSpec): Draft {
  const range = itemRange(target);
  const items = range ? (slide.items ?? itemRange(source)?.drawn ?? range.drawn) : undefined;
  if (range && items !== undefined && (items < range.min || items > range.max)) no();
  const places = new Map(effectiveSlots(target, items).filter(fillable).map((slot) => [slot.name, slot]));
  const draft: Draft = { slots: {}, changes: [], dropped: [] };
  if (items !== undefined) draft.items = items;
  for (const slot of effectiveSlots(source, slide.items).filter(fillable)) {
    const value = slide.slots[slot.name];
    if (empty(value)) continue;
    const place = places.get(slot.name);
    const same = place && place.kind === slot.kind && (place.kind !== 'image' || slot.kind !== 'image' || place.imageKind === slot.imageKind);
    if (same) draft.slots[slot.name] = value;
    else if (slot.kind === 'image' && illustrates(slot.imageKind)) draft.dropped.push(slot.name);
    else no();
  }
  return draft;
}

/** Puts the content into the target by what each slot is for. */
function byParts(content: Content, target: PatternDef, shape: Anatomy): Draft {
  const draft: Draft = { slots: {}, changes: [], dropped: [] };
  const put = (slot: TextSlotDef | undefined, text: string[] | undefined): void => {
    if (!text) return;
    if (!slot) return no();
    draft.slots[slot.name] = asValue(text);
  };
  put(shape.title, content.title);
  put(shape.subtitle, content.subtitle);
  put(shape.figure, content.figure);
  for (const [name, { value, slot }] of content.others) {
    if (!shape.others.some((candidate) => candidate.name === name && candidate.kind === slot.kind)) no();
    draft.slots[name] = value;
  }
  if (content.picture) {
    if (shape.picture && shape.picture.imageKind === content.picture.kind) draft.slots[shape.picture.name] = content.picture.ref;
    else if (illustrates(content.picture.kind)) draft.dropped.push(content.picture.from);
    else no();
  }

  const row = shape.row;
  const range = itemRange(target);
  const count = content.items.length;
  const itemWord = target.repeat?.item ?? 'item';
  const nameOf = (part: Part, n: number) => (row?.parts.get(part)?.template ?? '').replace('{n}', String(n));
  const oneLine = (item: Item) => [item.heading, item.text, item.figure].every((part) => (part?.length ?? 0) <= 1);

  if (count > 0 && row && range) {
    // Items into items: cards into columns, a row into a grid.
    if (!row.clear || count < range.min || count > range.max) no();
    put(shape.body, content.body);
    draft.items = count;
    // A row of single texts written "heading: text" can fill items that want the two apart.
    const halves = row.parts.has('heading') && row.parts.has('text') ? content.items.map((item) => (!item.heading && !item.figure && item.text?.length === 1 ? split(item.text[0] as string) : undefined)) : [];
    const apart = halves.length > 0 && halves.every((half) => half !== undefined);
    let merged = false;
    content.items.forEach((item, index) => {
      const n = index + 1;
      if (item.figure && item.heading) no();
      const leadPart: TextPart | undefined = item.figure ? 'figure' : item.heading ? 'heading' : undefined;
      const lead = item.figure ?? item.heading;
      if (apart) {
        const [heading, text] = halves[index] as [string, string];
        draft.slots[nameOf('heading', n)] = heading;
        draft.slots[nameOf('text', n)] = text;
      } else if (leadPart && row.parts.has(leadPart)) {
        draft.slots[nameOf(leadPart, n)] = (lead as string[]).join(' ');
        if (item.text) {
          if (!row.parts.has('text')) no();
          draft.slots[nameOf('text', n)] = asValue(item.text);
        }
      } else if (leadPart) {
        // The target's items have one text: the heading leads it.
        if (!row.parts.has('text') || row.parts.has('heading') || row.parts.has('figure') || !oneLine(item)) no();
        draft.slots[nameOf('text', n)] = joined(lead, item.text);
        merged = true;
      } else if (item.text) {
        if (!row.parts.has('text')) no();
        draft.slots[nameOf('text', n)] = asValue(item.text);
      }
      for (const part of ['icon', 'picture'] as const) {
        const place = row.parts.get(part)?.slot;
        const kind = content.itemPictures.find((picture) => picture.part === part)?.kind;
        if (item[part] && place?.kind === 'image' && place.imageKind === kind) draft.slots[nameOf(part, n)] = item[part] as ImageRef;
      }
    });
    for (const picture of content.itemPictures) {
      const place = row.parts.get(picture.part)?.slot;
      if (place?.kind === 'image' && place.imageKind === picture.kind) continue;
      if (!illustrates(picture.kind)) no();
      draft.dropped.push(picture.from);
    }
    if (content.itemName !== itemWord) draft.changes.push(`The ${plural(count, content.itemName)} became ${plural(count, itemWord)}.`);
    if (merged) draft.changes.push(`Each ${itemWord} is written as one text, "heading: text".`);
    if (apart) draft.changes.push(`Each ${itemWord} was split at the colon into a heading and a text.`);
  } else if (count > 0) {
    // Items into a list: one point for each item.
    const body = shape.body;
    if (!body || content.body || !isList(body) || !content.items.every(oneLine)) return no();
    if (content.itemPictures.some((picture) => !illustrates(picture.kind))) no();
    draft.slots[body.name] = content.items.map((item) => joined(item.figure ?? item.heading, item.text));
    draft.dropped.push(...content.itemPictures.map((picture) => picture.from));
    const lead = content.items.some((item) => item.figure) ? 'figure' : 'heading';
    const written = content.items.some((item) => (item.figure ?? item.heading) && item.text) ? `, each written as "${lead}: text"` : '';
    draft.changes.push(`The ${plural(count, content.itemName)} became a list of ${plural(count, 'point')}${written}.`);
  } else if (row && range) {
    // A list into items: one item for each point.
    const points = content.body;
    if (!points || !row.clear || points.length < 2 || points.length < range.min || points.length > range.max || row.parts.has('figure') || !row.parts.has('text')) return no();
    draft.items = points.length;
    if (row.parts.has('heading')) {
      // The items want a heading and a text each: only a list written "heading: text" gives both.
      const halves = points.map(split);
      if (!halves.every((half) => half !== undefined)) no();
      halves.forEach((half, index) => {
        draft.slots[nameOf('heading', index + 1)] = (half as [string, string])[0];
        draft.slots[nameOf('text', index + 1)] = (half as [string, string])[1];
      });
      draft.changes.push(`The ${plural(points.length, 'point')} became ${plural(points.length, itemWord)}, each split at the colon into a heading and a text.`);
    } else {
      points.forEach((point, index) => (draft.slots[nameOf('text', index + 1)] = point));
      draft.changes.push(`The ${plural(points.length, 'point')} became ${plural(points.length, itemWord)}.`);
    }
  } else {
    put(shape.body, content.body);
    // The words are the same; whether they show as points or as paragraphs is the new slot's own formatting.
    const lines = content.body?.length ?? 0;
    if (lines > 1 && content.bodyIsList !== isList(shape.body)) {
      draft.changes.push(content.bodyIsList ? `The ${plural(lines, 'point')} of the list are written as ${plural(lines, 'paragraph')}.` : `The ${plural(lines, 'paragraph')} became a list of ${plural(lines, 'point')}.`);
    }
  }
  return draft;
}

/**
 * Lists the other patterns that can hold what one slide of a deck says, with the slide written
 * for each. `slide` is the slide's number in the deck, from 1. No model is asked and no text is
 * rewritten: a title stays a title, and a row of items becomes other items or a list of points.
 *
 * A pattern is an alternative when every text, table, chart and supplied graphic of the slide
 * has a place in it and every slot it requires can be filled. A photo or an icon may be left
 * out, which `dropped` reports. A text that is longer than its new slot holds is reported in
 * `warnings`, as `validateDeck` would. Alternatives that fit without a warning come first.
 */
export function slideAlternatives(manifest: TemplateManifest, deck: DeckSpec, slide: number, options: AlternativesOptions = {}): SlideAlternative[] {
  const original = deck.slides[slide - 1];
  if (!original) throw new RangeError(`the deck has ${deck.slides.length} slides; there is no slide ${slide}`);
  // A slide the software lays out itself is its own layout: its blocks say what it shows.
  if (Array.isArray(original.blocks)) return [];
  const source = manifest.patterns.find((pattern) => pattern.id === original.pattern);
  if (!source) throw new Error(`slide ${slide} uses the pattern "${original.pattern}", which the manifest does not have`);
  const layout = options.layout ?? 'strict';
  const filled: SlideSpec = { ...original, slots: original.slots ?? {} };
  const content = contentOf(source, anatomy(source), filled);

  const found: SlideAlternative[] = [];
  for (const target of manifest.patterns) {
    if (target === source || (options.patterns && !options.patterns.includes(target.id))) continue;
    // Slot for slot where the names agree, which is exact; by what the slots are for where they do not.
    for (const attempt of [() => byName(source, target, filled), () => byParts(content, target, anatomy(target))]) {
      let draft: Draft;
      try {
        draft = attempt();
      } catch (error) {
        if (!(error instanceof NoPlace)) throw error;
        continue;
      }
      const written: SlideSpec = { pattern: target.id, slots: draft.slots };
      if (draft.items !== undefined) written.items = draft.items;
      // The renderer's own rules decide whether the slide can be made: a required slot left empty is an error.
      const issues = validateDeck(manifest, { ...deck, slides: [written] }, { layout });
      if (issues.some((issue) => issue.severity === 'error')) continue;
      found.push({ pattern: target.id, description: target.description, slide: written, changes: draft.changes, dropped: draft.dropped, warnings: issues.map((issue) => ({ ...issue, slide })) });
      break;
    }
  }
  const rank = (alternative: SlideAlternative) => (alternative.warnings.length > 0 ? 2 : alternative.dropped.length > 0 ? 1 : 0);
  return found
    .map((alternative, index) => ({ alternative, index }))
    .sort((a, b) => rank(a.alternative) - rank(b.alternative) || a.index - b.index)
    .map((entry) => entry.alternative);
}
