// A free slide when a deck is changed: how the deck lists its parts for the model that plans the
// change, how that plan names the parts the slide has afterwards, and what each of them keeps of
// the slide as it was. A part that stays is carried over by code, under its new number, and no
// model is shown its content or asked for it again: the numbers of a chart cannot drift because
// another chart was added beside it.
import { FREE_PATTERN, isChartValue, isImageRef, isTableValue, toParagraphs } from '../deck/spec.ts';
import type { FreeArrange, FreeBlockKind, SlideSpec, SlotValue } from '../deck/spec.ts';
import { FREE_PARTS, freeSlotName, isFreeBlockKind } from '../free/blocks.ts';
import type { FreePart } from '../free/blocks.ts';
import type { JsonSchema } from '../llm/provider.ts';
import type { PatternDef } from '../template/manifest.ts';
import { SHOWS_MAIN, SHOWS_MAIN_NOTE } from './catalog.ts';
import type { PlannedSlot, UsablePattern } from './catalog.ts';
import { MOST_PARTS, MOST_PHOTOS, describeEntry } from './free.ts';
import { partsWords } from './parts.ts';

// --- the slide as the model that plans a change sees it ---

/** A slide of more parts than this is listed in short: the first three and the last of each run of like parts. */
const MOST_LISTED = 24;
/** A text of a part is shown up to this many characters. */
const TEXT_SHOWN = 160;

/** A text as the model that plans a change is shown it: cut short where it is longer than {@link TEXT_SHOWN}. */
export const cut = (text: string): string => (text.length <= TEXT_SHOWN ? text : `${text.slice(0, TEXT_SHOWN - 3).trimEnd()}...`);
const some = (count: number, one: string, many: string): string => `${count} ${count === 1 ? one : many}`;
const isPicture = (part: FreePart): boolean => part.holds === 'icon' || part.holds === 'photo' || part.holds === 'graphic';
const isText = (part: FreePart): boolean => part.holds === 'text' || part.holds === 'paragraphs';

/** The kinds of a slide's parts, as the slide states them. */
const kindsOf = (slide: Pick<SlideSpec, 'blocks'>): FreeBlockKind[] => (slide.blocks ?? []) as FreeBlockKind[];

/** A picture as a writer is shown it: what it was made from, or else the id it has. */
export const shownValue = (value: SlotValue): unknown => (isImageRef(value) ? (value.prompt ? { prompt: value.prompt, alt: value.alt ?? '' } : value.asset) : value);

/**
 * What one value of a part is, in a few words and without its data: a text in quotes, cut short;
 * a chart by its type and how many categories it has; a table by its size; a picture by what it
 * shows. `undefined` for a part that has no value. `named` says a caption was said before it: a
 * chart without one is told by the names of its series, which are no data either.
 */
function said(part: FreePart, value: SlotValue | undefined, named = true): string | undefined {
  if (value === undefined || value === null) return undefined;
  if (isText(part)) {
    if (typeof value !== 'string' && !Array.isArray(value)) return undefined;
    const text = toParagraphs(value).filter(Boolean).join(part.holds === 'paragraphs' ? ' / ' : ' ');
    return text ? JSON.stringify(cut(text)) : undefined;
  }
  if (part.holds === 'chart') {
    if (!isChartValue(value)) return undefined;
    const names = value.series.map((series) => String(series?.name ?? '').trim()).filter(Boolean).join(', ');
    const series = !named && names ? `, series ${JSON.stringify(cut(names))}` : value.series.length > 1 ? `, ${value.series.length} series` : '';
    return `(${value.type}, ${some(value.categories.length, 'category', 'categories')}${series})`;
  }
  if (part.holds === 'table') return isTableValue(value) ? `(${some(value.header.length, 'column', 'columns')}, ${some(value.rows.length, 'row', 'rows')})` : undefined;
  if (!isImageRef(value)) return undefined;
  if (part.holds === 'icon') return `(icon ${JSON.stringify(value.asset)})`;
  return value.prompt ? `of ${JSON.stringify(cut(value.prompt))}` : JSON.stringify(value.asset);
}

/**
 * The parts of a kind of block. A stored deck may name kinds that are none, among them "constructor"
 * and "toString", which every object has: such a block has no parts, and says only its kind.
 */
const partsOf = (kind: unknown): readonly FreePart[] => (typeof kind === 'string' && Object.hasOwn(FREE_PARTS, kind) ? FREE_PARTS[kind as FreeBlockKind] : []);

/** One part in a line: its number, its kind and what it holds: `3 chart "North" (column, 4 categories)`. */
function partLine(number: number, kind: FreeBlockKind, slots: Record<string, SlotValue>): string {
  let line = `${number} ${kind}`;
  let afterText = false;
  for (const part of partsOf(kind)) {
    const piece = said(part, slots[freeSlotName(number, part.name)], afterText);
    if (!piece) continue;
    // A heading and its text read as one: "Hours": "From five."
    line += `${afterText && isText(part) ? ':' : ''} ${piece}`;
    afterText = isText(part);
  }
  return line;
}

/**
 * The parts of a free slide, one to a line, by their numbers. `numbers` names the parts to list,
 * in order; left out, all of them. Many parts are listed in short: of each run of like parts
 * that follow one another, the first three and the last, with a line between that says how many
 * more there are.
 */
export function listParts(blocks: readonly FreeBlockKind[], slots: Record<string, SlotValue>, numbers: readonly number[] = blocks.map((_, index) => index + 1)): string[] {
  const kind = (number: number) => blocks[number - 1] as FreeBlockKind;
  const line = (number: number) => partLine(number, kind(number), slots);
  if (numbers.length <= MOST_LISTED) return numbers.map(line);
  const lines: string[] = [];
  for (let at = 0; at < numbers.length; ) {
    let end = at;
    while (end + 1 < numbers.length && numbers[end + 1] === (numbers[end] as number) + 1 && kind(numbers[end + 1] as number) === kind(numbers[at] as number)) end++;
    const run = numbers.slice(at, end + 1);
    // Hiding one part saves nothing: a run is cut where two or more can be left out.
    if (run.length <= 5) lines.push(...run.map(line));
    else {
      const hidden = run.slice(3, -1);
      lines.push(...run.slice(0, 3).map(line), `... ${partsWords(hidden)}: ${hidden.length} more ${kind(run[0] as number)}s`, line(run.at(-1) as number));
    }
    at = end + 1;
  }
  return lines;
}

/**
 * A free slide as the model that plans a change is shown it: what it is made of, how it is set
 * out, its title, and its parts by number. No data: a chart is its caption, its type and the
 * number of its categories. `more` is said of the slide after how it is set out, such as the words
 * of a text set dense and the most it holds.
 */
export function listFree(slide: SlideSpec, more: readonly string[] = []): string {
  const blocks = kindsOf(slide);
  const title = said({ name: 'title', holds: 'text', required: false }, slide.slots.title);
  const head = `[${FREE_PATTERN}: ${[describeEntry({ blocks, ...(slide.arrange ? { arrange: slide.arrange } : {}) }), ...more].join('; ')}]`;
  return `${head} ${[...(title ? [`title: ${title}`] : []), ...listParts(blocks, slide.slots)].join('; ')}`;
}

// --- what the plan of a change says of a free slide ---

/** One part of a free slide after a change, as the plan of the change names it. */
export interface PartPlan {
  /** The number of the part of the slide as it was that this part is. 0 for a new part. */
  from: number;
  kind: FreeBlockKind;
  /** Whether the content of the part is written again. */
  rewrite: boolean;
}

/** What the plan of a change says of one free slide. */
export interface FreeChange {
  /** Every part of the slide after the change, in order. Empty when the parts stay as they are. */
  parts: PartPlan[];
  /** Where the main part stands, as the answer has it. What `auto` says depends on the slide (see {@link statesArrange}). */
  lead?: (typeof SHOWS_MAIN)[number];
  /** How many parts stand side by side, as the answer has it. What 0 says depends on the slide too. */
  columns?: number;
}

/** Room in an answer for each part a change may name: where it comes from, its kind and a word. A ceiling. */
const PART_ROOM = 16;

/**
 * Room for the parts the plan of a change may name: the parts the deck's free slides have, twice
 * over, and those the request may add. How many it adds is not known before it is planned, so a
 * number the request names is taken as a guess of it.
 */
export function partsRoom(parts: number, request: string): number {
  const named = [...request.matchAll(/\d+/g)].map((match) => Number(match[0])).filter((number) => number <= MOST_PARTS);
  const asked = Math.min(MOST_PARTS, named.reduce((sum, number) => sum + number, 0));
  return PART_ROOM * (2 * parts + Math.max(40, asked));
}

/** The rule the plan of a change is given when the deck may hold free slides. */
export const PARTS_RULE = 'parts: for a slide with the layout free whose parts change. Leave it [] otherwise.';

/**
 * Said with it. Parts that are taken off or moved are set by code, with nothing written. Words
 * in "change" on such a slide are taken to ask for something to be written, and its title is
 * then offered to the writer: the plan has no other way to name the title.
 */
export const MOVED_RULE = 'A slide with the layout free whose parts are only taken off or moved needs nothing written: leave its change empty unless something on the slide is to be written too, such as another title.';

/**
 * The rule given where a slide of a pattern may become a free slide. What such a slide holds is
 * handed to its first parts of that kind by code (see {@link carryOver}), unless the plan asks
 * for the part to be written again: the plan is told, since a new part is otherwise marked so.
 */
export const TURN_RULE =
  'A slide of a pattern that takes the layout free hands its chart, its table and its photos to the first parts of those kinds that it names. Give such a part rewrite false unless the request changes what it shows.';

/**
 * What an entry of the plan of a change says of a free slide: its parts one by one, and how they
 * are set out. The three properties are asked of every slide, as in an outline. `kinds` is never
 * empty.
 */
export function partsProperties(kinds: readonly FreeBlockKind[]): Record<string, JsonSchema> {
  const repeat = 'Repeat the current value unless the request is about the arrangement.';
  return {
    parts: {
      type: 'array',
      description:
        'Only for the layout free, and only when the request changes which parts the slide has, their order, or the content of some of them: every part of the slide after the change, in order. ' +
        "from: the number of the existing part it is, 0 for a new part. rewrite: true when the request changes that part's content. " +
        '[] when the parts stay as they are, and for every other layout.',
      items: {
        type: 'object',
        properties: { from: { type: 'integer', minimum: 0 }, kind: { type: 'string', enum: [...kinds] }, rewrite: { type: 'boolean' } },
        required: ['from', 'kind', 'rewrite'],
      },
    },
    shows_main: { type: 'string', enum: [...SHOWS_MAIN], description: `${SHOWS_MAIN_NOTE} ${repeat}` },
    side_by_side: { type: 'integer', minimum: 0, description: `Only for free, when the request asks for a number of columns; 0 lets the software decide. ${repeat}` },
  };
}

/**
 * Reads what an entry says of a free slide. A part of a kind that is not on offer is passed over,
 * as in an outline, unless it names a part the slide has (`before`, the parts of the slide the
 * entry is based on): that part is meant, whatever was written for its kind, and it keeps the
 * kind it has. Passing it over would take a part off the slide that nobody asked to remove.
 */
export function readParts(raw: unknown, kinds: readonly FreeBlockKind[], before?: readonly FreeBlockKind[]): FreeChange {
  const given = (typeof raw === 'object' && raw !== null ? raw : {}) as { parts?: unknown; shows_main?: unknown; side_by_side?: unknown };
  const parts = (Array.isArray(given.parts) ? given.parts : [])
    .filter((part): part is { from?: unknown; kind?: unknown; rewrite?: unknown } => typeof part === 'object' && part !== null)
    .flatMap((part): PartPlan[] => {
      const from = Math.floor(Number(part.from));
      const named = isFreeBlockKind(part.kind) && kinds.includes(part.kind) ? part.kind : undefined;
      const kind = named ?? (from >= 1 ? before?.[from - 1] : undefined);
      return kind && isFreeBlockKind(kind) ? [{ from: from >= 1 ? from : 0, kind, rewrite: part.rewrite === true }] : [];
    });
  const change: FreeChange = { parts };
  const lead = SHOWS_MAIN.find((value) => value === given.shows_main);
  if (lead) change.lead = lead;
  // Only a number says something: a property that was left out leaves the arrangement as it is.
  if (typeof given.side_by_side === 'number' && Number.isInteger(given.side_by_side) && given.side_by_side >= 0) change.columns = given.side_by_side;
  return change;
}

/**
 * Checks the parts a plan names for one slide against the slide it is based on. `before` is the
 * parts that slide has when it is a free slide already; `photos` the most photos a slide may
 * have made for it.
 *
 * What is wrong is a `problem` to send back once. With `final`, the last answer there will be,
 * it is put right here instead: a part that points at none is a new part, a slide that turns
 * free without parts is one text, and a slide with too many parts or photos is cut. A plan that
 * leaves most of a slide's parts unnamed is asked about once, and then done as it says. No
 * answer makes a change fail over the parts of a free slide.
 */
export function settleParts(change: FreeChange, slide: number, before: readonly FreeBlockKind[] | undefined, photos: number, final: boolean): { parts: PartPlan[]; problem?: string } {
  let parts = change.parts;
  let problem: string | undefined;
  const has = before?.length ?? 0;

  const astray = parts.find((part) => part.from > has);
  if (astray && !final) {
    problem = has > 0 ? `slide ${slide}: "from" of a part must be 0 or the number of an existing part (1 to ${has}), not ${astray.from}` : `slide ${slide}: the slide it is based on has no numbered parts, so "from" of each of its parts must be 0`;
  } else if (astray) {
    // A part that points at none is a new part. On a free slide it is written. On a slide that
    // had no numbered parts, "from 1, not to be written again" can only have meant what the slide
    // holds, its chart say: the part keeps the answer it gave, which decides what is handed to it
    // (see carryOver).
    parts = parts.map((part) => (part.from > has ? { ...part, from: 0, rewrite: has > 0 ? true : part.rewrite } : part));
  }

  if (!before && parts.length === 0) {
    if (final) parts = [{ from: 0, kind: 'text', rewrite: true }];
    else problem = `slide ${slide}: a slide that becomes a free slide needs "parts", one entry for each part it shows`;
  }

  // A part the plan does not name is taken off the slide. A plan that names the parts that change
  // and forgets the others would take most of a slide away, so that is asked about once. The same
  // answer again is taken at its word.
  const named = new Set(parts.filter((part) => part.from > 0).map((part) => part.from));
  const dropped = has - named.size;
  if (!final && !problem && parts.length > 0 && dropped >= 2 && dropped > has / 2) {
    problem = `slide ${slide}: ${dropped} of the ${has} parts of the slide are not named in "parts" and would be taken off it; name every part that stays, also those that do not change. If they are to go, answer the same again`;
  }

  // A slide the deck holds may have more parts than a model may name: it keeps them, and gets no more.
  const most = Math.max(MOST_PARTS, has);
  if (parts.length > most) {
    if (final) parts = parts.slice(0, most);
    else problem = `slide ${slide}: a free slide takes at most ${MOST_PARTS} parts, not ${parts.length}`;
  }

  const count = (list: readonly { kind: FreeBlockKind }[]) => list.filter((part) => part.kind === 'photo').length;
  const mostPhotos = Math.max(photos, count((before ?? []).map((kind) => ({ kind }))));
  if (count(parts) > mostPhotos) {
    if (final) {
      // The photos that are new go first, from the end; then those the slide had.
      let over = count(parts) - mostPhotos;
      for (const fresh of [true, false]) {
        for (let at = parts.length - 1; at >= 0 && over > 0; at--) {
          const part = parts[at] as PartPlan;
          if (part.kind !== 'photo' || (part.from === 0) !== fresh) continue;
          parts = parts.toSpliced(at, 1);
          over--;
        }
      }
    } else {
      problem = `slide ${slide}: at most ${MOST_PHOTOS} photos can be made for one slide; supply pictures for more`;
    }
  }
  return problem ? { parts, problem } : { parts };
}

/** Whether a plan names the parts of a slide exactly as the slide has them: nothing added, taken off, moved or written again. */
export function sameParts(parts: readonly PartPlan[], before: readonly FreeBlockKind[]): boolean {
  return parts.length === 0 || (parts.length === before.length && parts.every((part, index) => part.from === index + 1 && part.kind === before[index] && !part.rewrite));
}

/**
 * Whether a plan states how a slide is set out. A place for the main part (`left`, `right`,
 * `top`, `none`) and a number side by side of 1 or more always do. `auto` and 0 hand the slide
 * back to the software, and do so only on a slide the plan changes (`changed`: its `change` has
 * words, or its parts change). On a slide the plan leaves alone they are what an answer gives
 * for a slide it has nothing to say of, and state nothing: such a slide is copied as it is. A
 * property that was left out states nothing on any slide.
 */
export const statesArrange = (change: Pick<FreeChange, 'lead' | 'columns'>, changed: boolean): boolean =>
  (change.lead !== undefined && (changed || change.lead !== 'auto')) || (change.columns !== undefined && (changed || change.columns !== 0));

/**
 * How a slide is set out after a change: as the plan states it, and as it was in what the plan
 * does not state (see {@link statesArrange}). On a slide the plan changes (`changed`), `auto`
 * takes the place of the main part away from the slide and 0 its number side by side: the
 * software decides again. No more parts stand side by side than the slide has. `undefined` when
 * nothing is left stated: the slide then has no `arrange`.
 */
export function arrangeAfter(before: FreeArrange | undefined, change: Pick<FreeChange, 'lead' | 'columns'>, count: number, changed: boolean): FreeArrange | undefined {
  const arrange: FreeArrange = { ...before };
  if (change.lead === 'auto') {
    if (changed) delete arrange.lead;
  } else if (change.lead) arrange.lead = change.lead;
  if (change.columns === 0) {
    if (changed) delete arrange.columns;
  } else if (change.columns) arrange.columns = change.columns;
  if (typeof arrange.columns === 'number' && Number.isInteger(arrange.columns) && arrange.columns > count) arrange.columns = count;
  return Object.keys(arrange).length > 0 ? arrange : undefined;
}

/** Whether two slides are set out alike, as far as they state it. */
export const sameArrange = (a: FreeArrange | undefined, b: FreeArrange | undefined): boolean => a?.columns === b?.columns && a?.lead === b?.lead;

// --- what the slide keeps ---

/** A free slide after a change, before anything is written for it. */
export interface Reshaped {
  blocks: FreeBlockKind[];
  /** The values the slide keeps as they are, under the names they have after the change. */
  carried: Record<string, SlotValue>;
  /** The slots that are not asked of the writer, whether the slide has a value for them or not. */
  skip: Set<string>;
  /** What the slots that are written again hold now, under the names they have after the change. */
  current: Record<string, SlotValue>;
  /** For a part that takes the place of a part of another kind: what that one held, by the number of the new part. */
  replaced: Record<string, Record<string, unknown>>;
  /** The parts of which something is written, by their numbers after the change. */
  written: number[];
}

const reshaped = (blocks: FreeBlockKind[]): Reshaped => ({ blocks, carried: {}, skip: new Set(), current: {}, replaced: {}, written: [] });
const filled = (value: SlotValue | undefined): boolean => value !== undefined && value !== null && !(typeof value === 'string' && value.trim() === '') && !(Array.isArray(value) && value.length === 0);

/**
 * Works out what a free slide keeps when its parts change. A part that is the same kind of part
 * as the one it comes from, and is not written again, keeps every value of that one under its
 * new number. One that is written again keeps its pictures, unless the change is about pictures
 * (`images`), and shows the writer what else it holds now. A new part, and one of another kind
 * than the part it comes from, is written from nothing; what the old part held is kept as
 * material for the writer.
 *
 * The slots are found through the names the engine gives them (`freeSlotName`), part by part:
 * every value goes over exactly as the deck has it.
 */
export function carryParts(before: SlideSpec, parts: readonly PartPlan[], images: boolean): Reshaped {
  const old = kindsOf(before);
  const result = reshaped(parts.map((part) => part.kind));
  // The like parts that stay whole, and which of their optional values some of them has.
  const stays = new Map<FreeBlockKind, Set<string>>();

  parts.forEach((part, index) => {
    const number = index + 1;
    const same = part.from > 0 && old[part.from - 1] === part.kind;
    if (!same || part.rewrite) result.written.push(number);
    if (part.from === 0) return;
    if (!same) {
      const held = partsOf(old[part.from - 1]).flatMap((entry) => {
        const value = before.slots[freeSlotName(part.from, entry.name)];
        return value === undefined || value === null ? [] : [[entry.name.toLowerCase(), shownValue(value)] as const];
      });
      if (held.length > 0) result.replaced[String(number)] = Object.fromEntries(held);
      return;
    }
    const had = stays.get(part.kind) ?? new Set<string>();
    if (!part.rewrite) stays.set(part.kind, had);
    for (const entry of FREE_PARTS[part.kind]) {
      const name = freeSlotName(number, entry.name);
      const value = before.slots[freeSlotName(part.from, entry.name)];
      if (!part.rewrite || (isPicture(entry) && !images)) {
        result.skip.add(name);
        if (value !== undefined) result.carried[name] = structuredClone(value);
        if (!part.rewrite && filled(value)) had.add(entry.name);
      } else if (value !== undefined && value !== null) {
        result.current[name] = value;
      }
    }
  });

  // A new part beside like parts that stay takes after them in its pictures: an icon is drawn
  // when every card of a slide has one, so a new card is asked for none where the others have none.
  parts.forEach((part, index) => {
    const had = stays.get(part.kind);
    if (!had || (part.from > 0 && old[part.from - 1] === part.kind)) return;
    for (const entry of FREE_PARTS[part.kind]) if (!entry.required && isPicture(entry) && !had.has(entry.name)) result.skip.add(freeSlotName(index + 1, entry.name));
  });
  return result;
}

/**
 * The same for a slide that becomes a free slide, or is new: every part is written. A slide of a
 * pattern hands its chart, its table and its photos to the first new parts of those kinds, in
 * order, so that their numbers are not copied by a model and its pictures are not made again. A
 * chart or a table goes to a part the plan does not ask to be written again; a photo to any
 * photo part, unless the change is about pictures and the part is to be written again.
 * `moved` gets the names of the slots, in the slide as it was, whose values were handed on.
 */
export function carryOver(source: SlideSpec | undefined, definition: PatternDef | undefined, parts: readonly PartPlan[], images: boolean, moved: Set<string> = new Set()): Reshaped {
  const result = reshaped(parts.map((part) => part.kind));
  result.written = parts.map((_, index) => index + 1);
  const had = (kind: 'chart' | 'table' | 'photo') =>
    (definition?.slots ?? []).flatMap((slot) => {
      const value = source?.slots[slot.name];
      const fits = kind === 'chart' ? slot.kind === 'chart' && isChartValue(value) : kind === 'table' ? slot.kind === 'table' && isTableValue(value) : slot.kind === 'image' && slot.imageKind === 'photo' && isImageRef(value);
      return fits ? [{ name: slot.name, value: value as SlotValue }] : [];
    });
  const left = { chart: had('chart'), table: had('table'), photo: had('photo') };
  parts.forEach((part, index) => {
    if (part.kind !== 'chart' && part.kind !== 'table' && part.kind !== 'photo') return;
    if (part.kind === 'photo' ? images && part.rewrite : part.rewrite) return;
    const next = left[part.kind].shift();
    const entry = FREE_PARTS[part.kind].find((candidate) => candidate.holds === part.kind);
    if (!next || !entry) return;
    const name = freeSlotName(index + 1, entry.name);
    result.carried[name] = structuredClone(next.value);
    result.skip.add(name);
    moved.add(next.name);
  });
  return result;
}

/**
 * The slots that hold a value on a free slide after a change, under the names they then have:
 * those it keeps as they are and those that are written again. The slide is planned with them
 * (`withParts` of `freeWritePattern`), so that an optional part the software would not plan by its
 * rules is counted in where the slide has it: the caption under its one photo, the heading over
 * its one text. Such a part then has its room on the slide, and is offered again when it is to
 * be written again. The title is no part.
 */
export function heldParts(shape: Pick<Reshaped, 'carried' | 'current'>): string[] {
  return [...Object.entries(shape.carried), ...Object.entries(shape.current)].filter(([name, value]) => name !== 'title' && filled(value)).map(([name]) => name);
}

/**
 * A pattern to write against, with the limits of each chart and table no lower than what the
 * part holds now (`current`, by slot name). The limits of a planned chart say what reads well at
 * the size it is drawn at. A chart the slide holds may be larger: it was written when the slide
 * had fewer parts, or by code. The writer is shown it as it is and asked for it at that size. A
 * schema that allowed fewer series than the chart has would take one away in a change that is
 * not about its numbers, another language say.
 */
export function roomAsHeld(pattern: UsablePattern, current: Record<string, SlotValue>): UsablePattern {
  const most = (limit: number | undefined, held: number): number | undefined => (limit === undefined ? undefined : Math.max(limit, held));
  const slots = pattern.slots.map((planned): PlannedSlot => {
    const value = current[planned.slot.name];
    if (planned.mode === 'chart' && isChartValue(value)) {
      const slot = { ...planned.slot };
      const categories = most(slot.maxCategories, value.categories.length);
      const series = most(slot.maxSeries, value.series.length);
      if (categories !== undefined) slot.maxCategories = categories;
      if (series !== undefined) slot.maxSeries = series;
      return { ...planned, slot };
    }
    if (planned.mode === 'table' && isTableValue(value)) {
      const slot = { ...planned.slot };
      const columns = most(slot.maxColumns, value.header.length);
      const rows = most(slot.maxRows, value.rows.length);
      if (columns !== undefined) slot.maxColumns = columns;
      if (rows !== undefined) slot.maxRows = rows;
      return { ...planned, slot };
    }
    return planned;
  });
  return { ...pattern, slots };
}

/** The photos of a free slide, in the order of its parts: what a slide that takes another layout carries over. */
export function freePhotos(slide: SlideSpec): SlotValue[] {
  return kindsOf(slide).flatMap((kind, index) =>
    partsOf(kind)
      .filter((entry) => entry.holds === 'photo')
      .map((entry) => slide.slots[freeSlotName(index + 1, entry.name)])
      .filter(isImageRef),
  );
}
