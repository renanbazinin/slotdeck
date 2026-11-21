// The part of planning that planDeck and planEdit share: asking the model for the content of
// a set of slides, turning its answer into slide specs, and sending drafts back for repair.
import { validateDeck } from '../deck/validate.ts';
import type { Issue, LayoutMode, ValidateOptions } from '../deck/validate.ts';
import { FREE_PATTERN, isChartValue, isFreeSlide, isImageRef, isTableValue, toParagraphs } from '../deck/spec.ts';
import type { DeckSpec, FreeArrange, FreeBlockKind, ImageRef, SlideSpec, SlotValue } from '../deck/spec.ts';
import { FREE_PARTS, describeBlocks, freeContent, freeSlide, freeSlotName, freeSlots } from '../free/blocks.ts';
import type { FreeBlockContent } from '../free/blocks.ts';
import { blockRuns } from '../free/offer.ts';
import type { JsonSchema } from '../llm/provider.ts';
import { EN, said } from '../messages/en.ts';
import type { Template } from '../render/render.ts';
import { textBudget } from '../template/manifest.ts';
import type { TemplateManifest } from '../template/manifest.ts';
import type { TextMeasurer } from '../text/measure.ts';
import { readImageInfo } from '../util/image.ts';
import { budgetScale, contentSchema } from './catalog.ts';
import type { CatalogContext, UsablePattern, WriteItem } from './catalog.ts';
import { MOST_PHOTOS, partLabel } from './free.ts';
import { REPEAT_TITLE, mergeRange, partOf, partRanges, rangeSchema, rangeTask, readRange, titled } from './parts.ts';
import type { FreeJob, PartRange, RangeAnswer } from './parts.ts';

/**
 * The rules every planning and writing call is given. `document` says the deck is made from a
 * document: its text is then said in its own words, and its figures may be used.
 */
export function systemPrompt(language: string, source: 'brief' | 'deck', document = false): string {
  const limits = document
    ? `- Character limits are hard and count spaces. Stay under them. A slide that has text from the document says that text, in its words and order: shorten it only where a limit ${source === 'brief' ? 'makes you' : 'or the request makes you'}, by leaving out what matters least, never by summing it up. Text in another language than the deck's is translated closely, sentence by sentence. Where a slot says how much to write, write about that much; elsewhere keep slide text short: a phrase beats a sentence.`
    : '- Character limits are hard and count spaces. Stay under them. Where a slot says how much to write, write about that much; elsewhere keep slide text short: a phrase beats a sentence.';
  const figures =
    source === 'brief'
      ? `- Every figure on a slide must come from the brief${document ? ' or the document' : ''}. Never make up statistics, quotes or names.`
      : `- Every figure on a slide must come from the request${document ? ', the document or the deck as it is' : ' or from the deck as it is'}. Never make up statistics, quotes or names.`;
  const topics = source === 'brief' ? (document ? '- Follow the brief. Do not add topics that neither the brief nor the document calls for.' : '- Follow the brief. Do not add topics it does not call for.') : '- Do what the request asks and nothing more.';
  return [
    "You write presentation decks for a fixed slide template. You do not place or size anything: you choose from the template's slide patterns and write the content for their slots.",
    '',
    'Rules:',
    `- Write all slide text in the language with BCP 47 tag "${language}".`,
    limits,
    '- Plain text only: no markdown, no bullet characters, no line breaks inside a string, no placeholder text.',
    figures,
    topics,
  ].join('\n');
}

export function cleanText(value: unknown): string {
  return String(value ?? '')
    .replace(/\s+/g, ' ')
    .trim();
}

/** Stable short id for a generated picture, so the same description maps to the same asset across edits. */
export function generatedAssetId(prompt: string): string {
  let a = 0x811c9dc5;
  let b = 0x01000193;
  for (let i = 0; i < prompt.length; i++) {
    const code = prompt.charCodeAt(i);
    a = Math.imul(a ^ code, 0x01000193);
    b = Math.imul(b ^ code, 0x85ebca6b) + 0x9e3779b9;
  }
  const hex = (value: number) => (value >>> 0).toString(16).padStart(8, '0');
  return `gen-${hex(a)}${hex(b)}`;
}

/** Turns the model's answer for one slide into a slide spec, tidying text on the way. */
export function toSlide(pattern: UsablePattern, content: unknown, context: CatalogContext): SlideSpec {
  const values = (typeof content === 'object' && content !== null ? content : {}) as Record<string, unknown>;
  const slots: Record<string, SlotValue> = {};

  for (const planned of pattern.slots) {
    const value = values[planned.slot.name];
    if (value === undefined || value === null) continue;

    if (planned.mode === 'text') {
      slots[planned.slot.name] = cleanText(Array.isArray(value) ? value.join(' ') : value);
    } else if (planned.mode === 'paragraphs') {
      const items = Array.isArray(value) ? value : String(value).split(/\r?\n/);
      slots[planned.slot.name] = items.map((item) => cleanText(item).replace(/^[•\-–]\s+/, '')).filter(Boolean);
    } else if (planned.mode === 'table') {
      const table = (typeof value === 'object' ? value : {}) as { header?: unknown; rows?: unknown };
      const header = (Array.isArray(table.header) ? table.header : []).map(cleanText);
      if (header.length === 0) continue;
      // A row the model made too short or too long is squared up rather than rejected.
      const rows = (Array.isArray(table.rows) ? table.rows : [])
        .filter((row): row is unknown[] => Array.isArray(row))
        .map((row) => header.map((_, index) => cleanText(row[index])));
      slots[planned.slot.name] = { header, rows };
    } else if (planned.mode === 'chart') {
      const chart = (typeof value === 'object' ? value : {}) as { type?: unknown; categories?: unknown; series?: unknown; format?: unknown };
      const type = chart.type === 'bar' || chart.type === 'line' || chart.type === 'pie' ? chart.type : 'column';
      const categories = (Array.isArray(chart.categories) ? chart.categories : []).map(cleanText);
      const series = (Array.isArray(chart.series) ? chart.series : []).map((entry) => {
        const item = (entry ?? {}) as { name?: unknown; values?: unknown };
        // Lengths are left as the model gave them: a short series is a problem to send back, not to pad.
        // A value the model says is not known (null) stays a gap, where Number would make it 0.
        const values = (Array.isArray(item.values) ? item.values : []).map((number: unknown) => (number === null ? null : Number(number)));
        return { name: cleanText(item.name), values: values.filter((number) => number === null || Number.isFinite(number)) };
      });
      if (categories.length === 0 || series.length === 0) continue;
      slots[planned.slot.name] = chart.format === 'percent' ? { type, categories, series, format: 'percent' } : { type, categories, series };
    } else if (planned.mode === 'photo') {
      if (typeof value !== 'object') continue;
      const photo = value as { prompt?: unknown; alt?: unknown; asset?: unknown };
      const alt = cleanText(photo.alt);
      const supplied = context.assets.find((asset) => asset.kind === 'photo' && asset.id === photo.asset);
      const prompt = cleanText(photo.prompt);
      let ref: ImageRef | undefined;
      if (supplied) ref = { asset: supplied.id };
      else if (prompt) ref = { asset: generatedAssetId(prompt), prompt };
      if (!ref) continue;
      if (alt) ref.alt = alt;
      slots[planned.slot.name] = ref;
    } else {
      slots[planned.slot.name] = { asset: cleanText(value) };
    }
  }
  return { pattern: pattern.def.id, slots };
}

/** The icon and picture lists appended to a writing prompt, when the slides being written can use them. */
export function suppliesText(items: WriteItem[], context: CatalogContext): string[] {
  const usesIcons = items.some(({ pattern }) => pattern.slots.some((planned) => planned.mode === 'icon'));
  const usesPictures = items.some(({ pattern }) => pattern.slots.some((planned) => planned.mode === 'photo' || planned.mode === 'graphic'));
  return [
    usesIcons ? `Icons to choose from:\n${context.icons.map((icon) => `- ${icon.id}: ${icon.description}`).join('\n')}` : '',
    usesPictures && context.assets.length > 0
      ? `Supplied pictures. Match a picture's proportions to the slide: a wide picture for a wide area.\n${context.assets
          .map((asset) => `- ${asset.id} (${asset.kind}, ${proportions(asset.data)}): ${asset.description}`)
          .join('\n')}`
      : '',
  ].filter(Boolean);
}

/** A picture's shape in words the model can match to a slide, e.g. "wide, 3.2:1". */
function proportions(data: Uint8Array): string {
  try {
    const { width, height } = readImageInfo(data);
    const ratio = width / height;
    const shape = ratio >= 2 ? 'wide' : ratio > 1.15 ? 'landscape' : ratio >= 0.87 ? 'square' : 'portrait';
    return `${shape}, ${ratio >= 1 ? `${Math.round(ratio * 10) / 10}:1` : `1:${Math.round((1 / ratio) * 10) / 10}`}`;
  } catch {
    return 'unknown size';
  }
}

/** Ids the renderer will be able to resolve: supplied icons and assets, plus pictures that carry a description. */
function knownAssets(deck: DeckSpec, context: CatalogContext): Set<string> {
  const known = new Set<string>([...context.icons.map((icon) => icon.id), ...context.assets.map((asset) => asset.id)]);
  for (const slide of deck.slides) {
    for (const value of Object.values(slide.slots)) {
      if (isImageRef(value) && value.prompt) known.add(value.asset);
    }
  }
  return known;
}

export interface WriteJob extends WriteItem {
  /** Slot values taken over unchanged from an existing slide. They are not offered to the model. */
  carried?: Record<string, SlotValue>;
  /**
   * For a free slide: its parts, how they are set out and the layout whose frame it stands on.
   * They go to the written slide as they are. `pattern` is then the pattern worked out for these
   * parts (see `freeWritePattern`), whose slots are `b1Heading`, `b2Chart` and so on.
   */
  free?: { blocks: FreeBlockKind[]; arrange?: FreeArrange; base: string };
  /**
   * For a free slide that is being changed: what the slots to write hold now, by slot name. A
   * call that writes parts of the slide by themselves shows it in the shape it answers in.
   */
  current?: Record<string, SlotValue>;
  /**
   * For a free slide that is being changed: the slots that were named outright to be written
   * again, by the plan of the change or by a change of language. One of them that comes back
   * with nothing is asked for again like a missing slot, and when it still has not come the deck
   * is told that the change was not made. A slot that is merely offered with what it holds
   * (the plan gave words for the slide and named no part) may be left as it is without a word.
   */
  due?: ReadonlySet<string>;
  /**
   * The purpose says what to do with the slide, not what the slide says: the job is a change to
   * a deck. It decides what a free slide shows when none of its parts has content (see `fallback`).
   */
  task?: boolean;
  /**
   * For a free slide: a text of the deck, such as its title, that the slide can say when none of
   * its parts has content and it has no words of its own. Such a slide says its purpose, or else
   * the title that was written for it; where the purpose is a task, the title comes first and
   * the purpose last, since a slide should not show an instruction.
   */
  fallback?: string;
  /** For a slide that carries passages of the deck's document: what they are called ("p4-p6") and their text, as the writer is shown it. */
  passage?: { label: string; text: string };
  /**
   * The running slot that code set to the slide's passages, word for word (in `carried`). The
   * writer does not write it, and what the check of the deck says of it is never sent back to the
   * writer: it stays with the deck as it is.
   */
  filled?: string;
}

/**
 * A free slide of up to this many parts is written with the other slides, a property for each
 * slot. Their numbers have one digit, so the order of the properties is the order of the parts
 * also for a model that sorts them.
 */
export const FLAT_PARTS = 9;

/** Whether a slide is written by calls of its own, as arrays of like parts: a free slide of more than {@link FLAT_PARTS} parts. */
export const writtenApart = (job: Pick<WriteJob, 'free'>): boolean => (job.free?.blocks.length ?? 0) > FLAT_PARTS;

/** The parts of a free slide for the one who writes its slots: "b1 to b6: 6 charts; b7: a note". */
export function describeParts(blocks: readonly FreeBlockKind[]): string {
  return blockRuns(blocks)
    .map((run) => `${run.number === 1 ? `b${run.first}` : `b${run.first} to b${run.first + run.number - 1}`}: ${describeBlocks(Array<FreeBlockKind>(run.number).fill(run.kind))}`)
    .join('; ');
}

export interface WriteArgs {
  ask: (prompt: string, schema: JsonSchema, room?: number) => Promise<unknown>;
  /** The same for a call that only sets numbers down in charts and tables: it is asked with little reasoning. Defaults to `ask`. */
  askLight?: (prompt: string, schema: JsonSchema, room?: number) => Promise<unknown>;
  /** The writing prompt, for the slides that are written together. Repair rounds append the previous answer and its problems. */
  request: string;
  /**
   * The prompt of a call that writes some parts of one slide by themselves (see
   * {@link writtenApart}), built around the lines that say which parts. Left out, those lines
   * are the whole prompt.
   */
  requestFor?: (job: WriteJob, task: string) => string;
  jobs: WriteJob[];
  context: CatalogContext;
  /** Builds the whole deck from the freshly written slides, keyed by slide number. */
  assemble: (written: Map<number, SlideSpec>) => DeckSpec;
  manifest: TemplateManifest;
  /**
   * The template as it was loaded, before the deck's own layouts were built onto it. A deck that
   * has a free slide is checked with it, so that the slide is laid out and its text checked
   * against the room it gets. Without it a free slide is checked for what needs no template.
   */
  template?: Template | undefined;
  /** Measures text with the template's fonts, so drafts are checked and repaired by the lines they take. */
  measure?: TextMeasurer | undefined;
  /** A second look at each draft that renders: what it finds goes back to the model with the limits the draft broke. */
  review?: ((deck: DeckSpec) => Issue[] | Promise<Issue[]>) | undefined;
  /**
   * What the caller finds wrong with each draft besides the check of the deck, such as a text
   * written well short of the length asked for it. It goes back to the model like any problem of
   * a slide that was written. `round` is 0 for the first answer and counts the repairs after it.
   */
  notes?: ((deck: DeckSpec, issues: Issue[], round: number) => Issue[]) | undefined;
  layout: LayoutMode;
  maxRepairs: number;
  onRepair?: (attempt: number, issues: number) => void;
}

/**
 * How many tokens the content of some slides may take: every slot at its budget, with room to
 * spare. It is a ceiling for the model's answer, so that an answer that runs away is cut short.
 */
export function contentRoom(jobs: Array<{ pattern: UsablePattern }>, context: CatalogContext): number {
  let characters = 0;
  let slots = 0;
  for (const { pattern } of jobs) {
    for (const planned of pattern.slots) {
      slots++;
      const { slot } = planned;
      if (slot.kind === 'text') characters += textBudget(slot, budgetScale(context)).maxChars ?? 400;
      else if (slot.kind === 'table') characters += (slot.maxColumns ?? 4) * ((slot.maxRows ?? 4) + 1) * (slot.maxCellChars ?? 30);
      else if (slot.kind === 'chart') characters += 400;
      else characters += planned.mode === 'photo' ? 500 : 60; // a photo is described in a sentence or two
    }
  }
  // A token is rarely shorter than a character, even in scripts that tokenise badly; names and punctuation come on top.
  return Math.ceil((characters + slots * 40) * 1.5) + 500;
}

// --- free slides: what is written for their parts, and what becomes of a part without content ---

const isFreeJob = (job: WriteJob): job is FreeJob => job.free !== undefined;

/** The part of a block that holds its numbers. A brief that gives none leaves it empty, and the block is then not shown. */
const NUMBERS: Partial<Record<FreeBlockKind, string>> = { figure: 'Value', chart: 'Chart', table: 'Table' };

/**
 * Why a planned part is not on the slide: nothing was written for it, what was written cannot be
 * shown, or it is a photo to generate beyond the number that are made for one slide.
 */
type Lost = 'empty' | 'unusable' | 'surplus';

/** What the check of the last draft holds against a free slide for good: parts to take off, and optional values to drop. */
interface TakenOff {
  blocks: Map<number, Lost>;
  slots: Set<string>;
  /** Where a slide is being changed: the slots that get the value they had back, since what was written for them cannot be shown. */
  back: Set<string>;
}

const nothingOff = (): TakenOff => ({ blocks: new Map(), slots: new Set(), back: new Set() });

/** A free slide as a draft holds it, with what it took to get there. */
interface FreeDraft {
  slide: SlideSpec;
  /** For each part of the slide, in order: the number it had when the slide was planned. */
  kept: number[];
  /** How many planned parts are not on the slide, and how many were to be written again and are as they were. */
  left: Record<Lost | 'unchanged', number>;
  /**
   * Where a slide is being changed: the slots that were asked for outright, came back with
   * nothing and hold what they had, under the names they were planned with.
   */
  unwritten: string[];
}

/** A table without a row shows nothing: it counts as left empty, like a chart without categories or without a series. */
const filled = (value: SlotValue | undefined): boolean =>
  value !== undefined &&
  value !== null &&
  !(typeof value === 'string' && value.trim() === '') &&
  !(Array.isArray(value) && value.length === 0) &&
  !(isTableValue(value) && value.rows.length === 0) &&
  !(isChartValue(value) && (value.categories.length === 0 || value.series.length === 0 || value.series.every((entry) => !Array.isArray(entry?.values) || entry.values.length === 0)));

const REQUIRED = 'this slot is required';
/** Said of a slot that a change was to write again and that came back with nothing. It is sent back like any problem. */
const UNWRITTEN = EN['part-unwritten']({});
/** Said of a slide for which more photos were described than are made for one slide. */
const PHOTOS_OVER = EN['photos-over']({ most: MOST_PHOTOS });

/**
 * The smallest budget that text is held to. A planned slide is set out for texts of the usual
 * length, and each slot says what its box holds then: a budget is one to write to, and no text of
 * a line or more is told less than this many characters. What states less is a word or two (the
 * value of a figure, a heading on a narrow card), where a quarter over is a character or two.
 * That is left to the check of the deck, which reports what does not fit.
 */
const LEAST_BUDGET = 12;

const oneLine = (value: unknown): string => (typeof value === 'string' ? cleanText(value) : '');

/** The photos of a drafted slide that the writer described, to be generated: their part numbers on the slide as planned, in order. */
function describedPhotos(job: FreeJob, draft: FreeDraft): number[] {
  const numbers: number[] = [];
  for (const planned of job.pattern.slots) {
    const at = planned.mode === 'photo' ? partOf(planned.slot.name) : undefined;
    const now = at ? draft.kept.indexOf(at.block) + 1 : 0;
    if (!at || now === 0) continue;
    const value = draft.slide.slots[freeSlotName(now, at.part)];
    if (isImageRef(value) && value.prompt) numbers.push(at.block);
  }
  return numbers.sort((a, b) => a - b);
}

/**
 * Makes the slide from what was written for its parts, `written` holding the slots under the
 * numbers the parts were planned with. A chart, a table or a figure that came back empty is taken
 * off: the brief had no numbers for it. So is whatever `off` names. The parts that stay are
 * numbered anew, and a slide left with none says one text: its purpose, or else the title that
 * was written for it, or else the text the job brings for that (`fallback`).
 *
 * Where a deck is changed, a part that is written again starts from what it holds (`current`):
 * a value it must have and that did not come back is the one it had, and so is every value of a
 * part of which nothing came back. A change never takes a part off a slide by leaving it unsaid.
 * What was named outright to be written (`due`) and did not come is noted (`unwritten`), so that
 * it is asked for again.
 */
function freeDraft(job: FreeJob, written: Record<string, SlotValue>, off: TakenOff | undefined): FreeDraft {
  const { blocks, arrange, base } = job.free;
  const slots = { ...written };
  for (const name of off?.slots ?? []) delete slots[name];
  const had = new Map<number, Array<{ name: string; value: SlotValue; required: boolean }>>();
  for (const [name, value] of Object.entries(job.current ?? {})) {
    const at = partOf(name);
    const kind = at ? blocks[at.block - 1] : undefined;
    const part = at && kind ? FREE_PARTS[kind]?.find((entry) => entry.name === at.part) : undefined;
    if (at && part) had.set(at.block, [...(had.get(at.block) ?? []), { name, value, required: part.required }]);
  }
  const unwritten: string[] = [];
  for (const values of had.values()) {
    const none = values.every(({ name }) => !filled(slots[name]));
    for (const { name, value, required } of values) {
      const missing = none || (required && !filled(slots[name]));
      if (!missing && !off?.back.has(name)) continue;
      slots[name] = value;
      if (missing && job.due?.has(name) && !off?.back.has(name)) unwritten.push(name);
    }
  }
  // The title likewise: left unsaid, it is the one the slide had. An empty one is how a slide is given none.
  if (slots.title === undefined && filled(job.current?.title)) slots.title = job.current?.title as SlotValue;
  else if (!filled(slots.title)) delete slots.title;
  const gone = new Map<number, Lost>(off?.blocks ?? []);
  blocks.forEach((kind, index) => {
    const numbers = NUMBERS[kind];
    if (numbers && !gone.has(index + 1) && !filled(slots[freeSlotName(index + 1, numbers)])) gone.set(index + 1, 'empty');
  });
  const lost = [...gone.values()];
  const left = {
    empty: lost.filter((reason) => reason === 'empty').length,
    unusable: lost.filter((reason) => reason === 'unusable').length,
    surplus: lost.filter((reason) => reason === 'surplus').length,
    unchanged: new Set([...(off?.back ?? [])].map((name) => partOf(name)?.block)).size,
  };
  const slide = (parts: FreeBlockKind[], values: Record<string, SlotValue>): SlideSpec => {
    const set: FreeArrange = { ...arrange };
    if (set.columns !== undefined && set.columns > parts.length) set.columns = parts.length;
    // The title, then the parts in their order: also where some were kept and others written, so that a deck reads the same however it was made.
    const ordered: Record<string, SlotValue> = {};
    for (const name of ['title', ...freeSlots(parts).map((entry) => entry.name), ...Object.keys(values)]) {
      const value = values[name];
      if (value !== undefined && !(name in ordered)) ordered[name] = value;
    }
    return { pattern: FREE_PATTERN, blocks: parts, ...(Object.keys(set).length > 0 ? { arrange: set } : {}), base, slots: ordered };
  };
  if (gone.size === 0) return { slide: slide([...blocks], slots), kept: blocks.map((_, index) => index + 1), left, unwritten };

  const kept = blocks.map((_, index) => index + 1).filter((number) => !gone.has(number));
  const content = freeContent(slide([...blocks], slots));
  let rest: FreeBlockContent[] = kept.map((number) => content.blocks[number - 1] as FreeBlockContent);
  if (rest.length === 0) {
    // Nothing is left to show. The slide says what it is about, in the first words there are for
    // that: a purpose may be empty, and in a change to a deck it is an instruction, which no slide
    // should show while there is anything else to say.
    const title = oneLine(content.title);
    const words = (job.task ? [title, job.fallback, job.purpose] : [job.purpose, title, job.fallback]).map(oneLine).find(Boolean) ?? '';
    rest = [{ text: words }];
    // A slide that says its title as its text says it once.
    if (words === title) delete content.title;
  }
  const made = freeSlide({ ...content, blocks: rest });
  // A picture that was never written is no picture: it stays a missing value, to be said as one.
  for (const [name, value] of Object.entries(made.slots)) if (isImageRef(value) && value.asset === '') delete made.slots[name];
  return { slide: slide(made.blocks ?? [], made.slots), kept, left, unwritten };
}

/** The name a slot of the drafted slide had when the slide was planned. `undefined` for a slot of no planned part. */
function plannedSlot(draft: FreeDraft, slot: string | undefined): string | undefined {
  const at = partOf(slot);
  if (!at) return slot;
  const number = draft.kept[at.block - 1];
  return number === undefined ? undefined : freeSlotName(number, at.part);
}

/**
 * What is said of a free slide beside the check of the deck. Text is held to the budget the
 * writer was shown: the check lays the slide out with its content, where a text that is far too
 * long makes every like part smaller and may still pass. A slot that a change was to write again
 * and that came back with nothing is said, and so is a slide with more photos to generate than
 * are made for one. And parts that were left out are told, as a note on the deck: no rewriting
 * brings them back.
 */
function freeNotes(job: FreeJob, draft: FreeDraft, issues: Issue[], context: CatalogContext): Issue[] {
  const notes: Issue[] = [];
  const reported = new Set(issues.filter((issue) => issue.slide === job.number && issue.slot).map((issue) => issue.slot));
  for (const planned of job.pattern.slots) {
    const at = partOf(planned.slot.name);
    const now = at ? draft.kept.indexOf(at.block) + 1 : 0;
    if (!at || now === 0 || planned.slot.kind !== 'text') continue;
    const slot = freeSlotName(now, at.part);
    const most = textBudget(planned.slot, budgetScale(context)).maxChars;
    const value = draft.slide.slots[slot];
    if (most === undefined || most < LEAST_BUDGET || reported.has(slot) || (typeof value !== 'string' && !Array.isArray(value))) continue;
    const chars = toParagraphs(value as string | string[]).reduce((sum, paragraph) => sum + paragraph.length, 0);
    if (chars > 1.25 * most) notes.push({ severity: 'warning', slide: job.number, slot, ...said('slot-over-chars', { chars, most }) });
  }
  // A text set dense has the paragraphs its writer was told: the check of the deck takes as many as
  // it has, and says what they hold.
  if (job.free.arrange?.dense === true) {
    for (const planned of job.pattern.slots) {
      const at = partOf(planned.slot.name);
      const now = at ? draft.kept.indexOf(at.block) + 1 : 0;
      if (!at || now === 0 || planned.slot.kind !== 'text' || planned.mode !== 'paragraphs') continue;
      const slot = freeSlotName(now, at.part);
      const told = textBudget(planned.slot, budgetScale(context)).maxParagraphs;
      const value = draft.slide.slots[slot];
      if (told === undefined || (typeof value !== 'string' && !Array.isArray(value))) continue;
      const count = toParagraphs(value as string | string[]).filter((paragraph) => paragraph !== '').length;
      if (count > told) notes.push({ severity: 'warning', slide: job.number, slot, ...said('slot-over-paragraphs', { paragraphs: count, most: told }) });
    }
  }
  for (const name of draft.unwritten) {
    const at = partOf(name);
    const now = at ? draft.kept.indexOf(at.block) + 1 : 0;
    if (at && now > 0) notes.push({ severity: 'warning', slide: job.number, slot: freeSlotName(now, at.part), ...said('part-unwritten', {}) });
  }
  if (describedPhotos(job, draft).length > MOST_PHOTOS) notes.push({ severity: 'warning', slide: job.number, ...said('photos-over', { most: MOST_PHOTOS }) });
  const told = (count: number, code: 'parts-left-out-empty' | 'parts-left-out-unusable') => {
    if (count > 0) notes.push({ severity: 'warning', ...said(code, { slide: job.number, count }) });
  };
  told(draft.left.empty, 'parts-left-out-empty');
  told(draft.left.unusable, 'parts-left-out-unusable');
  const over = draft.left.surplus;
  if (over > 0) notes.push({ severity: 'warning', ...said('photos-left-out', { slide: job.number, count: over, most: MOST_PHOTOS }) });
  const same = draft.left.unchanged;
  if (same > 0) notes.push({ severity: 'warning', ...said('parts-unchanged', { slide: job.number, count: same }) });
  return notes;
}

/**
 * Whether a warning says of a chart or a table no more than was so before the change: it has more
 * categories, series, columns or rows than read well in its place, and the part held as many when
 * the writer was shown it. Such a part was written when the slide had fewer parts, or by code.
 * The change is not about its size, and asking again could only cost it a series or a row. So it
 * is no problem to send back. The deck is told all the same, as any check of it tells: the
 * warning stays among its issues. `slot` is the name the part was planned with, `now` its value
 * in the draft.
 */
function largeAsBefore(job: FreeJob, slot: string, issue: Issue, now: SlotValue | undefined): boolean {
  if (issue.severity !== 'warning' || !/^the (?:chart|table) has \d+ (?:categories|series|columns|rows?); the slot fits \d+$/.test(issue.message)) return false;
  const before = job.current?.[slot];
  if (isChartValue(before) && isChartValue(now)) return now.categories.length <= before.categories.length && now.series.length <= before.series.length;
  if (isTableValue(before) && isTableValue(now)) return now.header.length <= before.header.length && now.rows.length <= before.rows.length;
  return false;
}

/**
 * What is said of slots that came back with nothing, once no further answer is asked for: one
 * note on the deck for each slide, in place of a warning on every such slot. The parts hold what
 * they held, so there is nothing on the slide to point at.
 */
function tellUnwritten(issues: Issue[]): Issue[] {
  const parts = new Map<number, Set<number | string>>();
  const rest = issues.filter((issue) => {
    if (issue.message !== UNWRITTEN || issue.slide === undefined) return true;
    parts.set(issue.slide, (parts.get(issue.slide) ?? new Set()).add(partOf(issue.slot)?.block ?? issue.slot ?? ''));
    return false;
  });
  for (const [slide, set] of parts) {
    rest.push({ severity: 'warning', ...said('parts-unwritten', { slide, count: set.size }) });
  }
  return rest;
}

/**
 * Looks at the errors a draft still has on the parts of a free slide, and adds to `off` what is
 * to go: a part whose content is missing or cannot be shown, and the value of an optional slot
 * that cannot be shown. Says whether it found any. A free slide is made with what it has, and
 * does not stop a deck. Only what the writer was asked for is looked at: a value the slide kept
 * from before a change is the deck's own, and what is wrong with it stays an error. A part that
 * was written again, and whose new content cannot be shown, is not taken off: it gets back what
 * it had. Of the photos the writer described, those beyond the number that are made for one
 * slide go too, the last ones first: every one would be a picture to generate.
 */
function takeOff(job: FreeJob, draft: FreeDraft, issues: Issue[], off: TakenOff): boolean {
  const asked = new Set(job.pattern.slots.map((planned) => planned.slot.name));
  let found = false;
  for (const block of describedPhotos(job, draft).slice(MOST_PHOTOS)) {
    off.blocks.set(block, 'surplus');
    found = true;
  }
  for (const issue of issues) {
    const slot = issue.severity === 'error' && issue.slide === job.number ? plannedSlot(draft, issue.slot) : undefined;
    const at = slot !== undefined && asked.has(slot) ? partOf(slot) : undefined;
    const kind = at ? job.free.blocks[at.block - 1] : undefined;
    const part = at && kind ? FREE_PARTS[kind]?.find((entry) => entry.name === at.part) : undefined;
    if (!slot || !at || !part) continue;
    if (!part.required) off.slots.add(slot);
    else if (filled(job.current?.[slot]) && !off.back.has(slot)) off.back.add(slot);
    else off.blocks.set(at.block, issue.message === REQUIRED ? 'empty' : 'unusable');
    found = true;
  }
  return found;
}

/** How many calls that write parts of a slide run side by side. */
const SIDE_BY_SIDE = 3;

interface PartsCall {
  job: FreeJob;
  range: PartRange;
  schema: JsonSchema;
  request: string;
  answer: unknown;
}

/**
 * Writes the free slides that are written by calls of their own: each slide's parts in one call
 * or several, three calls at a time. A call whose answer has problems is asked again by itself,
 * with its own answer and its own problems, up to `maxRepairs` times. What it then returns is
 * laid over what it gave before, so that an answer that holds only what was to be mended loses
 * nothing of the rest. Returns what was written for each slide, by slot, and adds to `off` the
 * parts that still have no usable content.
 */
async function writeApart(args: WriteArgs, jobs: FreeJob[], check: (deck: DeckSpec) => Issue[], off: Map<number, TakenOff>): Promise<Map<number, Record<string, SlotValue>>> {
  const { context } = args;
  const calls: PartsCall[] = jobs.flatMap((job) => {
    const ranges = partRanges(job, context);
    return ranges.map((range) => {
      const task = rangeTask(job, range, ranges.length);
      return { job, range, schema: rangeSchema(job, range, context), request: args.requestFor ? args.requestFor(job, task) : task, answer: undefined };
    });
  });
  const askAll = async (list: PartsCall[], prompt: (call: PartsCall) => string): Promise<void> => {
    for (let at = 0; at < list.length; at += SIDE_BY_SIDE) {
      await Promise.all(
        list.slice(at, at + SIDE_BY_SIDE).map(async (call) => {
          const answer = await (call.range.light ? (args.askLight ?? args.ask) : args.ask)(prompt(call), call.schema, call.range.room);
          call.answer = call.answer === undefined ? answer : mergeRange(call.job, call.range, call.answer, answer);
        }),
      );
    }
  };
  // One of these slides is checked by itself, in a deck that has the settings of the whole and no other slide.
  const frame = args.assemble(new Map());
  const written = new Map<number, Record<string, SlotValue>>();
  const last = new Map<number, { draft: FreeDraft; issues: Issue[] }>();

  await askAll(calls, (call) => call.request);
  for (let attempt = 1; ; attempt++) {
    const problems = new Map<PartsCall, string[]>();
    for (const job of jobs) {
      const mine = calls.filter((call) => call.job === job);
      const answers = mine.map((call) => readRange(job, call.range, call.answer, context));
      const slots = { ...job.carried, ...toSlide(job.pattern, Object.assign({}, ...answers.map((answer) => answer.values)), context).slots };
      written.set(job.number, slots);
      const draft = freeDraft(job, slots, undefined);
      const issues: Issue[] = check({ ...frame, slides: [draft.slide] })
        .filter((issue) => issue.slide === 1)
        .map((issue) => ({ ...issue, slide: job.number }));
      issues.push(...freeNotes(job, draft, issues, context).filter((issue) => issue.slide !== undefined));
      last.set(job.number, { draft, issues });

      const absent = new Set(answers.flatMap((answer) => answer.absent));
      // Only what a call was asked for goes back to it: a value the slide kept from before a change is not its to mend.
      const asked = new Set(job.pattern.slots.map((planned) => planned.slot.name));
      mine.forEach((call, index) => {
        const lines = (answers[index] as RangeAnswer).counts.map((line) => `- ${line}`);
        const parts = new Set(call.range.parts);
        for (const issue of issues) {
          // Too many photos to generate is said of the slide, to every call that writes photos of it.
          if (issue.message === PHOTOS_OVER) {
            if (call.range.runs.some((run) => run.kind === 'photo')) lines.push(`- s${job.number}: ${issue.message}`);
            continue;
          }
          const slot = plannedSlot(draft, issue.slot);
          const at = partOf(slot);
          if (slot === undefined || !asked.has(slot)) continue;
          // A chart or a table that is as large as the slide had it is not sent back for its size.
          if (issue.slot && largeAsBefore(job, slot, issue, draft.slide.slots[issue.slot])) continue;
          if (slot === 'title' && call.range.title) lines.push(`- s${job.number}.title: ${issue.message}`);
          // A part that no item was given for is said once, by the line about its array.
          if (!at || !parts.has(at.block) || absent.has(at.block)) continue;
          lines.push(`- s${job.number} ${partLabel(job.free.blocks, slot)}: ${issue.message}`);
        }
        if (lines.length > 0) problems.set(call, lines);
      });
    }
    if (problems.size === 0 || attempt > args.maxRepairs) break;
    args.onRepair?.(attempt, [...problems.values()].reduce((sum, lines) => sum + lines.length, 0));
    await askAll([...problems.keys()], (call) =>
      [
        call.request,
        `Your previous answer:\n${JSON.stringify(call.answer)}`,
        `It has these problems:\n${(problems.get(call) ?? []).join('\n')}`,
        'Return the full content again with these problems fixed. Change nothing else.',
      ].join('\n\n'),
    );
  }

  for (const job of jobs) {
    const state = last.get(job.number);
    const taken = nothingOff();
    if (state && takeOff(job, state.draft, state.issues, taken)) off.set(job.number, taken);
  }
  return written;
}

/**
 * Asks for the content of the given slides and checks the resulting deck. Problems on the
 * written slides go back to the model, up to `maxRepairs` times. Returns the last draft with
 * every issue it still has; the caller decides what is fatal.
 *
 * A free slide never is: a part of one that has no content, or none that can be shown, once the
 * repairs are spent, is taken off the slide, and a warning on the deck says how many were.
 */
export async function writeSlides(args: WriteArgs): Promise<{ deck: DeckSpec; issues: Issue[] }> {
  const { jobs, context } = args;
  // A free slide of many parts is written by calls of its own; every other slide in one call, as before.
  const together = jobs.filter((job) => !writtenApart(job));
  // A slide whose every slot code sets (a text from a document, its title the heading it opens
  // under) is not asked for: a model is asked only where there is something to write, and an
  // object with nothing in it is no schema every model takes.
  const asking = together.filter((job) => job.pattern.slots.length > 0);
  const schema = contentSchema(asking, context);
  // A free slide that is being changed and has a title is told to say it again unless the request
  // is about it, as a call that writes parts by themselves is told. Here the title may be left
  // out, which keeps it, and "" takes it off the slide.
  for (const job of together) {
    const slots = isFreeJob(job) && titled(job) ? ((schema.properties as Record<string, JsonSchema>)[`s${job.number}`]?.properties as Record<string, JsonSchema> | undefined) : undefined;
    if (slots?.title) slots.title = { ...slots.title, description: `${String(slots.title.description ?? 'Text').replace(/\.$/, '')}. ${REPEAT_TITLE}` };
  }
  const numbers = new Set(asking.map((job) => job.number));

  const check = (deck: DeckSpec): Issue[] => {
    const options: ValidateOptions = { layout: args.layout, assets: knownAssets(deck, context) };
    // With the template a free slide is laid out and its text checked. The check then measures
    // with the fonts of the template as this deck builds it, as rendering will.
    if (args.template && deck.slides.some(isFreeSlide)) return validateDeck(args.template, deck, options);
    if (args.measure) options.measure = args.measure;
    return validateDeck(args.manifest, deck, options);
  };

  const off = new Map<number, TakenOff>();
  const apart = jobs.filter(isFreeJob).filter(writtenApart);
  const parts = apart.length > 0 ? await writeApart(args, apart, check, off) : new Map<number, Record<string, SlotValue>>();

  const build = async (answer: unknown, round: number): Promise<{ deck: DeckSpec; issues: Issue[]; drafts: Map<number, FreeDraft> }> => {
    const content = (typeof answer === 'object' && answer !== null ? answer : {}) as Record<string, unknown>;
    const written = new Map<number, SlideSpec>();
    const drafts = new Map<number, FreeDraft>();
    for (const job of jobs) {
      if (isFreeJob(job)) {
        const slots = parts.get(job.number) ?? { ...job.carried, ...toSlide(job.pattern, content[`s${job.number}`], context).slots };
        const draft = freeDraft(job, slots, off.get(job.number));
        drafts.set(job.number, draft);
        written.set(job.number, draft.slide);
        continue;
      }
      const slide = toSlide(job.pattern, content[`s${job.number}`], context);
      const next: SlideSpec = { pattern: slide.pattern, slots: { ...job.carried, ...slide.slots } };
      if (job.pattern.items !== undefined) next.items = job.pattern.items;
      written.set(job.number, next);
    }
    const deck = args.assemble(written);
    const issues = check(deck);
    for (const job of jobs) {
      const draft = drafts.get(job.number);
      if (draft && isFreeJob(job)) issues.push(...freeNotes(job, draft, issues, context));
    }
    if (args.notes) issues.push(...args.notes(deck, issues, round));
    // A draft with errors cannot be rendered, so there is nothing yet for a review to look at.
    if (args.review && !issues.some((issue) => issue.severity === 'error')) issues.push(...(await args.review(deck)));
    return { deck, issues, drafts };
  };
  // Only slides that were just written can be fixed by asking again. A problem is told under the
  // name the model wrote the slot by: a free slide that lost a part has numbered the others anew.
  // And of a free slide only what was asked for goes back: where a deck is changed, a part the
  // slide kept as it was is not the writer's to mend.
  const freeJobs = new Map(together.filter(isFreeJob).map((job) => [job.number, job]));
  const asked = new Map([...freeJobs].map(([number, job]) => [number, new Set(job.pattern.slots.map((planned) => planned.slot.name))]));
  // A slot code set (a text from a document, the title it opens under) and did not ask of the writer is not the writer's to mend.
  const setByCode = new Map(together.map((job) => [job.number, new Set(Object.keys(job.carried ?? {}).filter((name) => name === job.filled || !job.pattern.slots.some((planned) => planned.slot.name === name)))] as const));
  // `sent` gathers, for each free slide, the slots a round's problems are about.
  const repairable = (draft: { issues: Issue[]; drafts: Map<number, FreeDraft> }, sent?: Map<number, Set<string>>): string[] =>
    draft.issues.flatMap((issue) => {
      if (issue.slide === undefined || !numbers.has(issue.slide)) return [];
      const free = draft.drafts.get(issue.slide);
      const slot = free ? plannedSlot(free, issue.slot) : issue.slot;
      if (slot !== undefined && setByCode.get(issue.slide)?.has(slot)) return [];
      if (free && issue.slot && (!slot || !asked.get(issue.slide)?.has(slot))) return [];
      // Nor is a chart or a table that is as large as the slide had it.
      const job = freeJobs.get(issue.slide);
      if (free && job && slot && issue.slot && largeAsBefore(job, slot, issue, free.slide.slots[issue.slot])) return [];
      if (free && slot) sent?.set(issue.slide, (sent.get(issue.slide) ?? new Set()).add(slot));
      return [`- s${issue.slide}${slot ? `.${slot}` : ''}: ${issue.message}`];
    });
  /**
   * A repair answer laid over the answer it mends, for the free slides alone. A model that is
   * asked to shorten one card may return that card and nothing else. Any other slide then fails
   * for what is missing, as it always did; a free slide is made with what it has, and would be
   * made of that one card. So a slot of a free slide that the new answer leaves out keeps what
   * the answer before gave it, unless the slot is one the problems were about: leaving that one
   * out is an answer to the problem.
   */
  const mended = (before: unknown, next: unknown, sent: Map<number, Set<string>>): unknown => {
    const record = (value: unknown): Record<string, unknown> | undefined => (typeof value === 'object' && value !== null && !Array.isArray(value) ? (value as Record<string, unknown>) : undefined);
    const old = record(before);
    const now = record(next);
    if (!old || !now) return next;
    const merged = { ...now };
    for (const job of together) {
      const key = `s${job.number}`;
      const had = isFreeJob(job) ? record(old[key]) : undefined;
      if (!had) continue;
      const given = record(now[key]) ?? {};
      const stays = (name: string) => !(name in given) && !sent.get(job.number)?.has(name);
      // In the order the slots were first given, with what is new after them.
      merged[key] = { ...Object.fromEntries(Object.entries(had).filter(([name]) => name in given || stays(name))), ...given };
    }
    return merged;
  };

  const room = contentRoom(asking, context);
  // A deck whose every slide was written apart, or set by code, has nothing left to ask for here.
  let answer: unknown = asking.length > 0 ? await args.ask(args.request, schema, room) : {};
  let round = 0;
  let draft = await build(answer, round);
  for (let attempt = 1; attempt <= args.maxRepairs && asking.length > 0; attempt++) {
    const sent = new Map<number, Set<string>>();
    const problems = repairable(draft, sent);
    if (problems.length === 0) break;
    args.onRepair?.(attempt, problems.length);
    const repair = [
      args.request,
      `Your previous answer:\n${JSON.stringify(answer)}`,
      `It has these problems:\n${problems.join('\n')}`,
      'Return the full content again with these problems fixed. Change nothing else.',
    ].join('\n\n');
    answer = mended(answer, await args.ask(repair, schema, room), sent);
    round = attempt;
    draft = await build(answer, round);
  }

  // After the last repair a free slide is made with what it has: a part whose content is still
  // missing, or cannot be shown, is taken off, and the deck says so.
  let found = false;
  for (const job of together) {
    const free = draft.drafts.get(job.number);
    if (!free || !isFreeJob(job)) continue;
    const taken = off.get(job.number) ?? nothingOff();
    if (takeOff(job, free, draft.issues, taken)) {
      off.set(job.number, taken);
      found = true;
    }
  }
  if (found) draft = await build(answer, round);
  return { deck: draft.deck, issues: tellUnwritten(draft.issues) };
}
