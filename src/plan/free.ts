// Free slides as a planner meets them: when one may be planned, what the list of patterns says of
// them, how a plan names their parts, and how a model or a person is told which part something
// is about. How such a slide is laid out is the work of src/free.
import { FREE_BLOCK_KINDS, FREE_PATTERN } from '../deck/spec.ts';
import type { DeckSpec, FreeArrange, FreeBlockKind } from '../deck/spec.ts';
import { FREE_PARTS, describeBlocks, isFreeBlockKind } from '../free/blocks.ts';
import { freeKitOrReason } from '../free/kit.ts';
import { FREE_KIND_NOTES, blocksOfRuns, freeOffer, freePattern } from '../free/offer.ts';
import type { FreeOffer } from '../free/offer.ts';
import type { Template } from '../render/render.ts';
import type { SlotDef, TemplateManifest } from '../template/manifest.ts';
import { planPattern } from './catalog.ts';
import type { CatalogContext, UsablePattern } from './catalog.ts';

/** A slide of a plan that is a free slide: it lists its parts. */
export interface FreeEntry {
  blocks: FreeBlockKind[];
  arrange?: FreeArrange;
}

/** Whether a slide of a plan is a free slide. As in a deck, its list of parts is what says so. */
export const isFreeEntry = <T extends { blocks?: unknown }>(entry: T): entry is T & { blocks: FreeBlockKind[] } => typeof entry === 'object' && entry !== null && Array.isArray(entry.blocks);

/** The most parts a plan may give one slide. More is a slip of the model, not a wish. */
export const MOST_PARTS = 400;
/** The most photos that are generated for one slide. A slide may hold more when pictures were supplied. */
export const MOST_PHOTOS = 6;

/**
 * Whether `free` is an id a list of patterns may name on this template: its manifest lets free
 * slides be planned and it has a layout they can stand on. A template with a pattern of its own
 * called `free` has no free slides to plan: the id is that pattern's.
 */
export function freeKnown(template: Template): boolean {
  const { manifest } = template;
  if (manifest.free === false || manifest.patterns.some((pattern) => pattern.id === FREE_PATTERN)) return false;
  try {
    return typeof freeKitOrReason(template) !== 'string';
  } catch {
    // A template that cannot be read for free slides has none; what else is planned from it is not held up.
    return false;
  }
}

/** Whether a list of patterns, as a brief, a deck or a caller gives it, names `free`. The template is looked at only when one does. */
export const namesFree = (list: unknown): boolean => Array.isArray(list) && list.some((id) => typeof id === 'string' && id.trim() === FREE_PATTERN);

/** What a deck has to say about free slides before any is planned. */
export interface FreeSettings {
  /** The caller's option: `false` turns free slides off. */
  option?: boolean | undefined;
  /** The app's list of patterns, when it has one. */
  app?: string[] | undefined;
  /** The person's list in force, when there is one. */
  limit?: string[] | undefined;
  /** Every slide is to be plain text. */
  textOnly?: boolean;
}

/**
 * Whether a setting turns free slides off for new slides: the caller's option, the template's
 * manifest, or a list of patterns that does not name them. No slide's text is then set smaller,
 * which what is said of a slide that is cut says too. A template with a pattern of its own called
 * `free` has no free slides to turn off.
 */
export function freeTurnedOff(manifest: Pick<TemplateManifest, 'patterns' | 'free'>, settings: Pick<FreeSettings, 'option' | 'app' | 'limit'>): boolean {
  if (manifest.patterns.some((pattern) => pattern.id === FREE_PATTERN)) return false;
  const listed = (list: string[] | undefined) => !list || list.includes(FREE_PATTERN);
  return settings.option === false || manifest.free === false || !listed(settings.app) || !listed(settings.limit);
}

/**
 * Whether new slides of a deck may be free slides, and with which kinds of part.
 *
 * `template` is the template as it was loaded, before a deck's own layouts are built onto it.
 * `refusal` says why no free slide can be made here whatever the lists say, in words that follow
 * the id (`free (free slides are turned off)`, `free needs the deck field "unit"`). A list that
 * does not name `free` is no refusal: free slides are then simply not on offer.
 */
export function freeOnOffer(template: Template, context: CatalogContext, deck: Pick<DeckSpec, 'direction' | 'textScale'>, settings: FreeSettings): { offer?: FreeOffer; refusal?: string } {
  const { manifest } = template;
  // The id belongs to a pattern of the template: there is nothing else to offer under it.
  if (manifest.patterns.some((pattern) => pattern.id === FREE_PATTERN)) return {};
  if (settings.option === false) return { refusal: '(free slides are turned off)' };
  if (manifest.free === false) return { refusal: "(the template's manifest turns free slides off)" };
  // Under a list that does not name them there is nothing to work out: the template is not looked at.
  const listed = (list: string[] | undefined) => !list || list.includes(FREE_PATTERN);
  if (!listed(settings.app) || !listed(settings.limit)) return {};

  try {
    const kit = freeKitOrReason(template);
    if (typeof kit === 'string') return { refusal: `(${kit})` };
    // The frame a free slide stands on may show a text of the deck, which has to be there.
    const frame = freeWritePattern(template, deck, { blocks: ['text'] }, context);
    if ('reason' in frame) return { refusal: frame.reason };
    const offer = freeOffer(template, {
      figures: context.figures,
      photos: context.imageDensity !== 'none',
      graphics: context.assets.some((asset) => asset.kind === 'graphic'),
      icons: context.icons.length > 0,
      ...(settings.textOnly ? { textOnly: true } : {}),
    });
    return offer ? { offer } : {};
  } catch {
    // Free slides are an offer on top of the template's patterns. A template they cannot be worked
    // out for is planned from its patterns, as it was before there were free slides.
    return { refusal: '(free slides cannot be worked out for this template)' };
  }
}

/**
 * The pattern the content of a planned free slide is written against: its slots with the room
 * each has, and which of them the model fills. Optional parts stay optional whatever the caller's
 * layout mode is: a part of a free slide that has no content is simply not drawn. `skip` names
 * slots that are taken care of already. A reason, in words that follow "the free slide", when the
 * slide cannot be written with what was supplied.
 *
 * What a slot says of itself is the engine's: what the part is for, and its limits in its fields,
 * each said once. Nothing is taken from a description here and nothing put into one, with one
 * exception that is the planner's own to know: how many photos are made for one slide.
 *
 * `withParts` names optional slots to count in besides those the engine plans by rule: where a
 * deck is changed, the optional slots that hold a value on the slide (a caption under the one
 * photo, a heading over the one text), under the names they have after the change. The slide is
 * then set out with them, and one that is to be written again is offered again.
 *
 * `lengths` gives texts the length a person asked for, in characters, by slot name: the slide is
 * set out for them, and their budgets say what their boxes hold at that length (length.ts says
 * what the writer is then told). `paragraphs` gives some of them the number of paragraphs they
 * will have, where that is known: a text that code fills is set out as it will be. On a slide set
 * dense, `asIs` gives the lengths of those paragraphs, of a text code sets as it is.
 */
export function freeWritePattern(
  template: Template,
  deck: Pick<DeckSpec, 'direction' | 'textScale'>,
  entry: {
    blocks: readonly FreeBlockKind[];
    arrange?: FreeArrange | undefined;
    base?: string | undefined;
    lengths?: Readonly<Record<string, number>> | undefined;
    paragraphs?: Readonly<Record<string, number>> | undefined;
    asIs?: Readonly<Record<string, readonly number[]>> | undefined;
  },
  context: CatalogContext,
  skip?: ReadonlySet<string>,
  withParts?: readonly string[],
): { pattern: UsablePattern; base: string } | { reason: string } {
  const planned = freePattern(template, deck, { blocks: [...entry.blocks], arrange: entry.arrange, base: entry.base, lengths: entry.lengths, paragraphs: entry.paragraphs, asIs: entry.asIs }, context.icons.length > 0, withParts);
  if (typeof planned === 'string') return { reason: `cannot be made from this template: ${planned}` };
  const photos = entry.blocks.filter((kind) => kind === 'photo').length;
  const slots = planned.pattern.slots.map((slot): SlotDef => {
    // Every photo that is described is a picture to generate, and so many are made for one slide.
    if (slot.kind === 'image' && slot.imageKind === 'photo' && photos > MOST_PHOTOS) {
      return { ...slot, description: [slot.description, `Of the photos of this slide at most ${MOST_PHOTOS} can be made from a description: give the others a supplied photo`].filter(Boolean).join('. ') };
    }
    return slot;
  });
  const usable = planPattern({ ...planned.pattern, slots }, { ...context, layout: 'adaptive' }, skip);
  if ('reason' in usable) return usable;
  // A chart or a table is asked of the writer like any part: left optional, a model leaves every
  // chart of a slide out (seen with the production model on a slide of six). One the brief has no
  // numbers for is given empty, which its schema allows here, and such a part is then not shown.
  const written = usable.pattern.slots.map((slot) => (slot.mode === 'chart' || slot.mode === 'table' ? { ...slot, required: true, mayBeEmpty: true } : slot));
  return { pattern: { ...usable.pattern, slots: written }, base: planned.base };
}

/** The parts of a stored free slide that are kinds of part at all, and one text when none is. */
export function entryBlocks(entry: { blocks?: unknown }): FreeBlockKind[] {
  const blocks = (Array.isArray(entry.blocks) ? entry.blocks : []).filter(isFreeBlockKind);
  return blocks.length > 0 ? blocks : ['text'];
}

/**
 * What a slide may hold several of, in words: "several charts, tables or pictures on one slide".
 * Of the three it names those among the kinds on offer, and it is empty when none of them is: a
 * brief without numbers gets neither chart nor table, an app without an image model and without
 * pictures of its own gets no picture, and neither should read that it could.
 */
export function severalOnOneSlide(kinds: readonly FreeBlockKind[]): string {
  const several = [kinds.includes('chart') ? 'charts' : '', kinds.includes('table') ? 'tables' : '', kinds.includes('photo') || kinds.includes('graphic') ? 'pictures' : ''].filter(Boolean);
  return several.length > 0 ? `several ${several.length < 2 ? several.join('') : `${several.slice(0, -1).join(', ')} or ${several.at(-1)}`} on one slide` : '';
}

/**
 * The line of the list of patterns that offers free slides. It is written here and not with the
 * template's patterns: a free slide has no slots to list before its parts are chosen, and is no
 * plain text unless the deck is. `changes` says what is being changed, where it is no new plan: a
 * plan names the parts of a slide in "shows", and a change to a deck names them one by one in "parts".
 */
export function freeLine(
  kinds: readonly FreeBlockKind[],
  options: { comfortable?: number; alone?: boolean; plain?: boolean; closed?: boolean; asked?: 'brief' | 'request'; changes?: 'plan' | 'deck' },
): string {
  const parts = kinds.map((kind) => `${kind} (${FREE_KIND_NOTES[kind]})`).join(', ');
  const sentences = [`A slide the software lays out itself from the parts you name in "${options.changes === 'deck' ? 'parts' : 'shows'}", any number of each: ${parts}.`];
  if (options.closed) {
    // Free slides are not on offer for new slides; the plan or the deck has one, which stays what it is.
    sentences.push(`Only for the slides of the ${options.changes ?? 'plan'} that have it already.`);
  } else {
    if (!options.alone) {
      // Of charts, tables and pictures the sentence names those that are on offer.
      const several = severalOnOneSlide(kinds);
      sentences.push(`Only for content that no layout above shows: ${several ? `${several}, ` : ''}more items than a layout takes, or parts that no layout combines.`, 'A slide that a layout above shows well uses that layout.');
    }
    const most = options.comfortable ?? 3;
    sentences.push(`About ${most} ${most === 1 ? 'part reads' : 'parts read'} well at the template's text size; more are set smaller, so plan more only when the ${options.asked ?? 'brief'} asks for them.`);
  }
  if (options.plain) sentences.push('Plain text.');
  return `- ${FREE_PATTERN}: ${sentences.join(' ')}`;
}

/** The rule a new outline is given when free slides are on offer. */
export const FREE_OUTLINE_RULE =
  'A slide with the layout free names its parts in "shows", in reading order. ' +
  'A number of things the brief asks to see on one slide (twelve charts, forty logos) is that number of parts, however many.';

/** The rule a change to a plan is given when the plan may hold free slides. */
export const FREE_CHANGE_RULE = 'A slide with the layout free names its parts in "shows". Repeat them exactly as the plan has them unless the request changes what that slide shows.';

/** Said with the share of photos when free slides are on offer: a photo on a free slide is a photo too. */
export const withFreePhotos = (sentence: string): string => sentence.replace(/\.$/, ' (a free slide with a photo counts).');

/** How a free slide is set out, in the words a plan is listed in: "3 side by side; main part at the left". */
function describeArrange(arrange: FreeArrange | undefined): string[] {
  const words: string[] = [];
  if (arrange?.columns !== undefined) words.push(`${arrange.columns} side by side`);
  if (arrange?.lead === 'none') words.push('all parts alike');
  else if (arrange?.lead === 'top') words.push('main part on top');
  else if (arrange?.lead) words.push(`main part at the ${arrange.lead}`);
  return words;
}

/** A free slide of a plan in a few words: "6 charts and a note; 3 side by side; main part at the left". */
export function describeEntry(entry: FreeEntry): string {
  return [describeBlocks(entry.blocks), ...describeArrange(entry.arrange)].join('; ');
}

/**
 * The runs of an answer as counts. Like parts belong in one entry with their count. A model that
 * gives one entry for each part instead tends to number them: chart 1, chart 2, up to chart 8.
 * Read as counts that would be thirty-six charts, so entries of one kind that follow one another
 * and count up from 1 are taken for what they are: as many parts as there are entries.
 */
function countedRuns<Kind>(runs: ReadonlyArray<{ kind: Kind; number: number }>): Array<{ kind: Kind; number: number }> {
  const counted: Array<{ kind: Kind; number: number }> = [];
  for (let at = 0; at < runs.length; ) {
    let end = at;
    while (end + 1 < runs.length && runs[end + 1]?.kind === runs[at]?.kind) end++;
    const group = runs.slice(at, end + 1);
    const numbered = group.length >= 2 && group.every((run, index) => run.number === index + 1);
    if (numbered) counted.push({ kind: (group[0] as { kind: Kind }).kind, number: group.length });
    else counted.push(...group);
    at = end + 1;
  }
  return counted;
}

/** What the answer of a model may say about the parts of a free slide. */
export interface FreeReading {
  /** The kinds a slide may name. A part of another kind is passed over. */
  kinds: readonly FreeBlockKind[];
  /** The most photo parts on one slide: those that can be generated, and one for each photo supplied. */
  photos: number;
  /** The slides of the plan that is being changed, so that a free slide whose parts come back empty keeps the parts it had. */
  old?: ReadonlyArray<{ pattern: string; purpose: string; blocks?: FreeBlockKind[]; arrange?: FreeArrange }>;
}

/** What may be read from an answer, given what is on offer and what was supplied. */
export function freeReading(kinds: readonly FreeBlockKind[], context: CatalogContext, old?: FreeReading['old']): FreeReading {
  const reading: FreeReading = { kinds, photos: MOST_PHOTOS + context.assets.filter((asset) => asset.kind === 'photo').length };
  if (old) reading.old = old;
  return reading;
}

/**
 * Reads the parts of a free slide from one entry of an answer: `shows`, `shows_main` and
 * `side_by_side`. `number` is the slide's number in the answer, `purpose` its sentence.
 *
 * What is wrong with them is a `problem` to send back once. With `final`, the last answer there
 * will be, nothing is a problem any more: a slide without parts becomes one text, and a slide
 * with too many is cut. No answer makes planning fail over the parts of a free slide.
 */
export function readShows(raw: unknown, number: number, purpose: string, reading: FreeReading, final: boolean): FreeEntry & { problem?: string } {
  const given = (typeof raw === 'object' && raw !== null ? raw : {}) as { shows?: unknown; shows_main?: unknown; side_by_side?: unknown };
  // The runs that name a kind on offer and a number of them. A number is not taken at its word
  // beyond what makes a slide too full to plan: an answer may say a thousand million.
  let room = MOST_PARTS + 1;
  let asked = 0;
  const named = (Array.isArray(given.shows) ? given.shows : [])
    .filter((run): run is { kind: unknown; count?: unknown; number?: unknown } => typeof run === 'object' && run !== null)
    // `number` is what the count was called before; a stored answer or a stub may still say it.
    .map((run) => ({ kind: run.kind, number: Math.floor(Number(run.count ?? run.number)) }))
    .filter((run) => isFreeBlockKind(run.kind) && reading.kinds.includes(run.kind) && run.number >= 1);
  const runs = countedRuns(named)
    .map((run) => {
      asked += run.number;
      const number = Math.min(run.number, room);
      room -= number;
      return { kind: run.kind, number };
    });
  let blocks = blocksOfRuns(runs);
  const arrange: FreeArrange = {};
  const lead = given.shows_main;
  if (lead === 'left' || lead === 'right' || lead === 'top' || lead === 'none') arrange.lead = lead;
  const columns = Math.floor(Number(given.side_by_side));
  if (columns >= 1) arrange.columns = columns;
  let problem: string | undefined;

  // The free slide of the plan that is being changed which this one is: the one with its purpose, or else the one in its place.
  const free = (reading.old ?? []).map((entry, index) => ({ entry, index })).filter(({ entry }) => isFreeEntry(entry) && entry.blocks.length > 0);
  const before = (free.find(({ entry }) => entry.purpose === purpose) ?? free.find(({ index }) => index === number - 1))?.entry;
  if (before && (blocks.length === 0 || blocks.join() === entryBlocks(before).join())) {
    // A change that says nothing of the slide's parts, or repeats them, leaves the slide as it was
    // planned: its parts, and how they are set out wherever the answer does not say otherwise.
    const kept: FreeEntry = { blocks: entryBlocks(before) };
    const set: FreeArrange = blocks.length === 0 ? { ...before.arrange } : { ...before.arrange, ...arrange };
    if (arrange.columns !== undefined && set.columns !== undefined) set.columns = Math.min(set.columns, kept.blocks.length);
    if (Object.keys(set).length > 0) kept.arrange = set;
    return kept;
  }
  if (before?.arrange) {
    // Its parts change. How they are set out stays as it was planned wherever the answer does not
    // say otherwise: "auto" and 0 are what an answer gives for a slide whose arrangement the
    // request does not speak of, and a wish a person made is not undone by them.
    if (arrange.lead === undefined && before.arrange.lead !== undefined) arrange.lead = before.arrange.lead;
    const planned = before.arrange.columns;
    if (arrange.columns === undefined && typeof planned === 'number' && Number.isInteger(planned) && planned >= 1) arrange.columns = planned;
  }
  if (blocks.length === 0) {
    if (final) blocks = ['text'];
    else problem = `slide ${number}: a free slide needs "shows", such as 6 charts and 1 note`;
  }
  if (blocks.length > MOST_PARTS) {
    if (final) blocks = blocks.slice(0, MOST_PARTS);
    else problem = `slide ${number}: a free slide takes at most ${MOST_PARTS} parts, not ${asked}`;
  }
  if (blocks.filter((kind) => kind === 'photo').length > reading.photos) {
    if (final) {
      let photos = 0;
      blocks = blocks.filter((kind) => kind !== 'photo' || ++photos <= reading.photos);
    } else {
      problem = `slide ${number}: at most ${MOST_PHOTOS} photos can be made for one slide; supply pictures for more`;
    }
  }
  if (arrange.columns !== undefined) arrange.columns = Math.max(1, Math.min(arrange.columns, blocks.length));

  const entry: FreeEntry & { problem?: string } = { blocks };
  if (Object.keys(arrange).length > 0) entry.arrange = arrange;
  if (problem) entry.problem = problem;
  return entry;
}

/** The kinds of part some slides use, in the order kinds are listed to a model. */
export function kindsUsed(slides: ReadonlyArray<{ blocks?: unknown }>, also: readonly FreeBlockKind[] = []): FreeBlockKind[] {
  const used = new Set<unknown>(also);
  for (const slide of slides) if (Array.isArray(slide.blocks)) for (const kind of slide.blocks) used.add(kind);
  return FREE_BLOCK_KINDS.filter((kind) => used.has(kind));
}

/**
 * Which part of a free slide a slot is, in words: `part 7 (card) text` for `b7Text` of a slide
 * whose seventh part is a card. The slot's own name is left out where the part has one slot or
 * the slot is the part itself (`part 3 (chart)`). `undefined` for a slot that is no part: the title.
 */
export function partLabel(blocks: readonly FreeBlockKind[] | undefined, slot: string | undefined): string | undefined {
  const match = /^b(\d+)([A-Z][A-Za-z]*)$/.exec(slot ?? '');
  const kind = match ? blocks?.[Number(match[1]) - 1] : undefined;
  if (!match || !kind || !FREE_PARTS[kind]) return undefined;
  const name = (match[2] as string).toLowerCase();
  const whole = name === kind || FREE_PARTS[kind].length === 1 || (kind === 'list' && name === 'items');
  return `part ${match[1]} (${kind})${whole ? '' : ` ${name}`}`;
}
