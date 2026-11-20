import { parseHex } from '../color/space.ts';
import { unstatedChartValues } from '../deck/numbers.ts';
import { DeckValidationError } from '../deck/validate.ts';
import type { Issue, LayoutMode } from '../deck/validate.ts';
import { FREE_PATTERN } from '../deck/spec.ts';
import type { DeckColors, DeckSpec, FreeArrange, FreeBlockKind, SlotValue } from '../deck/spec.ts';
import { describeBlocks } from '../free/blocks.ts';
import { freeKit } from '../free/kit.ts';
import type { FreeKit } from '../free/kit.ts';
import { freeTitleLines } from '../free/plan.ts';
import { blockRuns, freeTextRoom } from '../free/offer.ts';
import { freeBodySize } from '../free/plan.ts';
import type { FreeOffer } from '../free/offer.ts';
import type { JsonRequest, JsonSchema, LlmProvider } from '../llm/provider.ts';
import type { Template } from '../render/render.ts';
import { TEXT_SIZES, itemRange, textBudget } from '../template/manifest.ts';
import { measurerFor } from '../text/measure.ts';
import type { Direction, PatternDef, TemplateManifest, TextSize, TextSlotDef } from '../template/manifest.ts';
import { budgetScale, buildCatalog, describeCatalog, isPlainText, outlineEntrySchema, outlineSchema, planPattern, readItems } from './catalog.ts';
import type { AssetDef, Catalog, CatalogContext, IconDef, ImageDensity, UsablePattern, WriteItem } from './catalog.ts';
import { withLayouts } from '../template/compose.ts';
import type { PatternDesign } from '../template/compose.ts';
import { FREE_CHANGE_RULE, FREE_OUTLINE_RULE, MOST_PARTS, describeEntry, entryBlocks, freeKnown, freeLine, freeOnOffer, freeReading, freeTurnedOff, freeWritePattern, isFreeEntry, kindsUsed, namesFree, readShows, withFreePhotos } from './free.ts';
import type { FreeReading } from './free.ts';
import {
  DOCUMENT_CHANGE_RULES,
  DOCUMENT_OPENING,
  DOCUMENT_VARY,
  MOST_DOCUMENT_CHARS,
  MOST_DOCUMENT_SLIDES,
  askedSlidesProperty,
  atomsOf,
  documentBlock,
  documentChars,
  documentLanguageProperty,
  documentRules,
  documentWriteRule,
  fillParagraphs,
  inLanguage,
  grossGap,
  isFixedCount,
  keptCheck,
  keptIssues,
  keptWarnings,
  leftOutProperty,
  linesJoined,
  minMax,
  needOf,
  openingSection,
  packPassages,
  passageText,
  purposeOf,
  readAskedSlides,
  readDocument,
  readLanguageTag,
  readPassageNumbers,
  settleDocumentSlide,
  settlePassages,
  shapeOf,
  slideDocument,
  slideText,
  textCut,
  textCutDense,
} from './document.ts';
import type { DocumentPlan, KeptCheck, SettleOptions, SlideDocument } from './document.ts';
import {
  WORDS_CHANGE_RULE,
  charsOfWords,
  densable,
  documentParagraphs,
  documentRoom,
  freeLengths,
  heldBy,
  homeFor,
  homeForChars,
  partPurpose,
  readWords,
  runningNames,
  runningSize,
  setSmall,
  shortOfWords,
  shortOfWordsDense,
  tensDown,
  textRoom,
  toldAsAsked,
  withLength,
  withPassages,
  wordsOfChars,
  wordsRule,
  writtenLength,
  writtenShort,
  writtenUnder,
} from './length.ts';
import type { LengthPlanning, TextFirst, TextHome, ToldLength } from './length.ts';
import { composeForPlan, usedLayouts } from './layouts.ts';
import type { ComposeLayoutsResult } from './layouts.ts';
import type { DeckReview } from './review.ts';
import { cleanText, describeParts, suppliesText, systemPrompt, writeSlides, writtenApart } from './write.ts';
import type { WriteJob } from './write.ts';
import { isRtlLanguage } from '../util/language.ts';

/** What the person asked for. */
export interface Brief {
  /** What the deck is about, in the person's own words. */
  prompt: string;
  /** A fixed number of slides, or a range. Defaults to 6 to 10. */
  slideCount?: number | { min: number; max: number };
  /** BCP 47 tag of the language to write in. Defaults to the template's language. */
  language?: string;
  /** Reading direction. Defaults to the language's. */
  direction?: Direction;
  /** How many slides should carry a photo. Defaults to `medium`. */
  imageDensity?: ImageDensity;
  /** Anything else to respect: audience, tone, points that must be covered. */
  instructions?: string;
  /** Deck-level values the template needs, e.g. the unit shown in the header. */
  fields?: Record<string, string>;
  /** Size of body text. Larger text means less of it fits on a slide. Defaults to `normal`. */
  textSize?: TextSize;
  /**
   * Plan every slide as plain text: a title with paragraphs or a list, and no pictures, icons,
   * cards, figures, tables or charts. Only the template's plain text patterns are offered.
   */
  textOnly?: boolean;
  /**
   * The only patterns the deck may use, by id, when the person asked for that: "just the text
   * layout". Every slide then uses one of them, the first and the last too. With `patterns` in
   * the options as well, the deck may use what is in both lists. It comes before `textOnly`. The
   * list is kept with the plan and copied to the deck (`DeckSpec.patterns`); a later change keeps
   * to it unless that change asks for another pattern. An id the template does not have is a
   * `PlanError` when the deck is planned, and so is a pattern that cannot be filled with what
   * was supplied (one that needs icons when there is no icon set): the error says which and why.
   * The list may name `free`, the id of free slides, alone or with patterns; under a list that
   * does not, no slide is a free slide.
   */
  patterns?: string[];
  /**
   * Colours in place of the template's, such as `{ accent: '1E6B3A' }`: six hex digits each.
   * They are copied to the deck (`DeckSpec.colors`) when it is written. No model reads a colour
   * out of the words of a brief: a colour the person asked for is stated here.
   */
  colors?: DeckColors;
  /**
   * A document to make the deck from, word for word: a report, minutes, an article. The slides
   * carry its own text, every part of it, in its order, cut only where a slide's limits leave no
   * room, over as many slides as it takes. With a fixed `slideCount`, or a number of slides the
   * brief names, it goes on that many slides, set smaller where it must (see `FreeArrange.dense`).
   * `prompt` then says what the deck is for, and `instructions` any wish ("only the risks", "a
   * short summary"). A line ends a paragraph; a line that starts with "-", "•" or a number is a
   * list item; a short line on its own, or one that starts with "#", is a heading. At most
   * `MOST_DOCUMENT_CHARS` characters. Left out or blank, the deck is planned as before.
   */
  document?: string;
}

export type PlanProgress =
  | { stage: 'outline' }
  | { stage: 'layouts' }
  | { stage: 'write'; slides: number }
  | { stage: 'repair'; attempt: number; issues: number };

export interface PlanOptions {
  llm: LlmProvider;
  /** Defaults to `strict`. */
  layout?: LayoutMode;
  /** The app's icon set. Patterns with icon slots are only offered when this is given. */
  icons?: IconDef[];
  /** Ready-made pictures the deck may use, e.g. rendered diagrams. */
  assets?: AssetDef[];
  /**
   * Restrict the deck to these pattern ids: what the app allows. A brief may narrow the list
   * further (`Brief.patterns`) and never widens it. Free slides obey the list like a pattern:
   * a list that does not name `free` gets none (see `free`).
   */
  patterns?: string[];
  /**
   * When patterns built around numbers may be used. `from-brief` (default) offers them only if
   * the brief contains figures; a model given such a pattern and no data tends to invent some.
   */
  numberPatterns?: 'from-brief' | 'always' | 'never';
  /** How many times the model may be asked to fix a draft that breaks the template's limits. Defaults to 2. */
  maxRepairs?: number;
  /**
   * A second look at every draft: `designReview` for the code-only design checks, `visionReview`
   * for a model that looks at pictures of the slides, or a function of your own. What it finds is
   * sent back to the writing model in the repair rounds, and what remains is returned with the
   * warnings. Nothing reviews a draft unless this is set.
   */
  review?: DeckReview;
  /**
   * Up to this many new layouts may be made for the deck, in the template's own look, for slides
   * whose content none of its layouts shows well: three goals on a slide of running text, a plan
   * in steps. After the outline, one light call decides; each layout then takes one or two calls,
   * side by side, about 10 to 20 seconds. The designs are kept in the plan and the deck
   * (`layouts`); the template is not changed. A layout that cannot be made leaves its slides as
   * planned. Off (0) by default. None is made for a deck whose brief limits it to a list of
   * patterns: a new layout would be one the person did not name. And none is made while free
   * slides are on offer (see `free`): a free slide is what a deck gets for content that no
   * layout shows.
   */
  composeLayouts?: number;
  /**
   * Whether a slide may be a free slide: one the software lays out itself from the parts the plan
   * names (six charts and a note, twelve cards), for content that no pattern of the template
   * shows. On unless `false`. Free slides are offered when the template can carry them, and when
   * no list of patterns is in force or the list names `free`: an app that passes `patterns`
   * gets none unless its list says so. While they are on offer `composeLayouts` makes nothing.
   * `false` gives the planning of 0.5 back. A plan or a deck that already has a free slide is
   * written and changed all the same: the option decides only what a new slide is offered.
   */
  free?: boolean;
  onProgress?: (event: PlanProgress) => void;
  signal?: AbortSignal;
}

export interface OutlineEntry {
  pattern: string;
  purpose: string;
  /** For a pattern with a row of items: how many items the slide has. */
  items?: number;
  /** For a free slide, whose `pattern` is `free`: its parts in reading order. No other slide has it. */
  blocks?: FreeBlockKind[];
  /** For a free slide: how its parts are set out, where the plan says so. */
  arrange?: FreeArrange;
  /**
   * The words of running text asked for this slide ("200 words"), where an amount was asked. The
   * plan gives such a slide a layout or a free slide that holds the text, and the writer is told
   * the length. A slide asked for more than one slide holds at the template's size is set smaller
   * on a free slide (`arrange.dense`); it is never spread over more slides. On a slide that carries
   * passages of a document it asks for a shorter text, and is read only when it is under four
   * fifths of the passages' words.
   */
  words?: number;
  /**
   * In a deck made from a document: the passages this slide carries, by number from 1, as
   * `documentPassages` numbers them. Set by the planner. A text that went over several slides has
   * its own run of passages on each.
   */
  passages?: number[];
  /**
   * Internal, set by a change to a plan (`editPlan`): the parts of this slide besides its text were
   * asked for in a change, so they stay beside its text in later changes too, whatever they leave of
   * it: a free slide's parts, and a layout's chart or table. A layout of cards, figures, a list or a
   * quote gives way to an amount of text asked for it all the same, as in 0.6.2. A plan stored
   * without it needs nothing. It is not shown to a model.
   */
  partsAsked?: true;
}

export interface PlanResult {
  deck: DeckSpec;
  /** Limits the final draft still exceeds, e.g. text a little over a slot's budget. */
  warnings: Issue[];
  outline: OutlineEntry[];
}

export class PlanError extends Error {
  constructor(message: string) {
    super(message);
    this.name = 'PlanError';
  }
}

export function directionOf(language: string): Direction {
  return isRtlLanguage(language) ? 'rtl' : 'ltr';
}

/** Whether a text contains any digit, in any script. */
export function hasFigures(text: string): boolean {
  return /\p{Nd}/u.test(text);
}

/**
 * A text without the codes of the colours a deck was given. "Make it 1E6B3A" names a colour, and
 * its digits are no figures of the brief: they must not bring the patterns built around numbers.
 */
export function withoutColorCodes(text: string, ...given: unknown[]): string {
  const codes = new Set<string>();
  for (const colors of given) {
    if (typeof colors !== 'object' || colors === null) continue;
    const { accent, accent2, replace } = colors as { accent?: unknown; accent2?: unknown; replace?: unknown };
    const pairs = typeof replace === 'object' && replace !== null ? Object.entries(replace as Record<string, unknown>).flat() : [];
    for (const value of [accent, accent2, ...pairs]) {
      const hex = parseHex(value);
      if (!hex) continue;
      codes.add(hex);
      // The short form too, where the colour has one: "#1b3" for 11BB33.
      if (hex[0] === hex[1] && hex[2] === hex[3] && hex[4] === hex[5]) codes.add(`${hex[0]}${hex[2]}${hex[4]}`);
    }
  }
  let result = text;
  for (const code of codes) result = result.replace(new RegExp(`#?\\b${code}\\b`, 'gi'), ' ');
  return result;
}

/**
 * Binds a provider to a system prompt and an optional abort signal. `room` is how many tokens
 * the answer may reasonably take: a model sometimes runs away inside a text field, and a limit
 * is what ends that after seconds instead of minutes. `effort` is `low` for calls that choose
 * and arrange rather than write.
 */
export function asker(llm: LlmProvider, system: string, signal: AbortSignal | undefined, effort?: 'low') {
  return (prompt: string, schema: JsonSchema, room?: number): Promise<unknown> => {
    const request: JsonRequest = { system, prompt, schema };
    if (room !== undefined) request.maxOutputTokens = Math.ceil(room);
    if (effort) request.effort = effort;
    if (signal) request.signal = signal;
    return llm.generateJson(request);
  };
}

/** The caller's review, bound to the caller's abort signal. */
export function reviewWith(options: Pick<PlanOptions, 'review' | 'signal'>): ((deck: DeckSpec) => Issue[] | Promise<Issue[]>) | undefined {
  const { review, signal } = options;
  return review ? (deck) => review(deck, signal ? { signal } : {}) : undefined;
}

/** Room for a plan: a title, and for each slide a pattern id and a sentence. Generous, since it is a ceiling. */
export const planRoom = (slides: number): number => 400 + 160 * slides;

/** Room added for each slide of a plan that may be a free slide and name its parts. A ceiling too. */
export const FREE_ENTRY_ROOM = 12;

/** How many slides carry a photo, for a deck planned from a brief. A number of pictures the brief itself names comes first. */
const DENSITY_GUIDANCE: Record<ImageDensity, string> = {
  none: 'Do not plan any photos.',
  low: 'Unless the brief says how many pictures it wants: use a pattern with a photo on at most a quarter of the slides.',
  medium: 'Unless the brief says how many pictures it wants: use a pattern with a photo on about half of the slides.',
  high: 'Unless the brief says how many pictures it wants: prefer patterns with a photo wherever they suit the content.',
};

/** The same for a change to a plan. It speaks of the slides the change adds: the others were planned already. */
const ADDED_DENSITY_GUIDANCE: Record<ImageDensity, string> = {
  none: DENSITY_GUIDANCE.none,
  low: 'Unless the request says how many pictures it wants: of the slides the change adds, at most a quarter use a pattern with a photo.',
  medium: 'Unless the request says how many pictures it wants: of the slides the change adds, about half use a pattern with a photo.',
  high: 'Unless the request says how many pictures it wants: for the slides the change adds, prefer patterns with a photo wherever they suit the content.',
};

/**
 * What a person asks about the form of a deck comes before every habit of the planner. It is the
 * first rule wherever patterns are chosen: for a new plan, a changed plan and a changed deck.
 */
export const wishesRule = (asked: 'brief' | 'request'): string =>
  `What the ${asked} asks about the form of the deck comes first, before variety, before the usual choices below and before what a pattern's description advises: ` +
  'how many slides, which kinds of slides, how many things on a slide, how many side by side. ' +
  'Do exactly that, even when it means one slide or one pattern for every slide.';

/** Said when a plan or a deck is changed that was limited to a list of patterns. The list was a wish, so a later wish may go beyond it. */
export const limitRule = (limit: string[]): string =>
  `This deck was limited to these layouts: ${limit.join(', ')}. Keep to them unless the requested change asks for another.`;

/**
 * What a model that plans a change is told about colour, beside the rule that colours are not its
 * work: whether the software set one with this change. Its summary is what the person reads, so
 * it must not claim a colour that nobody set, nor hide that a wish for one was not met.
 */
export const colorNote = (set: boolean): string =>
  set
    ? 'The software sets the colour that was asked for with this change: say nothing about colour in summary.'
    : 'No colour was given with this change: if the request asks for one, say in summary that the colour was not changed.';

/** A list of pattern ids as it is kept: each id once, without stray spaces. Anything that is not a list of ids gives none. */
export function patternIds(list: unknown): string[] {
  if (!Array.isArray(list)) return [];
  return [...new Set(list.filter((id): id is string => typeof id === 'string').map((id) => id.trim()).filter(Boolean))];
}

/** The lists of patterns a deck may use, as the stages of planning read them. */
export interface PatternLists {
  /** What the caller's options allow. Nothing outside it is offered, at any stage. */
  allow?: string[];
  /** What the person limited the deck to, of the patterns the template has. This is what a deck keeps. */
  asked?: string[];
  /**
   * The person's list in force: `asked`, narrowed to `allow` when both are given. A new plan is
   * offered these patterns and no others; a later change is told to keep to them.
   */
  limit?: string[];
}

/**
 * Works out which patterns a deck may use from the two lists there can be: the app's
 * (`PlanOptions.patterns`) and the person's (`Brief.patterns`, later `DeckSpec.patterns`).
 *
 * `fresh` says the person's list comes with a new brief. Then an id the template does not have,
 * or a list that shares nothing with the app's, is a `PlanError` that names what is wrong, so it
 * can be put right. At every later stage a list is read leniently instead, since templates change
 * after plans and decks are kept: ids the template does not have are dropped, and when nothing is
 * left the deck is planned as if it had no list.
 *
 * `free` says that a list may name `free`, the id of free slides, beside the template's own
 * patterns. Free slides obey a list like any pattern: they are planned only where no list is in
 * force or the list names them.
 */
export function allowedPatterns(
  manifest: Pick<TemplateManifest, 'name' | 'patterns'>,
  lists: { app?: string[] | undefined; person?: string[] | null | undefined },
  fresh: boolean,
  free = false,
): PatternLists {
  const { app } = lists;
  const result: PatternLists = {};
  if (app) result.allow = app;
  const given = patternIds(lists.person);
  if (given.length === 0) return result;

  const known = new Set(manifest.patterns.map((pattern) => pattern.id));
  const named = (id: string) => known.has(id) || (free && id === FREE_PATTERN);
  const unknown = given.filter((id) => !named(id));
  if (fresh && unknown.length > 0) {
    throw new PlanError(
      `Template "${manifest.name}" has no pattern called ${unknown.map((id) => `"${id}"`).join(', ')}, so the deck cannot be limited to ${unknown.length === 1 ? 'it' : 'them'}. Its patterns are: ${[...known].join(', ')}`,
    );
  }
  const asked = given.filter(named);
  if (asked.length === 0) return result;
  result.asked = asked;

  const limit = app ? asked.filter((id) => app.includes(id)) : asked;
  if (limit.length > 0) result.limit = limit;
  else if (fresh) {
    throw new PlanError(`The deck is limited to the patterns ${asked.join(', ')}, and the options allow only ${(app ?? []).join(', ')}: no pattern is in both lists`);
  }
  return result;
}

/**
 * A list of patterns set with a change to a plan or a deck, as it is kept. An empty list (`null`
 * or `[]` was given) lifts the limit. Ids the template does not have are dropped, without an
 * error: only a new brief is held to its ids. A list that names no pattern of the template sets
 * nothing (`undefined`), so that a slip never lifts a limit. `free` says that the list may name
 * `free`, the id of free slides (see {@link allowedPatterns}).
 */
export function givenPatterns(manifest: Pick<TemplateManifest, 'patterns'>, list: unknown, free = false): string[] | undefined {
  if (list === undefined) return undefined;
  if (list === null || (Array.isArray(list) && list.length === 0)) return [];
  const known = patternIds(list).filter((id) => (free && id === FREE_PATTERN) || manifest.patterns.some((pattern) => pattern.id === id));
  return known.length > 0 ? known : undefined;
}

/**
 * A deck's colours as they are kept: every colour in its stored form (six hex digits, upper
 * case, no `#`). A value that is no colour, or no set of colours at all, is kept as given, so
 * that the check of the deck says what is wrong with it. `undefined` when nothing is asked for.
 */
export function storedColors(colors: DeckColors | null | undefined): DeckColors | undefined {
  if (!colors) return undefined;
  if (typeof colors !== 'object' || Array.isArray(colors)) return colors;
  const stored: Record<string, unknown> = {};
  for (const [name, value] of Object.entries(colors as Record<string, unknown>)) {
    if (value === undefined || value === null || value === '') continue;
    if (name === 'replace' && typeof value === 'object' && !Array.isArray(value)) {
      const pairs = Object.entries(value as Record<string, unknown>).map(([from, to]) => [parseHex(from) ?? from, parseHex(to) ?? to] as const);
      if (pairs.length > 0) stored.replace = Object.fromEntries(pairs);
    } else {
      stored[name] = parseHex(value) ?? value;
    }
  }
  return Object.keys(stored).length > 0 ? (stored as DeckColors) : undefined;
}

function briefText(brief: Brief): string {
  return [`Brief:\n${brief.prompt.trim()}`, brief.instructions?.trim() ? `Also respect:\n${brief.instructions.trim()}` : '']
    .filter(Boolean)
    .join('\n\n');
}

/**
 * A wish for columns is a wish about how items sit. The catalogue says it for each pattern:
 * "side by side in one row", or "in lines of" a number for a grid.
 */
export const columnsRule = (asked: 'brief' | 'request'): string =>
  `When the ${asked} asks for a number of columns, or for items side by side, use a pattern that takes that number of items and whose items are "side by side in one row" or "in lines of" that number or more: either sets them on one line. Items "in lines of" a smaller number are broken over several lines. When the list has no such pattern, use the one that comes closest.`;

/** What a wish about the kind of slides means, in the planner's own terms. Shared by every prompt that chooses patterns. */
export const SLIDE_KIND_RULE =
  'A wish about the kind of slides (text only, no tables, only cards) decides which patterns may be used, and comes before variety. ' +
  '"Text only" means the patterns marked "Plain text.": a title with paragraphs or a list, and no cards, figures, tables, charts, icons or pictures. ' +
  'Using one pattern for many slides is right when the wish leaves no other.';

/**
 * The outline prompt. `limited` says the brief names the only patterns the deck may use: the list
 * shown is then that list, and the rules about variety and photos, which would pull against it,
 * are left out. `free` says free slides are on offer: the list then ends with their line, and
 * one rule says how such a slide names its parts. Without it the prompt is what it was before
 * there were free slides.
 */
function outlinePrompt(
  brief: Brief,
  catalog: Catalog,
  count: { min: number; max: number },
  density: ImageDensity,
  textScale = 1,
  limited = false,
  free?: FreeOffer,
  document?: { block: string; rules: string[] },
): string {
  const size = count.min === count.max ? `exactly ${count.min}` : `${count.min} to ${count.max}`;
  // The person's own list comes before the wish for text only: it names the patterns outright.
  const textOnly = brief.textOnly === true && !limited;
  const photos = free && density !== 'none' ? withFreePhotos(DENSITY_GUIDANCE[density]) : DENSITY_GUIDANCE[density];
  // A deck made from a document says it right after the brief, passage by passage, and is planned by rules of its own.
  const opening = document
    ? [`- ${DOCUMENT_OPENING}`]
    : // A deck of one or two slides has no room for an opening and a closing slide: its slides carry the content.
      count.max >= 3
      ? ['- A deck of three or more slides opens with an opening pattern and ends with a closing pattern when the list has them, unless the brief asks for a deck without them.']
      : [];
  return [
    briefText(brief),
    ...(document ? [document.block] : []),
    `Plan a deck of ${size} ${count.max === 1 ? 'slide' : 'slides'}. Choose one pattern for each slide from this list ("?" marks an optional slot):`,
    [describeCatalog(catalog, textScale), ...(free ? [freeLine(free.kinds, { comfortable: free.comfortable, alone: catalog.patterns.length === 0, plain: textOnly })] : [])].filter(Boolean).join('\n'),
    [
      `- ${wishesRule('brief')}`,
      ...(free ? [`- ${FREE_OUTLINE_RULE}`] : []),
      ...opening,
      '- For a pattern with a row of items, set "items" to the number the brief asks for, or else the number the content really has. Do not pad a list to fill a row.',
      `- ${columnsRule('brief')}`,
      `- ${wordsRule('brief')}`,
      ...(document ? document.rules.map((rule) => `- ${rule}`) : []),
      '- Choose a pattern only when the content can fill every required slot within its limits.',
      ...(textOnly
        ? ['- The deck is text only, and every pattern in the list is plain text. Using the same pattern for several slides in a row is expected.']
        : [`- ${SLIDE_KIND_RULE}`, ...(limited ? [] : [document ? `- ${DOCUMENT_VARY}` : '- Otherwise vary the patterns across the deck.', `- ${photos}`])]),
    ].join('\n'),
  ].join('\n\n');
}

/**
 * A slide in the list a writer is given: `s3: cards. Three goals`. A free slide says which of its
 * parts are what, and a slide that carries passages of a document has their text under its line.
 */
const slideLine = (job: WriteJob): string => `s${job.number}: ${job.pattern.def.id}${job.free ? ` (${describeParts(job.free.blocks)})` : ''}. ${job.purpose}${job.passage ? `\n${slideText(job.number, job.passage)}` : ''}`;

function writePrompt(brief: Brief, jobs: WriteJob[], context: CatalogContext): string {
  // A slide of many parts is written by calls of its own. The writer of the others is told that it is there.
  const here = jobs.filter((job) => !writtenApart(job));
  const apart = jobs.filter(writtenApart);
  // Slides with text from a document are written in its words: the rules for that come once, before the supplies.
  const fromDocument = here.some((job) => job.passage) ? [documentWriteRule(here.some((job) => job.filled))] : [];
  return [
    briefText(brief),
    `Write the content of each slide of this deck:\n${here.map(slideLine).join('\n')}`,
    ...(apart.length > 0 ? [`These slides of the deck are written apart, not here:\n${apart.map((job) => `s${job.number}: ${FREE_PATTERN} (${describeBlocks(job.free?.blocks ?? [])}). ${job.purpose}`).join('\n')}`] : []),
    ...fromDocument,
    ...suppliesText(here, context),
    'Fill every required property. Leave an optional property out when the slide is better without it.',
  ].join('\n\n');
}

/** The prompt of a call that writes some parts of one slide by themselves: the brief, the deck around the slide, then the task, and the slide's text from a document where it has one. */
function partsPrompt(brief: Brief, jobs: WriteJob[], job: WriteJob, task: string, context: CatalogContext): string {
  const deck = jobs.map((other) => `s${other.number}: ${other.pattern.def.id}${other.free ? ` (${describeBlocks(other.free.blocks)})` : ''}. ${other.purpose}`).join('\n');
  const fromDocument = job.passage ? [slideText(job.number, job.passage), documentWriteRule(job.filled !== undefined)] : [];
  return [briefText(brief), `The slides of the deck:\n${deck}`, task, ...fromDocument, ...suppliesText([job], context)].join('\n\n');
}

/** A deck before its content is written: the slides in order, each with a pattern and the point it makes. */
export interface DeckPlan {
  version: 1;
  title: string;
  /** The request, with everything needed to write the deck later. */
  brief: Brief;
  slides: OutlineEntry[];
  /** Layouts made for this deck (see `composeLayouts`). Slides use them by their ids. */
  layouts?: PatternDesign[];
  /**
   * For a deck made from a document (`Brief.document`): the language the document is written in,
   * as a BCP 47 tag, as the planner read it. Code sets the document into a slide word for word only
   * where that is the deck's language; in another, the writer translates it. Left out, as in a plan
   * stored before it was asked, the document is taken to be in the deck's language when it is
   * written in that language's script.
   */
  documentLanguage?: string;
}

interface Planning {
  layout: LayoutMode;
  language: string;
  direction: Direction;
  density: ImageDensity;
  fields: Record<string, string>;
  context: CatalogContext;
  catalog: Catalog;
  byId: Map<string, UsablePattern>;
  /** The lists of patterns the deck may use (see {@link allowedPatterns}). */
  lists: PatternLists;
  /** The template with the deck's own layouts built onto it. */
  template: Template;
  /** Free slides as they are on offer for a new slide: the kinds of part, and how many read well. `undefined` when none is offered. */
  free?: FreeOffer;
  /** A setting turns free slides off (see `freeTurnedOff`): no slide's text is set smaller, which a slide that is cut says. */
  freeOff?: true;
  /** The document the deck is made from, read into passages. `undefined` when there is none. */
  document?: DocumentPlan;
}

/**
 * The document a brief gives, read into passages: `undefined` for none, or a blank one. One that
 * is no text, or longer than a deck is made from, is a {@link PlanError} that says so. `written` is
 * the language it is written in, where the planner said.
 */
export function documentOf(brief: Pick<Brief, 'document'>, language: string, written?: string): DocumentPlan | undefined {
  const value: unknown = brief.document;
  if (value === undefined || value === null) return undefined;
  if (typeof value !== 'string') throw new PlanError('"document" must be text');
  if (value.trim() === '') return undefined;
  const chars = documentChars(value);
  if (chars > MOST_DOCUMENT_CHARS) throw new PlanError(`The document is ${chars} characters, and a deck is made from at most ${MOST_DOCUMENT_CHARS} (about 1,800 words): pass the part the deck is about`);
  const document = readDocument(value, language, written);
  return document.passages.length > 0 ? document : undefined;
}

interface PlanningStage {
  /** The layouts made for the deck so far. */
  layouts?: PatternDesign[] | undefined;
  /** Whether the brief is a new one, so that a mistake in its list of patterns is an error. */
  fresh?: boolean;
  /**
   * Whether the brief's list narrows what is offered. It does for a new plan. Later stages offer
   * everything the caller allows: a slide planned outside the list, on a later wish, must still
   * be written.
   */
  hard?: boolean;
  /** The plan at hand has a free slide, which is written and kept whatever is on offer for a new one. */
  hasFree?: boolean;
  /** The language the brief's document is written in, where the planner said (`DeckPlan.documentLanguage`). */
  documentLanguage?: string | undefined;
}

/**
 * Everything the planning stages derive from a brief and the caller's options, so they cannot
 * disagree. `loaded` is the template as it was loaded: the deck's own layouts (`stage.layouts`)
 * are built onto it here, and free slides are worked out on the template without them.
 */
function planning(loaded: Template, brief: Brief, options: PlanOptions, stage: PlanningStage = {}): Planning {
  // The brief is what the deck is about, in words. Anything else would reach the model as words nobody wrote.
  if (typeof brief?.prompt !== 'string') throw new PlanError('The brief has no prompt: "prompt" is the text of what the deck is about');
  const { layouts } = stage;
  const template = withLayouts(loaded, layouts);
  const { manifest } = template;
  const layout = options.layout ?? 'strict';
  const language = brief.language ?? manifest.language ?? 'en';
  const direction = brief.direction ?? directionOf(language);
  const document = documentOf(brief, language, stage.documentLanguage);
  // A share of pictures that is not one of the list ("many", or "toString", which every object has)
  // is the usual one. A deck made from a document is about its text, and has few.
  const density: ImageDensity = typeof brief.imageDensity === 'string' && Object.hasOwn(DENSITY_GUIDANCE, brief.imageDensity) ? brief.imageDensity : document ? 'low' : 'medium';
  const fields = brief.fields ?? {};

  const numbers = options.numberPatterns ?? 'from-brief';
  // The figures of a document are the brief's too.
  const figures = numbers === 'always' || (numbers === 'from-brief' && hasFigures(withoutColorCodes(`${brief.prompt} ${brief.instructions ?? ''}${document ? ` ${document.text}` : ''}`, brief.colors)));
  // A size that is not one of the list ("huge", or "__proto__", which every object has) is the usual one.
  const textScale = typeof brief.textSize === 'string' && Object.hasOwn(TEXT_SIZES, brief.textSize) ? TEXT_SIZES[brief.textSize] : 1;
  const context: CatalogContext = { layout, fields, icons: options.icons ?? [], assets: options.assets ?? [], imageDensity: density, figures, textScale, templateScale: manifest.textScale ?? 1 };
  // A template that has a pattern called `free` has that pattern and no free slides. In a new
  // brief the id may always be named otherwise: where no free slide can be planned, the reason is given
  // below. A stored list may name it where the template can carry free slides, which is looked at
  // only when a list does.
  const taken = manifest.patterns.some((pattern) => pattern.id === FREE_PATTERN);
  const known = !taken && namesFree(brief.patterns) && freeKnown(loaded);
  const lists = allowedPatterns(manifest, { app: options.patterns, person: brief.patterns }, stage.fresh === true, stage.fresh ? !taken : known);
  const allow = stage.hard ? (lists.limit ?? lists.allow) : lists.allow;
  if (allow) context.allow = allow;
  if (layouts?.length) {
    context.own = layouts.map((design) => design.id);
    // A deck's own layouts were made for it: a list of allowed patterns does not leave them out.
    if (context.allow) context.allow = [...context.allow, ...context.own];
  }
  const catalog = buildCatalog(manifest, context);
  // The wish for text only gives way to the person's own list, here as in the outline.
  const free = freeOnOffer(loaded, context, { direction, textScale }, { option: options.free, app: lists.allow, limit: lists.limit, textOnly: brief.textOnly === true && !lists.limit });
  if (stage.fresh && stage.hard && lists.limit) {
    // A pattern the person named and that cannot be filled here is said at once. Left out without
    // a word, the deck would be planned from the rest of the list, and still be called limited to it.
    const { limit } = lists;
    const unusable = catalog.excluded.filter((entry) => limit.includes(entry.id));
    if (limit.includes(FREE_PATTERN) && free.refusal) unusable.push({ id: FREE_PATTERN, reason: free.refusal });
    if (unusable.length > 0) {
      throw new PlanError(
        `The deck cannot be limited to ${unusable.length === 1 ? 'a pattern' : 'patterns'} that cannot be filled here: ${unusable.map((entry) => `${entry.id} ${entry.reason}`).join('; ')}`,
      );
    }
  }
  // A deck may be planned from free slides alone, and a plan that has one is written whatever else can be used.
  if (catalog.patterns.length === 0 && !free.offer && !stage.hasFree) {
    const reasons = catalog.excluded.map((entry) => `${entry.id} ${entry.reason}`).join('; ');
    throw new PlanError(`No pattern of template "${manifest.name}" can be used${reasons ? `: ${reasons}` : ''}`);
  }
  const planned: Planning = { layout, language, direction, density, fields, context, catalog, byId: new Map(catalog.patterns.map((pattern) => [pattern.def.id, pattern])), lists, template };
  if (free.offer) planned.free = free.offer;
  else if (freeTurnedOff(manifest, { option: options.free, app: lists.allow, limit: lists.limit })) planned.freeOff = true;
  if (document) planned.document = document;
  return planned;
}

/** What the answer of a model may say of free slides, when a prompt speaks of them. */
interface FreeAnswer extends FreeReading {
  /** Whether free slides are on offer for a new slide, and not only kept where a plan has them. */
  offered: boolean;
}

/**
 * Reads an outline from a model's answer. `free` says free slides may be among its slides, and
 * what their parts may be. `final` says no further answer will be asked for: what is wrong with
 * the parts of a free slide is then put right here and is no problem any more. `passages` is the
 * number of passages of the document a deck is made from: each slide then says which it carries.
 */
function readOutline(answer: unknown, byId: Map<string, UsablePattern>, free?: FreeAnswer, final = false, passages?: number): { title: string; entries: OutlineEntry[]; problem?: string; hint?: string } {
  const raw = answer as { title?: unknown; slides?: unknown } | null;
  const list = Array.isArray(raw?.slides) ? raw.slides : [];
  const problems: string[] = [];
  let counts = false;
  const entries = list.map((item, index) => {
    const raw = item as { pattern?: unknown; purpose?: unknown; items?: unknown; words?: unknown; passages?: unknown };
    const entry: OutlineEntry = { pattern: cleanText(raw?.pattern), purpose: cleanText(raw?.purpose) };
    const pattern = byId.get(entry.pattern);
    if (pattern) {
      const count = readItems(pattern.def, raw?.items);
      if (count.problem) {
        problems.push(`slide ${index + 1}: ${count.problem}`);
        counts = true;
      }
      if (count.items !== undefined) entry.items = count.items;
    } else if (free && entry.pattern === FREE_PATTERN) {
      // What a free slide says in "shows" is all there is to it; on any other slide "shows" is passed over.
      const shown = readShows(raw, index + 1, entry.purpose, free, final);
      entry.blocks = shown.blocks;
      if (shown.arrange) entry.arrange = shown.arrange;
      if (shown.problem) problems.push(shown.problem);
    }
    const words = readWords(raw?.words);
    if (words !== undefined) entry.words = words;
    if (passages !== undefined) {
      const own = readPassageNumbers(raw?.passages, passages);
      if (own.length > 0) entry.passages = own;
    }
    return entry;
  });
  const unknown = entries.filter((entry) => !byId.has(entry.pattern) && !isFreeEntry(entry)).map((entry) => entry.pattern || '(empty)');
  const outline = { title: cleanText(raw?.title), entries };
  if (entries.length === 0) return { ...outline, problem: 'the answer contained no slides' };
  if (unknown.length > 0) return { ...outline, problem: `these patterns are not in the list: ${[...new Set(unknown)].join(', ')}` };
  if (free && !free.offered) {
    // The plan keeps the free slides it has, and gets no new one: free slides are not on offer here.
    const had = (free.old ?? []).filter(isFreeEntry).length;
    const has = entries.filter(isFreeEntry).length;
    if (has > had) return { ...outline, problem: `the layout free is only for the ${had === 1 ? 'slide' : `${had} slides`} of the plan that ${had === 1 ? 'has' : 'have'} it already, and the answer gives it to ${has}` };
  }
  if (problems.length > 0) {
    // A row that cannot take the number asked of it is what a free slide is for.
    return { ...outline, problem: problems.join('; '), ...(counts && free?.offered ? { hint: ', or use free for a slide that needs more items than a pattern takes' } : {}) };
  }
  return outline;
}

/**
 * First stage: decides which pattern each slide uses and the point it makes. One model call.
 * The result is small and quick to produce, so it is the thing to show a person and let them change.
 */
export async function planOutline(template: Template, brief: Brief, options: PlanOptions): Promise<DeckPlan> {
  // A brief that names the only patterns to use is taken at its word: nothing else is offered.
  const planned = planning(template, brief, options, { fresh: true, hard: true });
  const { language, density, context, free } = planned;
  const limited = planned.lists.limit !== undefined;
  const catalog = outlineCatalog(planned, brief);
  // A text-only deck can be planned from free slides of text where the template has no plain text pattern.
  if (catalog.patterns.length === 0 && !free) {
    throw new PlanError(`Template "${template.manifest.name}" has no plain text pattern, so a text-only deck cannot be planned from it`);
  }
  const byId = new Map(catalog.patterns.map((pattern) => [pattern.def.id, pattern]));
  const given = typeof brief.slideCount === 'number' ? { min: brief.slideCount, max: brief.slideCount } : brief.slideCount;
  // A deck made from a document has as many slides as its text takes, where the brief names no number.
  const doc = planned.document;
  const rooms = doc ? documentRooms(template, planned, catalog) : undefined;
  const own = doc && rooms ? Math.max(1, Math.ceil(doc.chars / rooms.target)) : 0;
  const count = given ?? (doc ? { min: 1, max: Math.min(MOST_DOCUMENT_SLIDES, Math.max(10, own + 4)) } : { min: 6, max: 10 });
  // Choosing a pattern for each slide needs little reasoning, and is much quicker with little.
  const ask = asker(options.llm, systemPrompt(language, 'brief', doc !== undefined), options.signal, 'low');
  const reading: FreeAnswer | undefined = free ? { ...freeReading(free.kinds, context), offered: true } : undefined;
  const passages = doc?.passages.length;
  const asksCount = brief.slideCount === undefined;
  const schema = outlineSchema(catalog, count, free, passages !== undefined ? { passages, askedSlides: asksCount } : undefined);
  // A free slide names its parts, which is a little more to say for each slide; a slide of a deck
  // made from a document names its passages, and the answer those it leaves out, the slides the
  // brief asks for and the language of the document.
  const room = planRoom(count.max) + (free ? FREE_ENTRY_ROOM * count.max : 0) + (passages !== undefined ? 8 * count.max + 2 * passages + DOCUMENT_ROOM : 0);

  options.onProgress?.({ stage: 'outline' });
  const told = doc && rooms ? { block: documentBlock(doc), rules: documentRules({ slides: own, chars: rooms.target, given }) } : undefined;
  const request = outlinePrompt(brief, catalog, count, density, budgetScale(context), limited, free, told);
  const read = (answer: unknown, final: boolean) => readOutline(answer, byId, reading, final, passages);
  let answer = await ask(request, schema, room);
  let outline = read(answer, false);
  // Of a deck made from a document, an answer that places little of it, or that does not have the
  // number of slides it says the brief asks for, is asked for again with the problems of patterns.
  const missed = passages !== undefined ? documentProblems(outline.entries, answer, passages, asksCount, given) : [];
  if (outline.problem || missed.length > 0) {
    const problem = [outline.problem, ...missed].filter(Boolean).join('; ');
    const placing = passages !== undefined ? ', with every passage on one slide or in "left_out"' : '';
    const retry = outline.problem
      ? `${request}\n\nYour previous answer could not be used: ${problem}. Answer again using only the listed patterns${outline.hint ?? ''}${placing}.`
      : `${request}\n\nYour previous answer could not be used: ${problem}. Answer again${placing}.`;
    const first = { answer, outline };
    answer = await ask(retry, schema, room);
    outline = read(answer, true);
    if (outline.problem) {
      // How a document is placed is put right in code: an answer that was asked again for that alone is used as it was.
      if (first.outline.problem) throw new PlanError(`The model did not produce a usable outline: ${outline.problem}`);
      ({ answer, outline } = first);
    }
  }
  const lp = lengthPlanning(template, planned, catalog);
  let slides: OutlineEntry[];
  let documentLanguage: string | undefined;
  if (doc && rooms) {
    // The passages are put right in code, then each slide gets the room for its own, on that slide
    // where the number of slides is fixed, over more slides where it is not and they need them.
    const said = answer as { left_out?: unknown; asked_slides?: unknown; document_language?: unknown } | null;
    const asked = asksCount ? readAskedSlides(said?.asked_slides) : undefined;
    const fixed = isFixedCount(brief.slideCount) || asked !== undefined;
    const most = fixed ? outline.entries.length : asksCount ? MOST_DOCUMENT_SLIDES : count.max;
    // The language the document is written in, as the answer says: code sets it into a slide as it
    // is only where that is the deck's language. One the text's own script belies is not kept.
    documentLanguage = readLanguageTag(said?.document_language, doc.text);
    const read = documentLanguage ? inLanguage(doc, language, documentLanguage) : doc;
    const sizeOf = (numbers: readonly number[]) => needOf(read, numbers, language);
    const settled = settlePassages(outline.entries, readPassageNumbers(said?.left_out, read.passages.length), read, {
      carries: carriesText(byId, context, read),
      fixed,
      slides: own,
      room: Math.max(0, most - outline.entries.length),
      textHome: textEntryOf(lp, catalog, context, outline.entries[0]?.pattern ?? FREE_PATTERN),
      sizeOf,
      ...(fixed && outline.entries.length <= 3 ? { keepEnd: endKeeps(read, lp) } : {}),
    }).entries;
    const placed = fitDocument(settled, read, lp, { most: fixed ? settled.length : most, fixed, target: rooms.target });
    slides = fitOutline(placed, lp, (entry) => (entry.passages?.length ?? 0) > 0);
  } else {
    // A slide that was asked an amount of text gets the room for it, on that slide.
    slides = fitOutline(outline.entries, lp);
  }
  const plan: DeckPlan = { version: 1, title: outline.title, brief, slides };
  if (documentLanguage) plan.documentLanguage = documentLanguage;
  // A layout made for the deck would be one the person did not name. And where free slides are on
  // offer, a free slide is what the deck gets for content that no layout shows. A deck made from a
  // document carries its text, which the template's own layouts and free slides are planned for.
  if (!options.composeLayouts || limited || free || doc) return plan;
  options.onProgress?.({ stage: 'layouts' });
  return (await composeForPlan(template, plan, catalog.patterns, { llm: options.llm, max: options.composeLayouts, ...(options.signal ? { signal: options.signal } : {}) })).plan;
}

/**
 * What finding room for an asked length works with in a plan (see length.ts): the layouts of the
 * list on offer, sized for their items, and free slides where they are on offer, worked out on
 * the template as it was loaded (`loaded`). A slide of a layout that becomes a free slide keeps
 * its chart, its table and its photo, where the free slide can take one.
 */
function lengthPlanning(loaded: Template, planned: Planning, catalog: Catalog): LengthPlanning {
  const { context, direction, free, language } = planned;
  const byId = new Map(catalog.patterns.map((pattern) => [pattern.def.id, pattern]));
  const plain = catalog.patterns.filter((pattern) => isPlainText(pattern.def)).map((pattern) => pattern.def.id);
  const kinds = new Set<FreeBlockKind>(free?.kinds ?? []);
  const written: NonNullable<LengthPlanning['written']> = (blocks, arrange, lengths, paragraphs, asIs) => {
    const made = freeWritePattern(loaded, { direction, textScale: context.textScale }, { blocks, arrange, lengths, ...(paragraphs ? { paragraphs } : {}), ...(asIs ? { asIs } : {}) }, context);
    return 'pattern' in made ? made.pattern : undefined;
  };
  return {
    context,
    language,
    layout: (id, items) => {
      const usable = byId.get(id);
      if (!usable || items === undefined) return usable;
      const sized = planPattern(usable.def, context, undefined, items);
      return 'pattern' in sized ? sized.pattern : undefined;
    },
    plainText: (id) => plain.includes(id),
    plain,
    free: kinds.has('text') ? written : undefined,
    written,
    freeParts: (id) => freePartsOf(byId.get(id)?.def, kinds),
  };
}

/**
 * The parts a free slide is made of when a slide of this layout becomes one for a long text,
 * besides the text: its charts, tables and photos, of the kinds on offer. Anything else of the
 * layout (cards, figures, icons) is written into the text.
 */
export function freePartsOf(def: PatternDef | undefined, kinds: ReadonlySet<FreeBlockKind>): FreeBlockKind[] {
  return (def?.slots ?? []).flatMap((slot): FreeBlockKind[] => {
    if (slot.empty === 'drop' || slot.empty === 'keep') return [];
    if (slot.kind === 'chart' && kinds.has('chart')) return ['chart'];
    if (slot.kind === 'table' && kinds.has('table')) return ['table'];
    if (slot.kind === 'image' && slot.imageKind === 'photo' && kinds.has('photo')) return ['photo'];
    return [];
  });
}

/**
 * The slides of an outline, with room given to the text asked for some of them. Such a slide
 * keeps its layout when that holds the text, or moves to one that does (see `homeFor`): a free
 * slide of that text alone set smaller where it needs it (`arrange.dense`). A text is never spread
 * over more slides. A slide that still holds less is cut to what it holds: the plan says so
 * (`planTextFit`), and the deck when it is written. `settled` names the slides fitted already,
 * which a change to a plan passes as they are. `textFirst` says which slides may give up their
 * other parts for a text they would leave under half of its words (see `homeFor`): every slide
 * where left out, as in a new plan; in a change to a plan, only where the plan had those parts.
 */
export function fitOutline(entries: OutlineEntry[], planning: LengthPlanning, settled?: (entry: OutlineEntry) => boolean, textFirst?: (entry: OutlineEntry) => boolean): OutlineEntry[] {
  if (!entries.some((entry) => entry.words !== undefined)) return entries;
  return entries.map((entry) => {
    if (entry.words === undefined || settled?.(entry)) return entry;
    const current: TextHome = isFreeEntry(entry) ? { blocks: entry.blocks, arrange: entry.arrange } : { pattern: entry.pattern, items: entry.items };
    const found = homeFor(current, entry.words, planning, textFirst?.(entry) ?? true);
    return found.home === current ? entry : placeEntry(entry, found.home);
  });
}

/** Whether a pattern has a table slot, for a table of a slide's passages. */
const hasTableSlot = (pattern: UsablePattern | undefined): boolean => pattern?.slots.some((planned) => planned.slot.kind === 'table') ?? false;

/**
 * Whether a slide of a plan has a table for a table of its passages: a free slide with a table
 * part, or a layout with a table slot. On a slide without, a table's rows are running text.
 */
const showsTable = (entry: OutlineEntry, byId: ReadonlyMap<string, UsablePattern>): boolean =>
  byId.has(entry.pattern) ? hasTableSlot(byId.get(entry.pattern)) : isFreeEntry(entry) && entryBlocks(entry).includes('table');

/** An outline entry given another home for its text: a layout with its items, or a free slide of these parts. */
function placeEntry(entry: OutlineEntry, home: TextHome): OutlineEntry {
  const { pattern: _pattern, items: _items, blocks: _blocks, arrange: _arrange, ...rest } = entry;
  if ('blocks' in home) return { pattern: FREE_PATTERN, ...rest, blocks: home.blocks, ...(home.arrange ? { arrange: home.arrange } : {}) };
  return { pattern: home.pattern, ...rest, ...(home.items !== undefined ? { items: home.items } : {}) };
}

/**
 * The output tokens an outline made from a document, or a change to one, takes for what it says
 * besides its slides: the passages it leaves out, the slides the brief asks for, and the language
 * of the document.
 */
const DOCUMENT_ROOM = 40;

/**
 * Whether a slide can carry passages where it stands first or last: a free slide, a layout with
 * room for running text, or a layout with a table that is given only tables of the document
 * (`doc`), which it shows in its table, as it does in the middle of a deck. An opening or a closing
 * slide that cannot gives its passages to the slide next to it.
 */
const carriesText =
  (byId: ReadonlyMap<string, UsablePattern>, context: CatalogContext, doc?: DocumentPlan) =>
  (entry: { pattern: string; blocks?: unknown; passages?: readonly number[] }): boolean => {
    if (isFreeEntry(entry) || !byId.has(entry.pattern)) return true;
    const pattern = byId.get(entry.pattern) as UsablePattern;
    if (textRoom(pattern, context) > 0) return true;
    const passages = entry.passages ?? [];
    return doc !== undefined && passages.length > 0 && passages.every((number) => doc.passages[number - 1]?.kind === 'table') && hasTableSlot(pattern);
  };

/**
 * How much of a document one slide holds, in characters: `comfortable` at the template's own text
 * size, and `target`, what an outline is told to give a slide. Of a free slide of text where those
 * are on offer, as the engine sets one out, nine tenths of it in round fifties; else of the layout
 * of plain text that holds most.
 */
function documentRooms(loaded: Template, planned: Planning, catalog: Catalog): { comfortable: number; target: number } {
  const { context, direction, free } = planned;
  const room = free?.kinds.includes('text') ? freeTextRoom(loaded, { direction, textScale: context.textScale }) : undefined;
  if (room && room.comfortable > 0) return { comfortable: room.comfortable, target: Math.max(50, Math.floor((0.9 * room.comfortable) / 50) * 50) };
  const plain = catalog.patterns.filter((pattern) => isPlainText(pattern.def)).map((pattern) => heldBy(pattern, context));
  const comfortable = Math.max(0, ...plain) || 500;
  return { comfortable, target: comfortable };
}

/**
 * What an outline of a deck made from a document is asked again for: too few passages placed, not
 * the number of slides it says the brief asks for, or a slide of three or fewer that carries none.
 * `given` is the number of slides the brief gives.
 */
function documentProblems(entries: OutlineEntry[], answer: unknown, passages: number, asksCount: boolean, given?: { min: number; max: number }): string[] {
  const said = answer as { left_out?: unknown; asked_slides?: unknown } | null;
  const leftOut = readPassageNumbers(said?.left_out, passages);
  const gap = grossGap(entries, leftOut, passages);
  const asked = asksCount ? readAskedSlides(said?.asked_slides) : undefined;
  const count = asked !== undefined && entries.length > 0 && entries.length !== asked ? `the brief asks for ${asked} ${asked === 1 ? 'slide' : 'slides'} and the answer has ${entries.length}: answer with exactly ${asked}` : undefined;
  const bare = bareSlides(entries, asked ?? (given && given.min === given.max ? given.min : undefined), passages - leftOut.length, 'brief');
  return [gap, count, bare].filter((problem): problem is string => problem !== undefined);
}

/**
 * What an answer of three slides or fewer, where their number is fixed (`count`), is asked again
 * for: a slide that carries no passage, where there are passages enough (`placeable`) for each.
 * Such a deck carries the document on every slide unless the brief or the request asks for an
 * opening or a closing slide, which only the model can tell: an answer that keeps one when asked
 * again stands.
 */
function bareSlides(entries: readonly OutlineEntry[], count: number | undefined, placeable: number, asked: 'brief' | 'request'): string | undefined {
  if (count === undefined || count > 3 || entries.length > 3 || placeable < entries.length) return undefined;
  const bare = entries.flatMap((entry, at) => ((entry.passages?.length ?? 0) > 0 ? [] : [at + 1]));
  if (bare.length === 0) return undefined;
  const which = bare.length === 1 ? `slide ${bare[0]} carries` : `slides ${bare.join(' and ')} carry`;
  return `${which} no passage, and a deck of three slides or fewer carries the document on every slide unless the ${asked} asks for an opening or a closing slide`;
}

/**
 * The pattern of a slide of text that settling the passages of an answer adds, or that an opening
 * or a closing slide becomes (see `settlePassages`): as {@link textHomeOf}, and where neither a
 * layout of plain text nor a free slide is on offer, the layout of most running text, else `fallback`.
 */
const textEntryOf =
  (planning: LengthPlanning, catalog: Catalog, context: CatalogContext, fallback: string) =>
  (): { pattern: string; blocks?: ['text'] } => {
    const home = textHomeOf(planning);
    if (home) return 'blocks' in home ? { pattern: FREE_PATTERN, blocks: ['text'] } : { pattern: home.pattern };
    const roomiest = [...catalog.patterns].sort((a, b) => textRoom(b, context) - textRoom(a, context))[0];
    return { pattern: roomiest?.def.id ?? fallback };
  };

/**
 * Where a slide added for a text goes: the layout of plain text on offer that holds least, else a
 * free slide of text. `undefined` where neither is on offer.
 */
function textHomeOf(planning: LengthPlanning): TextHome | undefined {
  const rooms = planning.plain
    .flatMap((id) => {
      const pattern = planning.layout(id);
      return pattern ? [{ id, room: heldBy(pattern, planning.context) }] : [];
    })
    .filter((one) => one.room > 0)
    .sort((a, b) => a.room - b.room);
  if (rooms[0]) return { pattern: rooms[0].id };
  return planning.free ? { blocks: ['text'] } : undefined;
}

/** Whole slides shared out in proportion to what each asks, by the largest remainders, never more than it asks. */
function sharedOut(wants: readonly number[], spare: number): number[] {
  const total = wants.reduce((sum, want) => sum + want, 0);
  if (total <= spare) return [...wants];
  const exact = wants.map((want) => (spare * want) / total);
  const given = exact.map(Math.floor);
  let left = spare - given.reduce((sum, one) => sum + one, 0);
  const order = exact.map((value, at) => ({ at, rest: value - Math.floor(value) })).sort((a, b) => b.rest - a.rest || a.at - b.at);
  for (const { at } of order) {
    if (left <= 0) break;
    if ((given[at] as number) < (wants[at] as number)) {
      given[at] = (given[at] as number) + 1;
      left--;
    }
  }
  return given;
}

/** What fitting the passages of a document to the slides of a plan needs besides them. */
export interface DocumentFit {
  /** The most slides the plan may have. */
  most: number;
  /** The number of slides is fixed: each slide stays one slide, and none is added. */
  fixed: boolean;
  /** The characters of the document a slide is planned to hold, where the number of slides is not fixed. */
  target?: number;
  /** Slides fitted already, which are passed as they are, unless the passages are shared again (`balance`). */
  settled?: (entry: OutlineEntry) => boolean;
  /**
   * With a fixed number of slides: where a slide holds less than its passages and the slides carry
   * the passages in runs, the passages are shared again over those slides, fitted already or not,
   * and where one still holds less, over the slides of text that carry none too. Defaults to `fixed`.
   */
  balance?: boolean;
  /**
   * Whether the text of a slide comes before its other parts (see `fitDocument`). Every slide's
   * does where left out; a change to a plan keeps a part a change asks for, or asked for before.
   */
  textFirst?: (entry: OutlineEntry) => boolean;
}

/** Where the text of a slide of a plan goes as it stands: its layout with its items, or a free slide of its parts. */
const homeOfEntry = (entry: OutlineEntry): TextHome => (isFreeEntry(entry) ? { blocks: entry.blocks, arrange: entry.arrange } : { pattern: entry.pattern, items: entry.items });

/** Whether a home has a table for a table of its passages. On one without, the table's rows are running text. */
const tableInHome = (home: TextHome, planning: LengthPlanning): boolean => ('blocks' in home ? home.blocks.includes('table') : hasTableSlot(planning.layout(home.pattern)));

/**
 * The person's text comes first (see `fitDocument`): what the text of some passages needs alone,
 * where a table of them is running text, and words asked for the slide ask for it shorter as they
 * would. Where free slides are not on offer, a layout of plain text is tried for it (`plain`).
 */
const textFirstOf = (doc: DocumentPlan, numbers: readonly number[], language: string | undefined, words?: number): TextFirst => ({
  need: slideDocument(doc, numbers, words, language, false)?.need ?? needOf(doc, numbers, language, false),
  shape: shapeOf(doc, numbers, false),
  below: 1,
  plain: true,
});

/**
 * Passages shorter than this, in characters, are a heading or a title line, which an opening or a
 * closing slide shows: as a slot of fewer characters is no running text (length.ts).
 */
const HEADING_CHARS = 200;

/**
 * In a deck of three slides or fewer whose number is fixed, whether an opening or a closing slide
 * that shows no running text keeps the passages the answer gave it, as a slide of text (see
 * `settlePassages`): where they are running text, more than a title line, and the slide next to
 * it would not hold them beside its own, so that giving them to it would cut them. A cover given
 * the document's title line stays a cover, and so does one whose passages fit next door. A layout
 * with a table given only tables of the document is never asked: it carries them where it is, in
 * its table (`carriesText`).
 */
function endKeeps(doc: DocumentPlan, planning: LengthPlanning): NonNullable<SettleOptions['keepEnd']> {
  const { language } = planning;
  return (end, to) => {
    if (needOf(doc, end.passages, language, false) < HEADING_CHARS) return false;
    if (!to) return true;
    const numbers = [...to.passages, ...end.passages].sort((a, b) => a - b);
    const home = homeOfEntry(to as OutlineEntry);
    const tables = tableInHome(home, planning);
    const found = homeForChars(home, needOf(doc, numbers, language, tables), planning, { shape: shapeOf(doc, numbers, tables), dense: true, join: true, first: textFirstOf(doc, numbers, language) });
    return !found.holds;
  };
}

/**
 * The slides of an outline made from a document, each given the room for its passages. A slide
 * keeps its layout where that holds them, or moves to one that does (see `homeForChars`): a free
 * slide of that text, set smaller where it needs it (`arrange.dense`). The person's text comes
 * first: a slide that would be cut beside cards, a chart, a table, a list or anything else it
 * shows is its text alone, set dense, where free slides are on offer and that keeps more of the
 * text. Its other parts give way, as a photo does. Where free slides are not on offer, such a
 * slide is a layout of plain text where that keeps more of it. `fit.textFirst` says which slides
 * may change so: all may in a new plan. Words asked for a slide of at least four fifths of the
 * words of its passages ask for no shorter text, and are dropped. A
 * table of the passages goes into a table where the slide has one, and needs no room of its text;
 * on a slide without, its rows are running text, and need their room. Where passages are shared
 * out again, the table of a slide goes to the slide that carries the table of its passages.
 *
 * Where the number of slides is not fixed, a slide whose passages are much more than a slide holds
 * at the template's size goes over more slides, each with its own run of passages, as many as the
 * plan has room for; a slide asked to be shorter never does. Where it is fixed, no slide is added:
 * a slide that holds less is set smaller, and when one still holds less the passages are shared
 * again over the slides that carry them, where those carry them in order. Where one holds less even
 * so, the slides of text that carry none take their share too: no slide of the number asked for
 * stands empty while another is cut. A slide asked to be shorter than its passages keeps them, and
 * the passages are shared again only between such slides. A slide that still holds less is cut,
 * and the plan says so (`planTextFit`).
 */
export function fitDocument(entries: OutlineEntry[], doc: DocumentPlan, planning: LengthPlanning, fit: DocumentFit): OutlineEntry[] {
  const { language } = planning;
  const sizeOf = (numbers: readonly number[]) => needOf(doc, numbers, language);
  const homeOf = homeOfEntry;
  const tableIn = (home: TextHome): boolean => tableInHome(home, planning);
  /** Whether the text of each slide of the outline may come before its other parts. */
  const firstAt = entries.map((entry) => fit.textFirst?.(entry) ?? true);
  /**
   * An entry homed for some passages, from the home it has (`home`): it stays as it is where that
   * holds them. With `textFirst` false, what else the slide shows stays (see `textFirstOf`).
   */
  const homed = (entry: OutlineEntry, home: TextHome, numbers: readonly number[], need = needOf(doc, numbers, language, tableIn(home)), textFirst = true): { entry: OutlineEntry; holds: boolean } => {
    const found = homeForChars(home, need, planning, { shape: shapeOf(doc, numbers, tableIn(home)), dense: true, join: true, ...(textFirst ? { first: textFirstOf(doc, numbers, language, entry.words) } : {}) });
    return { entry: found.home === home ? entry : placeEntry(entry, found.home), holds: found.holds };
  };
  const isTable = (number: number): boolean => doc.passages[number - 1]?.kind === 'table';
  const text = textHomeOf(planning);
  /**
   * Where passages that a slide with a table carried (its home `tabled`, `undefined` where none
   * did) are shared out again, the home of one slide of them (`from`): the table goes with the
   * table of the passages. A slide whose passages (`numbers`) have it gets a table, and the slide
   * that had it loses its table where they do not. A layout with a table goes whole to the slide
   * with the table, and one that loses it starts from a slide of text; a free slide's table is a
   * part of the slide with the table.
   */
  const tableFor = (from: TextHome, tabled: TextHome | undefined, numbers: readonly number[]): TextHome => {
    const table = numbers.some(isTable);
    if (!tabled || tableIn(from) === table) return from;
    if (table) return !('blocks' in tabled) ? tabled : 'blocks' in from ? { ...from, blocks: [...from.blocks, 'table'] } : { blocks: ['text', 'table'] };
    if (from !== tabled) return from;
    return 'blocks' in from ? { ...from, blocks: from.blocks.filter((kind) => kind !== 'table') } : (text ?? from);
  };
  /** The home with a table of a slide that carried a table of its passages. */
  const tabledOf = (home: TextHome, numbers: readonly number[]): TextHome | undefined => (tableIn(home) && numbers.some(isTable) ? home : undefined);
  const balance = fit.balance ?? fit.fixed;
  /**
   * Whether a slide is a slide of text, which passages shared again may go into where it carries
   * none: a free slide of texts and photos, or a layout of plain text with room for running text.
   * A chart, a table, a list or a layout of any other kind is about what it shows, and stays so.
   */
  const ofText = (entry: OutlineEntry): boolean => {
    if (isFreeEntry(entry)) return densable(entryBlocks(entry), true);
    const pattern = planning.layout(entry.pattern, entry.items);
    return planning.plainText(entry.pattern) && pattern !== undefined && textRoom(pattern, planning.context) > 0;
  };
  const carried = entries.map((entry) => {
    // Passages shared again are shared over every slide that carries some, in order: a slide
    // fitted already is no gap they are moved across.
    if (!balance && fit.settled?.(entry)) return { entry };
    const home = homeOf(entry);
    const carry = slideDocument(doc, entry.passages, entry.words, language, tableIn(home));
    if (!carry) return { entry };
    const base: OutlineEntry = { ...entry, passages: carry.own };
    if (carry.asked === undefined) delete base.words;
    return { entry: base, carry, home };
  });

  if (!fit.fixed) {
    const target = fit.target ?? Number.POSITIVE_INFINITY;
    // A slide is cut into parts by all its text: a table may fall on a part with no table, where it is running text.
    const allText = (numbers: readonly number[]) => needOf(doc, numbers, language, false);
    const parts = carried.map(({ carry }) => {
      if (!carry || carry.asked !== undefined || carry.need <= target) return 1;
      const atoms = atomsOf(doc, carry.own, allText);
      let k = 1;
      while (k < atoms.length && minMax(atoms, k) > target) k++;
      return k;
    });
    const more = sharedOut(
      parts.map((k) => k - 1),
      Math.max(0, fit.most - entries.length),
    );
    return carried.flatMap(({ entry, carry, home }, at): OutlineEntry[] => {
      if (!carry || !home) return [entry];
      const k = 1 + (more[at] as number);
      if (k === 1) return [homed(entry, home, carry.own, carry.need, firstAt[at]).entry];
      const groups = packPassages(atomsOf(doc, carry.own, allText), k);
      return groups.map((group, part) => {
        const own: OutlineEntry = { ...entry, passages: group, purpose: partPurpose(entry.purpose, part + 1, groups.length) };
        delete own.words;
        // A later part starts from a slide of text: a photo or a chart is not shown twice. The table
        // of the slide goes to the part that carries the table of its passages.
        const from = tableFor(part === 0 ? home : (text ?? home), tabledOf(home, carry.own), group);
        const tables = tableIn(from);
        const found = homeForChars(from, needOf(doc, group, language, tables), planning, { shape: shapeOf(doc, group, tables), dense: true, join: true, ...(firstAt[at] ? { first: textFirstOf(doc, group, language) } : {}) });
        return found.home === home ? own : placeEntry(own, found.home);
      });
    });
  }

  // Sharing the passages again comes before what else a slide shows gives way to its text: where
  // that makes room, a slide keeps its chart, its table or its cards.
  const placed = carried.map(({ entry, carry, home }) => (carry && home ? homed(entry, home, carry.own, carry.need, false) : { entry, holds: true }));
  const carriers = carried.flatMap((one, at) => (one.carry ? [at] : []));
  const lists = carriers.map((at) => (carried[at]?.carry as SlideDocument).own);
  // The slides carry the passages in runs, one after another: each a run with no gap, after the run before.
  const inRuns = lists.every((list, at) => list.every((number, place) => place === 0 || number === (list[place - 1] as number) + 1) && (at === 0 || (list[0] as number) > ((lists[at - 1] as number[]).at(-1) as number)));
  /**
   * Shares the passages of the slides from `from` up to `to` again, where one of them is cut: over
   * the slides that carry them first. Only where one is cut even so do the slides of text there that
   * carry none take their share too, so that none stands empty while another is cut.
   */
  const shareAgain = (from: number, to: number): void => {
    const own = carriers.filter((at) => at >= from && at < to);
    if (own.every((at) => placed[at]?.holds)) return;
    const atoms = atomsOf(doc, own.flatMap((at) => (carried[at]?.carry as SlideDocument).own), sizeOf);
    const bare = carried.flatMap((one, at) => (at >= from && at < to && !one.carry && ofText(one.entry) ? [at] : []));
    // The table of a slide goes with the table of its passages, to the slide that now carries it.
    const tabled = own.map((at) => carried[at] as { home: TextHome; carry: SlideDocument }).find(({ home, carry }) => tabledOf(home, carry.own))?.home;
    const share = (over: readonly number[]): void =>
      packPassages(atoms, over.length).forEach((group, place) => {
        const at = over[place] as number;
        const one = carried[at] as { entry: OutlineEntry; home?: TextHome };
        const home = one.home ?? homeOf(one.entry);
        // A slide that carried nothing is about the passages it now carries.
        const entry: OutlineEntry = one.home ? one.entry : { ...one.entry, purpose: purposeOf(doc.passages, group) };
        if (!one.home) delete entry.words;
        const from = tableFor(home, tabled, group);
        placed[at] = homed({ ...(from === home ? entry : placeEntry(entry, from)), passages: group }, from, group, undefined, false);
      });
    if (atoms.length < own.length) return;
    share(own);
    if (bare.length > 0 && atoms.length >= own.length + bare.length && own.some((at) => !placed[at]?.holds)) share([...own, ...bare].sort((a, b) => a - b));
  };
  if (balance && inRuns) {
    // A slide asked for fewer words than its passages keeps them: it is to be shorter than they are.
    // So does a slide with a table that carries only tables of the document, which it shows in its
    // table, at an end of the deck as in its middle. The passages are shared again over the slides
    // between two such slides.
    const keeps = (at: number): boolean => {
      const one = carried[at] as { carry?: SlideDocument; home?: TextHome };
      return one.carry?.asked !== undefined || (one.carry !== undefined && one.home !== undefined && one.carry.own.every(isTable) && tableIn(one.home));
    };
    let from = 0;
    for (const fence of [...carriers.filter(keeps), entries.length]) {
      shareAgain(from, fence);
      from = fence + 1;
    }
  }
  // A slide still cut is its text alone, set dense, or without free slides a layout of plain text,
  // where that keeps more of its text: the person's text comes before the other parts of its slide.
  return placed.map((one, at) => {
    if (one.holds || !firstAt[at]) return one.entry;
    const home = homeOf(one.entry);
    const carry = slideDocument(doc, one.entry.passages, one.entry.words, language, tableIn(home));
    return carry ? homed(one.entry, home, carry.own, carry.need).entry : one.entry;
  });
}

/** What a slide of a plan carries before anything is written: for each slide given words or text from a document. */
export interface SlideTextFit {
  /** The slide's number, from 1. */
  slide: number;
  /** The words it is to carry: those asked for it, or those of its passages. */
  words: number;
  /** The words it holds: `words` when it holds them all. */
  holds: number;
  /** The size its running text is set at, in points, on a free slide. */
  size?: number;
  /** The size running text has on a free slide of this template when nothing makes it smaller. */
  templateSize?: number;
  /** The slide is set dense (`FreeArrange.dense`). */
  dense?: true;
  /** For a slide set dense: the most words of a text of up to 12 paragraphs it holds at the least size that reads. `holds` counts the slide's own words in its own paragraphs, and can be more or less. */
  most?: number;
  /** For a slide of a document: its passages, as "p3-p5". */
  passages?: string;
}

/** A provider for work that asks no model. Planning a slide's room never calls it; should anything, it is told. */
const NO_MODEL: LlmProvider = {
  generateJson: async () => {
    throw new PlanError('No model is asked to work out what the slides of a plan hold');
  },
};

/**
 * Works out, by code alone, what each slide of a plan that has asked words or passages of a
 * document will carry: the size its text is set at and whether it is cut, as the deck will be
 * written. What a chat tool says before writing, and what a caller can show beside a plan. A slide
 * that cannot be built is left out, as is every slide without words or passages: `[]` for a plan
 * with neither, or one that cannot be planned at all (`writeDeck` says why).
 */
export function planTextFit(template: Template, plan: DeckPlan, options: Pick<PlanOptions, 'patterns' | 'free' | 'icons' | 'assets' | 'layout' | 'numberPatterns'> = {}): SlideTextFit[] {
  const slides: unknown[] = Array.isArray(plan?.slides) ? plan.slides : [];
  const isEntry = (slide: unknown): slide is OutlineEntry => typeof slide === 'object' && slide !== null && !Array.isArray(slide);
  const given = plan?.brief?.document;
  if (!(typeof given === 'string' && given.trim()) && !slides.some((slide) => isEntry(slide) && readWords(slide.words) !== undefined)) return [];
  let planned: Planning;
  try {
    planned = planning(template, plan.brief, { ...options, llm: NO_MODEL }, { layouts: plan.layouts, hasFree: slides.some((slide) => isEntry(slide) && isFreeEntry(slide)), documentLanguage: readLanguageTag(plan.documentLanguage) });
  } catch (error) {
    if (error instanceof PlanError) return [];
    throw error;
  }
  const { language, direction, context, byId, document: doc } = planned;
  const deck = { direction, textScale: context.textScale };
  const kit = freeKit(template);
  const templateSize = kit ? freeBodySize(kit, context.textScale) : undefined;
  return slides.flatMap((entry, index): SlideTextFit[] => {
    if (!isEntry(entry)) return [];
    // The passages of a document the slide carries, or else the words asked for it.
    const carry = doc ? slideDocument(doc, entry.passages, entry.words, language, showsTable(entry, byId)) : undefined;
    const words = carry ? carry.words : readWords(entry.words);
    if (words === undefined) return [];
    const need = carry ? carry.need : charsOfWords(words, language);
    const roomIn = (pattern: UsablePattern, names?: ReadonlySet<string>): number => (carry ? documentRoom(pattern, context, carry.shape, names) : heldBy(pattern, context, { names, language }));
    const fit = (room: number): SlideTextFit => ({ slide: index + 1, words, holds: room >= need ? words : Math.floor((words * room) / need) });
    const named = (found: SlideTextFit): SlideTextFit[] => [carry ? { ...found, passages: carry.label } : found];
    if (isFreeEntry(entry) && !byId.has(entry.pattern)) {
      // As writeDeck builds it: its texts set out for the length asked, or for the passages it carries.
      const blocks = entryBlocks(entry);
      if (blocks.length > MOST_PARTS) return [];
      const lengths = freeLengths(blocks, need);
      const names = new Set(Object.keys(lengths));
      const written = freeWritePattern(template, deck, { blocks, arrange: entry.arrange, lengths, ...paragraphsFor(carry, names) }, context);
      if ('reason' in written) return [];
      const found = fit(roomIn(written.pattern, names));
      const size = runningSize(written.pattern, names);
      if (size !== undefined) found.size = size;
      if (size !== undefined && templateSize !== undefined) found.templateSize = templateSize;
      if (entry.arrange?.dense === true) {
        found.dense = true;
        const room = freeTextRoom(template, deck);
        if (room) found.most = wordsOfChars(room.densest, language);
      }
      return named(found);
    }
    const usable = byId.get(entry.pattern);
    if (!usable) return [];
    const range = itemRange(usable.def);
    if (entry.items !== undefined && range && !(Number.isInteger(entry.items) && entry.items >= range.min && entry.items <= range.max)) return [];
    const sized = entry.items !== undefined ? planPattern(usable.def, context, undefined, entry.items) : undefined;
    if (sized && !('pattern' in sized)) return [];
    return named(fit(roomIn(sized ? sized.pattern : usable)));
  });
}

/**
 * Whether a setting turns free slides off for a plan (see `freeTurnedOff`): the option, the
 * template's manifest, or a list of patterns that does not name them. No slide's text is then set
 * smaller, which what is said of a slide that is cut says too. `false` for a plan that cannot be
 * planned at all.
 */
export function freeSlidesOff(template: Template, plan: DeckPlan, options: Pick<PlanOptions, 'patterns' | 'free' | 'icons' | 'assets' | 'layout' | 'numberPatterns'> = {}): boolean {
  const slides: unknown[] = Array.isArray(plan?.slides) ? plan.slides : [];
  try {
    const hasFree = slides.some((slide) => typeof slide === 'object' && slide !== null && isFreeEntry(slide as OutlineEntry));
    return planning(template, plan.brief, { ...options, llm: NO_MODEL }, { layouts: plan.layouts, hasFree }).freeOff === true;
  } catch (error) {
    if (error instanceof PlanError) return false;
    throw error;
  }
}

/**
 * The paragraphs a free slide's text from a document is set out for, where it goes to one slot,
 * and the lengths of those paragraphs where code may set them as they are (`TextShape.asIs`):
 * neither asked to be shorter nor a table that goes into a table (see `filledText`).
 */
export function paragraphsFor(carry: SlideDocument | undefined, names: ReadonlySet<string>): { paragraphs?: Record<string, number>; asIs?: Record<string, readonly number[]> } {
  const [name] = [...names];
  if (!carry || name === undefined || names.size !== 1) return {};
  const asIs = carry.shape.asIs && carry.asked === undefined && !carry.table && carry.shape.asIs.chars === carry.need ? { asIs: { [name]: carry.shape.asIs.each } } : {};
  return { paragraphs: { [name]: carry.shape.paragraphs }, ...asIs };
}

/**
 * Where code sets a slide's running text to its passages, as given: the slot and its paragraphs.
 * Only for a slide whose text is in the deck's script, not asked to be shorter, with no table that
 * goes into a table, whose pattern has one running text and holds the passages, in as many
 * paragraphs as that takes or one. A layout's slot of paragraphs takes them as its writer is told
 * (`documentParagraphs`): one fewer than its limit, of four or more, for a text of three or more.
 * `undefined` for any other: the writer writes it, and joins paragraphs where there are more.
 */
export function filledText(doc: DocumentPlan, carry: SlideDocument, pattern: UsablePattern, context: CatalogContext, names?: ReadonlySet<string>): { slot: string; value: string | string[] } | undefined {
  if (!doc.same || carry.asked !== undefined || carry.table) return undefined;
  if (documentRoom(pattern, context, carry.shape, names) < carry.need) return undefined;
  const running = runningNames(pattern, names);
  const [slot] = running;
  if (slot === undefined || running.length !== 1 || (names && !/^b\d+Text$/.test(slot))) return undefined;
  const planned = pattern.slots.find((one) => one.slot.name === slot);
  if (!planned || planned.slot.kind !== 'text') return undefined;
  // A free slide's text draws no bullets, so an item keeps its own. A layout's list draws its own
  // bullets, as text from the writer has it: an item's bullet there would show twice.
  const paragraphs = fillParagraphs(doc.passages, carry.own, { bullets: names !== undefined });
  if (paragraphs.length === 0) return undefined;
  if (planned.mode === 'text') return paragraphs.length === 1 ? { slot, value: paragraphs[0] as string } : undefined;
  const scale = budgetScale(context);
  const most = names ? textBudget(planned.slot, scale).maxParagraphs : documentParagraphs(planned.slot, scale, carry.shape.paragraphs);
  return most === undefined || paragraphs.length <= most ? { slot, value: paragraphs } : undefined;
}

/**
 * The title code sets on a slide whose running text code sets from its passages (`filledText`):
 * the heading its first passage stands under, word for word, where the slide's title holds it,
 * and on a free slide in one line (`kit`), the line its plan keeps for a title. Its text then does
 * not say the heading again (`fillParagraphs` leaves out the slide's first heading), and the
 * writer is not asked for a title. `undefined` for a slide whose text opens under no heading, or
 * whose title would not hold it: the writer titles it.
 */
export function headingTitle(doc: DocumentPlan, carry: SlideDocument, pattern: UsablePattern, context: CatalogContext, kit?: FreeKit): string | undefined {
  const heading = openingSection(doc.passages, carry.own)?.trim();
  if (!doc.same || !heading) return undefined;
  const planned = pattern.slots.find((one) => one.slot.name === 'title');
  if (!planned || planned.slot.kind !== 'text') return undefined;
  const chars = [...heading].length;
  const most = textBudget(planned.slot, budgetScale(context)).maxChars;
  if ((most !== undefined && chars > most) || (kit && freeTitleLines(kit, chars) > 1)) return undefined;
  return heading;
}

/**
 * What an outline may choose from. A text-only deck is offered the plain text patterns and nothing
 * else, so the wish cannot be missed. A deck limited to a list of patterns is offered that list:
 * it names the patterns outright, so it comes before the wish for text only.
 */
function outlineCatalog(planned: Planning, brief: Brief): Catalog {
  if (!brief.textOnly || planned.lists.limit) return planned.catalog;
  return { ...planned.catalog, patterns: planned.catalog.patterns.filter((pattern) => isPlainText(pattern.def) || pattern.own) };
}

/**
 * A plan read back from a store may be anything. One that is no object with a list of slides, or
 * a slide of it that is no object, is told in a {@link PlanError}: nothing of it can be read.
 */
function checkPlanShape(plan: unknown): void {
  const isObject = (value: unknown) => typeof value === 'object' && value !== null && !Array.isArray(value);
  if (!isObject(plan) || !Array.isArray((plan as { slides?: unknown }).slides)) throw new PlanError('The plan is no object with a list of slides');
  const loose = ((plan as { slides: unknown[] }).slides).flatMap((slide, index) => (isObject(slide) ? [] : [index + 1]));
  if (loose.length > 0) throw new PlanError(`The plan has slides that are no object with a pattern and a purpose: ${loose.slice(0, 10).join(', ')}${loose.length > 10 ? ` and ${loose.length - 10} more` : ''}`);
}

/**
 * Looks at a plan for slides that a layout the template lacks would show better, makes those
 * layouts in the template's own look, and gives them to their slides. What `composeLayouts` in
 * {@link PlanOptions} does inside {@link planOutline}, as a step of its own: for a plan that was
 * already shown, or to learn what was made and what could not be. `max` defaults to 2. Nothing is
 * made for a plan whose brief limits the deck to a list of patterns, and nothing while free slides
 * are on offer (pass `free: false` to have layouts made in their place).
 */
export async function composeLayouts(template: Template, plan: DeckPlan, options: PlanOptions & { max?: number }): Promise<ComposeLayoutsResult> {
  checkPlanShape(plan);
  const planned = planning(template, plan.brief, options, { layouts: plan.layouts, hasFree: plan.slides.some(isFreeEntry) });
  if (planned.lists.limit || planned.free) return { plan, made: [], failed: [] };
  options.onProgress?.({ stage: 'layouts' });
  return composeForPlan(template, plan, outlineCatalog(planned, plan.brief).patterns, { llm: options.llm, max: options.max ?? 2, ...(options.signal ? { signal: options.signal } : {}) });
}

/** What a change to a plan may set beside what its instruction says. Code sets these; the model is not asked. */
export interface PlanSettings {
  /**
   * The colours the deck will have, as `Brief.colors`. `null` returns to the template's. Left
   * out, the plan keeps the colours it has.
   */
  colors?: DeckColors | null;
  /**
   * The only patterns the deck may use from now on, as `Brief.patterns`. `null` or an empty list
   * lifts the limit. Left out, the plan keeps the list it has. This is the person's wish, kept
   * with the plan; `patterns` is what the app allows, and is not kept. Slides that are planned
   * already keep their patterns unless the instruction says otherwise. No id is an error here:
   * one the template does not have is dropped, and a list that names none of its patterns sets
   * nothing.
   */
  deckPatterns?: string[] | null;
}

/** A plan with the colours and the list of patterns a change sets by code. The plan given is not changed. */
export function withSettings(template: Template, plan: DeckPlan, settings: PlanSettings): DeckPlan {
  const brief: Brief = { ...plan.brief };
  if (settings.colors !== undefined) {
    const colors = storedColors(settings.colors);
    if (colors) brief.colors = colors;
    else delete brief.colors;
  }
  // The list may name `free` where the template can carry free slides.
  const list = givenPatterns(template.manifest, settings.deckPatterns, namesFree(settings.deckPatterns) && freeKnown(template));
  if (list?.length) brief.patterns = list;
  else if (list) delete brief.patterns;
  return { ...plan, brief };
}

/**
 * Changes a plan by instruction ("add a slide about pricing", "make it shorter", "in English").
 * One model call. The instruction is kept with the brief, so facts it adds reach the writing stage.
 * Colours and the list of allowed patterns are not read from the instruction: give them in the
 * options (`colors`, `deckPatterns`) and they are written to the plan's brief.
 */
export async function editPlan(
  template: Template,
  plan: DeckPlan,
  instruction: string,
  options: PlanOptions & PlanSettings,
): Promise<{ plan: DeckPlan; summary: string }> {
  checkPlanShape(plan);
  const request = instruction.trim();
  const brief: Brief = { ...withSettings(template, plan, options).brief, instructions: [plan.brief.instructions?.trim(), request].filter(Boolean).join('\n') };
  const freeSlides = plan.slides.filter(isFreeEntry);
  // A document the plan was made from is read in the language the planner said it is written in.
  const storedLanguage = readLanguageTag(plan.documentLanguage);
  // The list is read leniently, the one given with this change too: only a new brief is held to its ids.
  const planned = planning(template, brief, options, { layouts: plan.layouts, hasFree: freeSlides.length > 0, documentLanguage: storedLanguage });
  const { language, density, catalog, byId, context, free, document: doc } = planned;
  // Of a plan made from a document: how many passages it has, and those its slides carry. The
  // others are on no slide, and the change is told so.
  const passages = doc?.passages.length;
  const placedBefore = new Set(doc ? plan.slides.flatMap((slide) => readPassageNumbers(slide.passages, doc.passages.length)) : []);
  // Of the list the deck is limited to, what is on offer here. The rest would only mislead.
  const onOffer = (planned.lists.limit ?? []).filter((id) => byId.has(id) || (id === FREE_PATTERN && free));
  const limit = onOffer.length > 0 ? onOffer : undefined;
  // A free slide the plan has stays a free slide whatever is on offer for a new one, so the
  // answer may name the kinds the plan uses beside the kinds on offer.
  const keepsFree = freeSlides.length > 0 && !byId.has(FREE_PATTERN);
  const shows = free || keepsFree ? { kinds: kindsUsed(freeSlides.map((slide) => ({ blocks: entryBlocks(slide) })), free?.kinds) } : undefined;
  const reading: FreeAnswer | undefined = shows ? { ...freeReading(shows.kinds, context, plan.slides), offered: free !== undefined } : undefined;
  const addedPhotos = free && density !== 'none' ? withFreePhotos(ADDED_DENSITY_GUIDANCE[density]) : ADDED_DENSITY_GUIDANCE[density];

  options.onProgress?.({ stage: 'outline' });
  // Where the plan has words or a document, its lines say what code made of them: a slide cut to
  // what it holds, text set small, and the most a slide set dense holds. A plan with neither is
  // shown as before.
  const fits = doc || plan.slides.some((slide) => readWords(slide.words) !== undefined) ? new Map(planTextFit(template, plan, options).map((fit) => [fit.slide, fit])) : undefined;
  /** What a slide of a plan made from a document carries of it: its passages and their length. */
  const carried = (slide: OutlineEntry): string[] => {
    const carry = doc ? slideDocument(doc, slide.passages, slide.words, language, showsTable(slide, byId)) : undefined;
    return carry ? [`${carry.label}, ${carry.need} characters`] : [];
  };
  const marks = (number: number): string[] => {
    const fit = fits?.get(number);
    if (!fit) return [];
    return [
      ...(fit.holds < fit.words ? ['cut to fit'] : []),
      ...(fit.size !== undefined && fit.templateSize !== undefined && setSmall(fit.size, fit.templateSize) ? [`text at ${fit.size} pt`] : []),
      ...(fit.most !== undefined ? [`at most about ${tensDown(fit.most)} words`] : []),
    ];
  };
  const prompt = [
    briefText(plan.brief),
    // The document comes after the brief, with the passages no slide carries marked.
    ...(doc ? [documentBlock(doc, { change: true, onNoSlide: new Set(doc.passages.map((passage) => passage.number).filter((number) => !placedBefore.has(number))) })] : []),
    `The plan as it is now${plan.title ? ` ("${plan.title}")` : ''}, in the language with BCP 47 tag "${language}":\n${plan.slides
      .map((slide, index) => {
        const words = readWords(slide.words);
        const shape = isFreeEntry(slide) ? describeEntry({ ...slide, blocks: entryBlocks(slide) }) : slide.items !== undefined ? `${slide.items} items` : '';
        const said = [shape, words !== undefined ? `${words} words` : '', ...carried(slide), ...marks(index + 1)].filter(Boolean).join('; ');
        return `${index + 1}. ${slide.pattern}${said ? ` (${said})` : ''}: ${slide.purpose}`;
      })
      .join('\n')}`,
    `Requested change:\n${request}`,
    `Return the whole plan after the change. Choose patterns from this list ("?" marks an optional slot):`,
    [
      describeCatalog(catalog, budgetScale(context)),
      ...(shows ? [freeLine(shows.kinds, free ? { comfortable: free.comfortable, alone: catalog.patterns.length === 0, plain: plan.brief.textOnly === true && !planned.lists.limit, asked: 'request' } : { closed: true })] : []),
    ]
      .filter(Boolean)
      .join('\n'),
    [
      `- ${wishesRule('request')}`,
      '- Change only what the request asks for. Keep every other slide, its pattern, its purpose and its place.',
      ...(shows ? [`- ${FREE_CHANGE_RULE}`] : []),
      '- Choose a pattern only when the content can fill every required slot within its limits.',
      `- ${columnsRule('request')}`,
      `- ${wordsRule('request', true)}`,
      ...(fits ? [`- ${WORDS_CHANGE_RULE}`] : []),
      ...(doc ? DOCUMENT_CHANGE_RULES.map((rule) => `- ${rule}`) : []),
      `- ${SLIDE_KIND_RULE}`,
      // The whole list stays on offer: the limit was a wish, and the requested change may be another.
      ...(limit
        ? [`- ${limitRule(limit)}`]
        : plan.brief.textOnly
          ? ['- This deck was asked to be text only: keep to the patterns marked "Plain text." or "Made for this deck." unless the requested change asks for something else.']
          : []),
      '- Colours are set by the software, not by you: a wish about colour changes no slide.',
      `- ${colorNote(options.colors !== undefined)}`,
      // Photos are a matter for the slides a change adds, and for none in a deck of text or of named patterns.
      ...(limit || plan.brief.textOnly ? [] : [`- ${addedPhotos}`]),
    ].join('\n'),
  ].join('\n\n');
  // Of a plan made from a document, each slide names its passages, and the answer those the
  // request leaves out, the slides it asks for and the language of the document.
  const more: Record<string, JsonSchema> =
    passages !== undefined ? { left_out: leftOutProperty(passages, 'request'), asked_slides: askedSlidesProperty('request'), document_language: documentLanguageProperty } : {};
  const schema: JsonSchema = {
    type: 'object',
    properties: {
      summary: { type: 'string', description: 'One sentence saying what you changed, written for the person who asked, in their language' },
      language: { type: 'string', description: 'BCP 47 tag of the deck language after the change' },
      title: { type: 'string', description: 'Title of the deck after the change' },
      slides: { type: 'array', minItems: 1, items: outlineEntrySchema(catalog, shows, passages !== undefined ? { passages } : undefined) },
      ...more,
    },
    required: ['summary', 'language', 'title', 'slides', ...Object.keys(more)],
  };

  const ask = asker(options.llm, systemPrompt(language, 'brief', doc !== undefined), options.signal, 'low');
  // An edit may add slides, so the room is for a plan somewhat longer than the one it starts from.
  // Where slides name their parts, that is a little more for each slide, and for each run of parts
  // the plan has; where they name their passages, a little more again.
  const runs = plan.slides.reduce((sum, slide) => sum + (isFreeEntry(slide) ? blockRuns(entryBlocks(slide)).length : 0), 0);
  const room =
    planRoom(plan.slides.length + 10) +
    (shows ? FREE_ENTRY_ROOM * (plan.slides.length + 10 + runs) : 0) +
    (passages !== undefined ? 8 * (plan.slides.length + 10) + 2 * passages + DOCUMENT_ROOM : 0);
  type Answer = { summary?: unknown; language?: unknown; left_out?: unknown; asked_slides?: unknown; document_language?: unknown } | null;
  let answer = (await ask(prompt, schema, room)) as Answer;
  let outline = readOutline(answer, byId, reading, false, passages);
  // An answer that puts no passage on a slide, where the plan had some, and leaves most of them
  // out of "left_out" too, is asked for again, with any problems of patterns.
  const gap =
    passages !== undefined && placedBefore.size > 0 && !outline.entries.some((entry) => (entry.passages?.length ?? 0) > 0)
      ? grossGap(outline.entries, readPassageNumbers(answer?.left_out, passages), passages)
      : undefined;
  // So is one that names three slides or fewer and leaves one of them without passages.
  const bare = passages !== undefined ? bareSlides(outline.entries, readAskedSlides(answer?.asked_slides), placedBefore.size, 'request') : undefined;
  if (outline.problem || gap || bare) {
    const problem = [outline.problem, gap, bare].filter(Boolean).join('; ');
    const placing = doc ? ', with every passage on one slide or in "left_out"' : '';
    const retry = outline.problem
      ? `${prompt}\n\nYour previous answer could not be used: ${problem}. Answer again using only the listed patterns${outline.hint ?? ''}${placing}.`
      : `${prompt}\n\nYour previous answer could not be used: ${problem}. Answer again${placing}.`;
    const first = { answer, outline };
    answer = (await ask(retry, schema, room)) as Answer;
    outline = readOutline(answer, byId, reading, true, passages);
    if (outline.problem) {
      // How a document is placed is put right in code: an answer that was asked again for that alone is used as it was.
      if (first.outline.problem) throw new PlanError(`The model did not produce a usable plan: ${outline.problem}`);
      ({ answer, outline } = first);
    }
  }

  const nextLanguage = cleanText(answer?.language) || language;
  if (nextLanguage !== language) {
    brief.language = nextLanguage;
    delete brief.direction; // follows the new language
  }
  delete brief.slideCount; // the plan itself now says how many slides there are
  const lp = lengthPlanning(template, planned, catalog);
  // Of a plan made from a document, the passages are put right in code: a passage the answer
  // dropped goes back next to its neighbour, and one that was on no slide stays so unless it is
  // placed now. The language of the document is the answer's, or else the one the plan had.
  let entries = outline.entries;
  const documentLanguage = doc ? (readLanguageTag(answer?.document_language, doc.text) ?? storedLanguage) : undefined;
  const read = doc ? inLanguage(doc, nextLanguage, documentLanguage) : undefined;
  if (read && passages !== undefined) {
    const given = entries;
    entries = settlePassages(given, readPassageNumbers(answer?.left_out, passages), read, {
      carries: carriesText(byId, context, read),
      // A change keeps the slides its answer has: code adds none.
      fixed: true,
      slides: 0,
      room: 0,
      textHome: textEntryOf(lp, catalog, context, given[0]?.pattern ?? FREE_PATTERN),
      sizeOf: (numbers) => needOf(read, numbers, language),
      previous: placedBefore,
      // Where the request names a number of slides, three or fewer, every slide carries the document.
      ...(readAskedSlides(answer?.asked_slides) !== undefined && given.length <= 3 ? { keepEnd: endKeeps(read, lp) } : {}),
    }).entries;
  }
  // A slide asked an amount of text gets the room for it, on that slide: code adds no slide. A
  // slide the plan has already, with the same layout and words (and passages), was fitted when it
  // was planned and stays as it is, set dense where it was: an answer that names its parts again
  // sets out their arrangement anew, which would lose that.
  const key = (slide: OutlineEntry) =>
    JSON.stringify([slide.pattern, slide.purpose, readWords(slide.words) ?? null, slide.items ?? null, slide.blocks ?? null, ...(doc ? [slide.passages ?? null] : [])]);
  const had = new Map(plan.slides.map((slide) => [key(slide), slide]));
  for (const entry of entries) {
    const before = had.get(key(entry));
    if (before?.arrange?.dense === true && isFreeEntry(entry)) entry.arrange = { ...entry.arrange, dense: true };
  }
  const settled = (slide: OutlineEntry) => had.has(key(slide));
  // Each slide whose passages changed gets the room for them on that slide, set smaller where it
  // needs it. They are shared again only where the request names a number of slides: then over
  // every slide of that number, one kept as it was too.
  // A slide's text comes before its other parts only where the plan had that slide with those
  // parts, and they were not asked for in a change (`partsAsked`): a part the request asks for
  // stays, now and in later changes, and the plan says what of the text it leaves out.
  const first = textFirstIn(plan.slides, entries, lp);
  const textFirst = (entry: OutlineEntry): boolean => first.get(entry) ?? false;
  const placed = read ? fitDocument(entries, read, lp, { most: entries.length, fixed: true, settled, balance: readAskedSlides(answer?.asked_slides) !== undefined, textFirst }) : entries;
  // So for words asked for a slide: a chart or cards the request adds beside its text stay.
  const slides = fitOutline(placed, lp, read ? (slide) => settled(slide) || (slide.passages?.length ?? 0) > 0 : settled, textFirst);
  const next: DeckPlan = { version: 1, title: outline.title || plan.title, brief, slides };
  if (documentLanguage) next.documentLanguage = documentLanguage;
  const layouts = usedLayouts(plan.layouts, slides);
  if (layouts) next.layouts = layouts;
  return { plan: next, summary: cleanText(answer?.summary) };
}

/**
 * The paragraphs the running slots of a pattern take in all, as their writer is told them, or
 * `undefined` where one of them is not a slot of paragraphs with a limit.
 */
function toldParagraphsOf(pattern: UsablePattern, context: CatalogContext, names?: ReadonlySet<string>): number | undefined {
  const scale = budgetScale(context);
  let told = 0;
  for (const name of runningNames(pattern, names)) {
    const planned = pattern.slots.find((one) => one.slot.name === name);
    if (!planned || planned.mode !== 'paragraphs') return undefined;
    const most = textBudget(planned.slot as TextSlotDef, scale).maxParagraphs;
    if (most === undefined) return undefined;
    told += most;
  }
  return told > 0 ? told : undefined;
}

/** The parts of a slide, as a change to a plan matches it with a slide the plan has: its layout, or of a free slide its kinds of part in any order. */
const partsKey = (slide: OutlineEntry): string => JSON.stringify([slide.pattern, isFreeEntry(slide) ? [...entryBlocks(slide)].sort() : null]);

/**
 * Whether a slide has parts that a text may push off (see `homeFor`): of a free slide any part but
 * a text or a photo, beside a text or a list; of a layout one that is not plain text, words asked
 * for it or not, so that a chart a change asks for stays when words are asked for it later.
 */
function hasOtherParts(slide: OutlineEntry, planning: LengthPlanning): boolean {
  const text = readWords(slide.words) !== undefined || (slide.passages?.length ?? 0) > 0;
  if (!isFreeEntry(slide)) return !planning.plainText(slide.pattern);
  const blocks = entryBlocks(slide);
  return (text || blocks.some((kind) => kind === 'text' || kind === 'list')) && blocks.some((kind) => kind !== 'text' && kind !== 'photo');
}

/**
 * Of the slides of a change to a plan, those whose text comes before their other parts (see
 * `fitDocument` and `fitOutline`): the slides the plan had with those parts. Each slide the plan
 * has is paired with one slide of the change of the same parts (`partsKey`), one to one: first one
 * of the same purpose, then one in the same place, then in order. A slide of the change beyond the
 * slides the plan had of its parts is new, and keeps the parts it was asked with; so does one
 * whose parts were asked for in an earlier change (`OutlineEntry.partsAsked`). Those are marked
 * so on the slide of the change, for the changes after it, in a plan of any kind: a part added to
 * a plan with no document and no words stays when words are asked for its slide later. Without a
 * mark a stored slide needs nothing: its parts were planned.
 */
function textFirstIn(before: readonly OutlineEntry[], entries: readonly OutlineEntry[], planning: LengthPlanning): Map<OutlineEntry, boolean> {
  const stored = new Map<string, number[]>();
  before.forEach((slide, at) => stored.set(partsKey(slide), [...(stored.get(partsKey(slide)) ?? []), at]));
  const pairs = new Map<number, number>();
  const taken = new Set<number>();
  const pair = (match: (from: number, to: number) => boolean): void => {
    entries.forEach((entry, to) => {
      if (pairs.has(to)) return;
      const from = (stored.get(partsKey(entry)) ?? []).find((at) => !taken.has(at) && match(at, to));
      if (from === undefined) return;
      pairs.set(to, from);
      taken.add(from);
    });
  };
  pair((from, to) => (before[from] as OutlineEntry).purpose === (entries[to] as OutlineEntry).purpose);
  pair((from, to) => from === to);
  pair(() => true);
  const first = new Map<OutlineEntry, boolean>();
  entries.forEach((entry, to) => {
    const from = pairs.get(to);
    const asked = from === undefined ? hasOtherParts(entry, planning) : (before[from] as OutlineEntry).partsAsked === true;
    if (asked) entry.partsAsked = true;
    else delete entry.partsAsked;
    first.set(entry, !asked && from !== undefined);
  });
  return first;
}

/**
 * Second stage: writes the content of every slot of a plan's slides, against a schema built for
 * exactly that sequence. Drafts that break the template's limits go back to the model.
 */
export async function writeDeck(loaded: Template, plan: DeckPlan, options: PlanOptions): Promise<PlanResult> {
  checkPlanShape(plan);
  const { brief } = plan;
  const layouts = usedLayouts(plan.layouts, plan.slides);
  // The deck's own layouts are built onto the template: written, checked and repaired as any other.
  // The brief's list of patterns narrows nothing here: the plan has chosen, and a slide a later
  // wish put outside the list is written like any other. Ids the template has lost are dropped.
  // A free slide of the plan is written whatever is on offer for a new one.
  // A document is set into slides as it is only where the language the planner read it in is the deck's.
  const documentLanguage = readLanguageTag(plan.documentLanguage);
  const { layout, language, direction, fields, context, byId, lists, template, document: doc, freeOff } = planning(loaded, brief, options, { layouts, hasFree: plan.slides.some(isFreeEntry), documentLanguage });
  const colors = storedColors(brief.colors);
  // A slide that lists its parts is a free slide, unless the template has a pattern of its own under that id.
  const isFree = (entry: OutlineEntry) => isFreeEntry(entry) && !byId.has(entry.pattern);
  const missing = plan.slides.filter((slide) => !isFree(slide) && !byId.has(slide.pattern)).map((slide) => slide.pattern);
  if (missing.length > 0) {
    throw new PlanError(`The plan uses patterns that cannot be filled with what was supplied: ${[...new Set(missing)].join(', ')}`);
  }
  // Slides that hold less text than was asked for them, said with the deck.
  const short: Issue[] = [];
  // Slots told the whole of the length asked for their slide, held to it once they are written.
  const told: ToldLength[] = [];
  // Slides written from their passages of a document, held to its words; and each slide's passages.
  const kept: KeptCheck[] = [];
  const carries = new Map<number, number[]>();
  const items: WriteJob[] = plan.slides.map((entry, index): WriteItem | WriteJob => {
    const number = index + 1;
    // The passages of the deck's document the slide carries: the writer is told their text, and
    // where the slide holds them, code sets them into it.
    const carry = doc ? slideDocument(doc, entry.passages, entry.words, language, showsTable(entry, byId)) : undefined;
    if (carry) carries.set(number, carry.own);
    const passage = carry && doc ? { passage: { label: carry.label, text: passageText(doc.passages, carry.own) } } : {};
    // An amount of text asked for the slide: the writer is told it, and the deck when the slide holds less.
    const words = carry ? undefined : readWords(entry.words);
    // A slide set dense says the size its text is set at when even that holds less.
    const dense = isFree(entry) && entry.arrange?.dense === true;
    const asked = (pattern: UsablePattern, names?: ReadonlySet<string>): UsablePattern => {
      if (carry && doc) {
        const room = documentRoom(pattern, context, carry.shape, names);
        const size = dense ? runningSize(pattern, names) : undefined;
        if (room < carry.need) short.push(size !== undefined ? textCutDense(number, carry.label, carry.need, room, size) : textCut(number, carry.label, carry.need, room, freeOff));
        const check = keptCheck(doc, number, carry, room, runningNames(pattern, names)[0]);
        if (check) kept.push(check);
        const shaped = withPassages(pattern, carry.need, context, { names, language, label: carry.label, shape: carry.shape, words: carry.asked });
        // A slide that holds the text only with its paragraphs joined says so, as one that cuts it does.
        const joined = room >= carry.need && carry.asked === undefined ? toldParagraphsOf(shaped, context, names) : undefined;
        if (joined !== undefined && joined < carry.shape.paragraphs) short.push(linesJoined(number, carry.label, carry.shape.paragraphs, joined));
        return shaped;
      }
      if (words === undefined) return pattern;
      const room = heldBy(pattern, context, { names, language });
      const size = dense ? runningSize(pattern, names) : undefined;
      if (room < charsOfWords(words, language)) short.push(size !== undefined ? shortOfWordsDense(number, words, room, size, language) : shortOfWords(number, words, room, language, freeOff));
      const whole = toldAsAsked(pattern, words, context, { names, language });
      if (whole) told.push({ slide: number, words, ...whole });
      return withLength(pattern, words, context, { names, language });
    };
    if (isFree(entry)) {
      // Its slots are worked out for these parts, on the template without the deck's own layouts.
      const blocks = entryBlocks(entry);
      // A plan is something an app keeps and may show for editing, so its counts are not taken at
      // their word. A slide is held to the parts a model's answer is held to: every part beyond
      // costs its share of a writing call, and a hundred thousand of them never end.
      if (blocks.length > MOST_PARTS) throw new PlanError(`Slide ${number}: a free slide takes at most ${MOST_PARTS} parts, not ${blocks.length}`);
      // Its texts are set out for the length asked, where one was, or for the passages it carries.
      const need = carry ? carry.need : words !== undefined ? charsOfWords(words, language) : undefined;
      const lengths = need !== undefined ? freeLengths(blocks, need) : undefined;
      const names = new Set(Object.keys(lengths ?? {}));
      const set = { blocks, arrange: entry.arrange, lengths, ...paragraphsFor(carry, names) };
      let written = freeWritePattern(loaded, { direction, textScale: context.textScale }, set, context);
      if ('reason' in written) throw new PlanError(`Slide ${number}: the free slide ${written.reason}`);
      // A text that holds its passages is set by code, and is not the writer's to write.
      let fill = carry && doc ? filledText(doc, carry, written.pattern, context, names) : undefined;
      // Its title then is the heading its text opens under, where the title holds it.
      const titled = fill && carry && doc ? headingTitle(doc, carry, written.pattern, context, freeKit(loaded, written.base)) : undefined;
      if (fill) {
        const rest = freeWritePattern(loaded, { direction, textScale: context.textScale }, set, context, new Set(titled ? [fill.slot, 'title'] : [fill.slot]));
        if ('pattern' in rest) written = rest;
        else fill = undefined;
      }
      // The frame it was planned on is kept with the slide, so that it stands on the same one later.
      const pattern = fill ? written.pattern : asked(written.pattern, names);
      const job: WriteJob = { number, pattern, purpose: entry.purpose, free: { blocks, ...(entry.arrange ? { arrange: entry.arrange } : {}), base: written.base }, ...passage };
      if (fill) {
        job.carried = titled ? { [fill.slot]: fill.value, title: titled } : { [fill.slot]: fill.value };
        job.filled = fill.slot;
      }
      // What the slide can say should all its parts come back empty and its purpose be empty too:
      // the title of the deck, or else the first words of the brief. A free slide stops no deck.
      const about = cleanText(brief.prompt);
      const said = cleanText(plan.title) || (about.length <= 120 ? about : about.slice(0, 120).replace(/\s+\S*$/, ''));
      if (said) job.fallback = said;
      return job;
    }
    const usable = byId.get(entry.pattern) as UsablePattern;
    // The number of items is checked before anything is built for it or written: the slots of a
    // million items are never worked out, and a number the row does not take costs no writing
    // call. The sentence is written here because `readItems` passes over a number that is not
    // whole (it reads an answer, where the number drawn then stands in).
    const range = itemRange(usable.def);
    if (entry.items !== undefined && range && !(Number.isInteger(entry.items) && entry.items >= range.min && entry.items <= range.max)) {
      const takes = range.min === range.max ? `exactly ${range.min}` : `${range.min} to ${range.max}`;
      throw new PlanError(`Slide ${number}: pattern "${entry.pattern}" takes ${takes} ${usable.def.repeat?.item ?? 'item'}s, not ${String(entry.items)}`);
    }
    // A row of items gets the slots of exactly the number the plan settled on.
    const sized = entry.items !== undefined ? planPattern(usable.def, context, undefined, entry.items) : undefined;
    const pattern = sized && 'pattern' in sized ? sized.pattern : usable;
    // A text that holds its passages is set by code, and is not the writer's to write.
    const fill = carry && doc ? filledText(doc, carry, pattern, context) : undefined;
    if (fill) {
      // Its title then is the heading its text opens under, where the title holds it.
      const titled = carry && doc ? headingTitle(doc, carry, pattern, context) : undefined;
      const set = new Set(titled ? [fill.slot, 'title'] : [fill.slot]);
      const carried: Record<string, SlotValue> = titled ? { [fill.slot]: fill.value, title: titled } : { [fill.slot]: fill.value };
      return { number, pattern: { ...pattern, slots: pattern.slots.filter((planned) => !set.has(planned.slot.name)) }, purpose: entry.purpose, ...passage, carried, filled: fill.slot };
    }
    return { number, pattern: asked(pattern), purpose: entry.purpose, ...passage };
  });
  const system = systemPrompt(language, 'brief', doc !== undefined);
  // A text written well short of the length asked for it goes back once, with the first answer's
  // problems, and so does a slide that keeps too few of the words of its text from the document.
  // The issues that send them are the writer's alone: the deck is told after the last answer.
  const sent = new WeakSet<Issue>();
  const notes = (deck: DeckSpec, issues: Issue[], round: number): Issue[] =>
    round > 0
      ? []
      : [...writtenUnder(deck, told).map(writtenLength), ...(doc ? keptIssues(doc, kept, deck, issues) : [])].map((issue) => {
          sent.add(issue);
          return issue;
        });

  options.onProgress?.({ stage: 'write', slides: items.length });
  const draft = await writeSlides({
    ask: asker(options.llm, system, options.signal),
    askLight: asker(options.llm, system, options.signal, 'low'),
    request: writePrompt(brief, items, context),
    requestFor: (job, task) => partsPrompt(brief, items, job, task, context),
    template: loaded,
    jobs: items,
    context,
    assemble: (written) => {
      const deck: DeckSpec = {
        version: 1,
        language,
        direction,
        brief: [brief.prompt.trim(), brief.instructions?.trim()].filter(Boolean).join('\n'),
        slides: items.map((item) => {
          const slide = written.get(item.number) ?? { pattern: item.pattern.def.id, slots: {} };
          const own = carries.get(item.number);
          // Its document's headings in its text are marked, a heading in brackets is set plain, and a text does not open with its title.
          return own ? (doc ? settleDocumentSlide({ ...slide, passages: own }, doc.passages, own, { translated: !doc.same }) : { ...slide, passages: own }) : slide;
        }),
      };
      if (context.textScale !== 1) deck.textScale = context.textScale;
      if (plan.title) deck.title = plan.title;
      if (layouts) deck.layouts = layouts;
      if (Object.keys(fields).length > 0) deck.fields = fields;
      // What the person asked of the whole deck goes with it, so that a later change can keep to it.
      if (colors) deck.colors = colors;
      if (lists.asked) deck.patterns = lists.asked;
      // The document goes with the deck, so that a later change can draw on it, with the language it is written in.
      if (doc && documentLanguage) deck.documentLanguage = documentLanguage;
      if (doc) deck.document = brief.document as string;
      return deck;
    },
    manifest: template.manifest,
    measure: measurerFor(template),
    review: reviewWith(options),
    ...(told.length > 0 || kept.length > 0 ? { notes } : {}),
    layout,
    maxRepairs: options.maxRepairs ?? 2,
    onRepair: (attempt, issues) => options.onProgress?.({ stage: 'repair', attempt, issues }),
  });

  const errors = draft.issues.filter((issue) => issue.severity === 'error');
  if (errors.length > 0) throw new DeckValidationError(errors);
  const issues = draft.issues.filter((issue) => !sent.has(issue));
  const shortWritten = writtenUnder(draft.deck, told).map((under) => writtenShort(under, language));
  const keptLeft = doc ? keptWarnings(doc, kept, draft.deck) : [];
  return { deck: draft.deck, warnings: [...issues, ...unstatedChartValues(template.manifest, draft.deck), ...short, ...keptLeft, ...shortWritten], outline: plan.slides };
}

/** Brief to deck spec in one go: {@link planOutline}, then {@link writeDeck}. */
export async function planDeck(template: Template, brief: Brief, options: PlanOptions): Promise<PlanResult> {
  return writeDeck(template, await planOutline(template, brief, options), options);
}
