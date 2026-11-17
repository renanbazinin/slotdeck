import { parseHex } from '../color/space.ts';
import { unstatedChartValues } from '../deck/numbers.ts';
import { DeckValidationError } from '../deck/validate.ts';
import type { Issue, LayoutMode } from '../deck/validate.ts';
import { FREE_PATTERN } from '../deck/spec.ts';
import type { DeckColors, DeckSpec, FreeArrange, FreeBlockKind } from '../deck/spec.ts';
import { describeBlocks } from '../free/blocks.ts';
import { blockRuns } from '../free/offer.ts';
import type { FreeOffer } from '../free/offer.ts';
import type { JsonRequest, JsonSchema, LlmProvider } from '../llm/provider.ts';
import type { Template } from '../render/render.ts';
import { TEXT_SIZES, itemRange } from '../template/manifest.ts';
import { measurerFor } from '../text/measure.ts';
import type { Direction, PatternDef, TemplateManifest, TextSize } from '../template/manifest.ts';
import { budgetScale, buildCatalog, describeCatalog, isPlainText, outlineEntrySchema, outlineSchema, planPattern, readItems } from './catalog.ts';
import type { AssetDef, Catalog, CatalogContext, IconDef, ImageDensity, UsablePattern, WriteItem } from './catalog.ts';
import { withLayouts } from '../template/compose.ts';
import type { PatternDesign } from '../template/compose.ts';
import { FREE_CHANGE_RULE, FREE_OUTLINE_RULE, MOST_PARTS, describeEntry, entryBlocks, freeKnown, freeLine, freeOnOffer, freeReading, freeWritePattern, isFreeEntry, kindsUsed, namesFree, readShows, withFreePhotos } from './free.ts';
import type { FreeReading } from './free.ts';
import { charsOfWords, freeLengths, heldBy, homeFor, partPurpose, partsFor, readWords, shortOfWords, spreads, withLength, wordsOfChars, wordsRule } from './length.ts';
import type { LengthPlanning, TextHome } from './length.ts';
import { composeForPlan, usedLayouts } from './layouts.ts';
import type { ComposeLayoutsResult } from './layouts.ts';
import type { DeckReview } from './review.ts';
import { cleanText, describeParts, suppliesText, systemPrompt, writeSlides, writtenApart } from './write.ts';
import type { WriteJob } from './write.ts';

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
   * the length. A slide whose text was spread over several slides has its share here.
   */
  words?: number;
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

const RTL_LANGUAGES = new Set(['ar', 'dv', 'fa', 'he', 'ps', 'sd', 'ug', 'ur', 'yi']);

export function directionOf(language: string): Direction {
  return RTL_LANGUAGES.has(language.toLowerCase().split('-')[0] ?? '') ? 'rtl' : 'ltr';
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
function outlinePrompt(brief: Brief, catalog: Catalog, count: { min: number; max: number }, density: ImageDensity, textScale = 1, limited = false, free?: FreeOffer): string {
  const size = count.min === count.max ? `exactly ${count.min}` : `${count.min} to ${count.max}`;
  // The person's own list comes before the wish for text only: it names the patterns outright.
  const textOnly = brief.textOnly === true && !limited;
  const photos = free && density !== 'none' ? withFreePhotos(DENSITY_GUIDANCE[density]) : DENSITY_GUIDANCE[density];
  return [
    briefText(brief),
    `Plan a deck of ${size} ${count.max === 1 ? 'slide' : 'slides'}. Choose one pattern for each slide from this list ("?" marks an optional slot):`,
    [describeCatalog(catalog, textScale), ...(free ? [freeLine(free.kinds, { comfortable: free.comfortable, alone: catalog.patterns.length === 0, plain: textOnly })] : [])].filter(Boolean).join('\n'),
    [
      `- ${wishesRule('brief')}`,
      ...(free ? [`- ${FREE_OUTLINE_RULE}`] : []),
      // A deck of one or two slides has no room for an opening and a closing slide: its slides carry the content.
      ...(count.max >= 3
        ? ['- A deck of three or more slides opens with an opening pattern and ends with a closing pattern when the list has them, unless the brief asks for a deck without them.']
        : []),
      '- For a pattern with a row of items, set "items" to the number the brief asks for, or else the number the content really has. Do not pad a list to fill a row.',
      `- ${columnsRule('brief')}`,
      `- ${wordsRule('brief')}`,
      '- Choose a pattern only when the content can fill every required slot within its limits.',
      ...(textOnly
        ? ['- The deck is text only, and every pattern in the list is plain text. Using the same pattern for several slides in a row is expected.']
        : [`- ${SLIDE_KIND_RULE}`, ...(limited ? [] : ['- Otherwise vary the patterns across the deck.', `- ${photos}`])]),
    ].join('\n'),
  ].join('\n\n');
}

/** A slide in the list a writer is given: `s3: cards. Three goals`. A free slide says which of its parts are what. */
const slideLine = (job: WriteJob): string => `s${job.number}: ${job.pattern.def.id}${job.free ? ` (${describeParts(job.free.blocks)})` : ''}. ${job.purpose}`;

function writePrompt(brief: Brief, jobs: WriteJob[], context: CatalogContext): string {
  // A slide of many parts is written by calls of its own. The writer of the others is told that it is there.
  const here = jobs.filter((job) => !writtenApart(job));
  const apart = jobs.filter(writtenApart);
  return [
    briefText(brief),
    `Write the content of each slide of this deck:\n${here.map(slideLine).join('\n')}`,
    ...(apart.length > 0 ? [`These slides of the deck are written apart, not here:\n${apart.map((job) => `s${job.number}: ${FREE_PATTERN} (${describeBlocks(job.free?.blocks ?? [])}). ${job.purpose}`).join('\n')}`] : []),
    ...suppliesText(here, context),
    'Fill every required property. Leave an optional property out when the slide is better without it.',
  ].join('\n\n');
}

/** The prompt of a call that writes some parts of one slide by themselves: the brief, the deck around the slide, then the task. */
function partsPrompt(brief: Brief, jobs: WriteJob[], job: WriteJob, task: string, context: CatalogContext): string {
  const deck = jobs.map((other) => `s${other.number}: ${other.pattern.def.id}${other.free ? ` (${describeBlocks(other.free.blocks)})` : ''}. ${other.purpose}`).join('\n');
  return [briefText(brief), `The slides of the deck:\n${deck}`, task, ...suppliesText([job], context)].join('\n\n');
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
  // A share of pictures that is not one of the list ("many", or "toString", which every object has) is the usual one.
  const density: ImageDensity = typeof brief.imageDensity === 'string' && Object.hasOwn(DENSITY_GUIDANCE, brief.imageDensity) ? brief.imageDensity : 'medium';
  const fields = brief.fields ?? {};

  const numbers = options.numberPatterns ?? 'from-brief';
  const figures = numbers === 'always' || (numbers === 'from-brief' && hasFigures(withoutColorCodes(`${brief.prompt} ${brief.instructions ?? ''}`, brief.colors)));
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
 * the parts of a free slide is then put right here and is no problem any more.
 */
function readOutline(answer: unknown, byId: Map<string, UsablePattern>, free?: FreeAnswer, final = false): { title: string; entries: OutlineEntry[]; problem?: string; hint?: string } {
  const raw = answer as { title?: unknown; slides?: unknown } | null;
  const list = Array.isArray(raw?.slides) ? raw.slides : [];
  const problems: string[] = [];
  let counts = false;
  const entries = list.map((item, index) => {
    const raw = item as { pattern?: unknown; purpose?: unknown; items?: unknown; words?: unknown };
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
  const count =
    typeof brief.slideCount === 'number'
      ? { min: brief.slideCount, max: brief.slideCount }
      : (brief.slideCount ?? { min: 6, max: 10 });
  // Choosing a pattern for each slide needs little reasoning, and is much quicker with little.
  const ask = asker(options.llm, systemPrompt(language, 'brief'), options.signal, 'low');
  const reading: FreeAnswer | undefined = free ? { ...freeReading(free.kinds, context), offered: true } : undefined;
  const schema = outlineSchema(catalog, count, free);
  // A free slide names its parts, which is a little more to say for each slide.
  const room = planRoom(count.max) + (free ? FREE_ENTRY_ROOM * count.max : 0);

  options.onProgress?.({ stage: 'outline' });
  const request = outlinePrompt(brief, catalog, count, density, budgetScale(context), limited, free);
  let outline = readOutline(await ask(request, schema, room), byId, reading);
  if (outline.problem) {
    const retry = `${request}\n\nYour previous answer could not be used: ${outline.problem}. Answer again using only the listed patterns${outline.hint ?? ''}.`;
    outline = readOutline(await ask(retry, schema, room), byId, reading, true);
    if (outline.problem) throw new PlanError(`The model did not produce a usable outline: ${outline.problem}`);
  }
  // A slide that was asked an amount of text gets the room for it, or the slides it takes.
  const slides = fitOutline(outline.entries, lengthPlanning(template, planned, catalog), count.max);
  const plan: DeckPlan = { version: 1, title: outline.title, brief, slides };
  // A layout made for the deck would be one the person did not name. And where free slides are on
  // offer, a free slide is what the deck gets for content that no layout shows.
  if (!options.composeLayouts || limited || free) return plan;
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
    free: kinds.has('text')
      ? (blocks, arrange, lengths) => {
          const written = freeWritePattern(loaded, { direction, textScale: context.textScale }, { blocks, arrange, lengths }, context);
          return 'pattern' in written ? written.pattern : undefined;
        }
      : undefined,
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
 * keeps its layout when that holds the text, or moves to one that does (see `homeFor`). Where
 * none does and the slide is plain text, its text goes over as many slides as it needs while the
 * deck may have up to `most`, each with its share of the words and its purpose numbered
 * ("(1/2)"). A slide that still holds less is told when the deck is written. `settled` names the
 * slides fitted already, which a change to a plan passes as they are: a slide left short because
 * the deck had no room for more stays so until its words change.
 */
export function fitOutline(entries: OutlineEntry[], planning: LengthPlanning, most: number, settled?: (entry: OutlineEntry) => boolean): OutlineEntry[] {
  if (!entries.some((entry) => entry.words !== undefined)) return entries;
  let spare = Math.max(0, most - entries.length);
  return entries.flatMap((entry): OutlineEntry[] => {
    if (entry.words === undefined || settled?.(entry)) return [entry];
    const current: TextHome = isFreeEntry(entry) ? { blocks: entry.blocks, arrange: entry.arrange } : { pattern: entry.pattern, items: entry.items };
    const found = homeFor(current, entry.words, planning);
    const placed = found.home === current ? entry : placeEntry(entry, found.home);
    const parts = found.holds || !spreads(found.home, planning) ? 1 : partsFor(entry.words, wordsOfChars(found.room, planning.language), spare);
    spare -= parts - 1;
    if (parts === 1) return [placed];
    const words = Math.ceil(entry.words / parts);
    return Array.from({ length: parts }, (_, at) => ({ ...placed, words, purpose: partPurpose(entry.purpose, at + 1, parts) }));
  });
}

/** An outline entry given another home for its text: a layout with its items, or a free slide of these parts. */
function placeEntry(entry: OutlineEntry, home: TextHome): OutlineEntry {
  const { pattern: _pattern, items: _items, blocks: _blocks, arrange: _arrange, ...rest } = entry;
  if ('blocks' in home) return { pattern: FREE_PATTERN, ...rest, blocks: home.blocks, ...(home.arrange ? { arrange: home.arrange } : {}) };
  return { pattern: home.pattern, ...rest, ...(home.items !== undefined ? { items: home.items } : {}) };
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
  // The list is read leniently, the one given with this change too: only a new brief is held to its ids.
  const planned = planning(template, brief, options, { layouts: plan.layouts, hasFree: freeSlides.length > 0 });
  const { language, density, catalog, byId, context, free } = planned;
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
  const prompt = [
    briefText(plan.brief),
    `The plan as it is now${plan.title ? ` ("${plan.title}")` : ''}, in the language with BCP 47 tag "${language}":\n${plan.slides
      .map((slide, index) => {
        const words = readWords(slide.words);
        const shape = isFreeEntry(slide) ? describeEntry({ ...slide, blocks: entryBlocks(slide) }) : slide.items !== undefined ? `${slide.items} items` : '';
        const said = [shape, words !== undefined ? `${words} words` : ''].filter(Boolean).join('; ');
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
  const schema: JsonSchema = {
    type: 'object',
    properties: {
      summary: { type: 'string', description: 'One sentence saying what you changed, written for the person who asked, in their language' },
      language: { type: 'string', description: 'BCP 47 tag of the deck language after the change' },
      title: { type: 'string', description: 'Title of the deck after the change' },
      slides: { type: 'array', minItems: 1, items: outlineEntrySchema(catalog, shows) },
    },
    required: ['summary', 'language', 'title', 'slides'],
  };

  const ask = asker(options.llm, systemPrompt(language, 'brief'), options.signal, 'low');
  // An edit may add slides, so the room is for a plan somewhat longer than the one it starts from.
  // Where slides name their parts, that is a little more for each slide, and for each run of parts the plan has.
  const runs = plan.slides.reduce((sum, slide) => sum + (isFreeEntry(slide) ? blockRuns(entryBlocks(slide)).length : 0), 0);
  const room = planRoom(plan.slides.length + 10) + (shows ? FREE_ENTRY_ROOM * (plan.slides.length + 10 + runs) : 0);
  let answer = (await ask(prompt, schema, room)) as { summary?: unknown; language?: unknown } | null;
  let outline = readOutline(answer, byId, reading);
  if (outline.problem) {
    answer = (await ask(`${prompt}\n\nYour previous answer could not be used: ${outline.problem}. Answer again using only the listed patterns${outline.hint ?? ''}.`, schema, room)) as typeof answer;
    outline = readOutline(answer, byId, reading, true);
    if (outline.problem) throw new PlanError(`The model did not produce a usable plan: ${outline.problem}`);
  }

  const nextLanguage = cleanText(answer?.language) || language;
  if (nextLanguage !== language) {
    brief.language = nextLanguage;
    delete brief.direction; // follows the new language
  }
  delete brief.slideCount; // the plan itself now says how many slides there are
  // A slide asked an amount of text gets the room for it, and a text that needs more slides may
  // take up to ten more. A slide the plan has already, with the same layout and words, was fitted
  // when it was planned and stays as it is: "a better title" does not spread a text that was left
  // on one slide because the deck had room for no more.
  const key = (slide: OutlineEntry) => JSON.stringify([slide.pattern, slide.purpose, readWords(slide.words) ?? null, slide.items ?? null, slide.blocks ?? null]);
  const had = new Set(plan.slides.map(key));
  const slides = fitOutline(outline.entries, lengthPlanning(template, planned, catalog), outline.entries.length + 10, (slide) => had.has(key(slide)));
  const next: DeckPlan = { version: 1, title: outline.title || plan.title, brief, slides };
  const layouts = usedLayouts(plan.layouts, slides);
  if (layouts) next.layouts = layouts;
  return { plan: next, summary: cleanText(answer?.summary) };
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
  const { layout, language, direction, fields, context, byId, lists, template } = planning(loaded, brief, options, { layouts, hasFree: plan.slides.some(isFreeEntry) });
  const colors = storedColors(brief.colors);
  // A slide that lists its parts is a free slide, unless the template has a pattern of its own under that id.
  const isFree = (entry: OutlineEntry) => isFreeEntry(entry) && !byId.has(entry.pattern);
  const missing = plan.slides.filter((slide) => !isFree(slide) && !byId.has(slide.pattern)).map((slide) => slide.pattern);
  if (missing.length > 0) {
    throw new PlanError(`The plan uses patterns that cannot be filled with what was supplied: ${[...new Set(missing)].join(', ')}`);
  }
  // Slides that hold less text than was asked for them, said with the deck.
  const short: Issue[] = [];
  const items: WriteJob[] = plan.slides.map((entry, index): WriteItem | WriteJob => {
    const number = index + 1;
    // An amount of text asked for the slide: the writer is told it, and the deck when the slide holds less.
    const words = readWords(entry.words);
    const asked = (pattern: UsablePattern, names?: ReadonlySet<string>): UsablePattern => {
      if (words === undefined) return pattern;
      const room = heldBy(pattern, context, { names, language });
      if (room < charsOfWords(words, language)) short.push(shortOfWords(number, words, room, language));
      return withLength(pattern, words, context, { names, language });
    };
    if (isFree(entry)) {
      // Its slots are worked out for these parts, on the template without the deck's own layouts.
      const blocks = entryBlocks(entry);
      // A plan is something an app keeps and may show for editing, so its counts are not taken at
      // their word. A slide is held to the parts a model's answer is held to: every part beyond
      // costs its share of a writing call, and a hundred thousand of them never end.
      if (blocks.length > MOST_PARTS) throw new PlanError(`Slide ${number}: a free slide takes at most ${MOST_PARTS} parts, not ${blocks.length}`);
      // Its texts are set out for the length asked, where one was.
      const lengths = words !== undefined ? freeLengths(blocks, charsOfWords(words, language)) : undefined;
      const written = freeWritePattern(loaded, { direction, textScale: context.textScale }, { blocks, arrange: entry.arrange, lengths }, context);
      if ('reason' in written) throw new PlanError(`Slide ${number}: the free slide ${written.reason}`);
      // The frame it was planned on is kept with the slide, so that it stands on the same one later.
      const pattern = asked(written.pattern, new Set(Object.keys(lengths ?? {})));
      const job: WriteJob = { number, pattern, purpose: entry.purpose, free: { blocks, ...(entry.arrange ? { arrange: entry.arrange } : {}), base: written.base } };
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
    return { number, pattern: asked(sized && 'pattern' in sized ? sized.pattern : usable), purpose: entry.purpose };
  });
  const system = systemPrompt(language, 'brief');

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
        slides: items.map((item) => written.get(item.number) ?? { pattern: item.pattern.def.id, slots: {} }),
      };
      if (context.textScale !== 1) deck.textScale = context.textScale;
      if (plan.title) deck.title = plan.title;
      if (layouts) deck.layouts = layouts;
      if (Object.keys(fields).length > 0) deck.fields = fields;
      // What the person asked of the whole deck goes with it, so that a later change can keep to it.
      if (colors) deck.colors = colors;
      if (lists.asked) deck.patterns = lists.asked;
      return deck;
    },
    manifest: template.manifest,
    measure: measurerFor(template),
    review: reviewWith(options),
    layout,
    maxRepairs: options.maxRepairs ?? 2,
    onRepair: (attempt, issues) => options.onProgress?.({ stage: 'repair', attempt, issues }),
  });

  const errors = draft.issues.filter((issue) => issue.severity === 'error');
  if (errors.length > 0) throw new DeckValidationError(errors);
  return { deck: draft.deck, warnings: [...draft.issues, ...unstatedChartValues(template.manifest, draft.deck), ...short], outline: plan.slides };
}

/** Brief to deck spec in one go: {@link planOutline}, then {@link writeDeck}. */
export async function planDeck(template: Template, brief: Brief, options: PlanOptions): Promise<PlanResult> {
  return writeDeck(template, await planOutline(template, brief, options), options);
}
