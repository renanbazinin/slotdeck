import { DeckValidationError, validateDeck } from '../deck/validate.ts';
import type { Issue } from '../deck/validate.ts';
import { keptGaps, statedNumbers, unstatedChartValues, unstatedDetails } from '../deck/numbers.ts';
import { FREE_PATTERN, isChartValue, isFreeSlide, isImageRef, toParagraphs } from '../deck/spec.ts';
import type { DeckColors, DeckSpec, FreeArrange, FreeBlockKind, SlideSpec, SlotValue } from '../deck/spec.ts';
import { isFreeBlockKind } from '../free/blocks.ts';
import { freeKit } from '../free/kit.ts';
import { freeTextRoom } from '../free/offer.ts';
import { freeBodySize } from '../free/plan.ts';
import { freeFor } from '../free/resolve.ts';
import type { JsonSchema } from '../llm/provider.ts';
import { said } from '../messages/en.ts';
import type { Template } from '../render/render.ts';
import { withLayouts } from '../template/compose.ts';
import { toldCounter } from '../template/told.ts';
import { TEXT_SIZES, effectiveSlots, itemCount, itemRange } from '../template/manifest.ts';
import type { Direction, PatternDef, TextSize } from '../template/manifest.ts';
import { measurerFor } from '../text/measure.ts';
import { budgetScale, buildCatalog, describeCatalog, isPlainText, planPattern, readItems } from './catalog.ts';
import type { CatalogContext, ImageDensity, UsablePattern } from './catalog.ts';
import {
  DOCUMENT_EDIT_RULE,
  documentBlock,
  documentWriteRule,
  inLanguage,
  keptCheck,
  keptIssues,
  keptShare,
  keptWarnings,
  passageLabel,
  passageText,
  passagesChangeProperty,
  documentLimit,
  readDocument,
  readLanguageTag,
  openingSection,
  readPassageNumbers,
  settleDocumentSlide,
  slideDocument,
  slideText,
  textCut,
  textCutDense,
} from './document.ts';
import type { KeptCheck, SlideDocument } from './document.ts';
import {
  DENSE_EDIT_RULE,
  MOST_WORDS,
  WORDS_NOTE,
  charsOfWords,
  documentRoom,
  freeLengths,
  holdsOnlyAsIs,
  heldBy,
  homeForChars,
  isDenseHome,
  photoLeftOut,
  readWords,
  runningNames,
  runningSize,
  sameHome,
  setSmall,
  shortOfWords,
  shortOfWordsDense,
  tensDown,
  toldAsAsked,
  undense,
  withLength,
  withPassages,
  wordsOfChars,
  wordsRule,
  writtenLength,
  writtenShort,
  writtenUnder,
} from './length.ts';
import type { LengthPlanning, TextHome, ToldLength } from './length.ts';
import {
  FREE_ENTRY_ROOM,
  PlanError,
  SLIDE_KIND_RULE,
  allowedPatterns,
  asker,
  colorNote,
  columnsRule,
  directionOf,
  asIsOf,
  filledText,
  freePartsOf,
  headingTitle,
  givenPatterns,
  hasFigures,
  limitRule,
  paragraphsFor,
  planRoom,
  reviewWith,
  storedColors,
  wishesRule,
  withoutColorCodes,
} from './plan.ts';
import type { PlanOptions, SlideTextFit } from './plan.ts';
import { MOST_PHOTOS, freeKnown, freeLine, freeOnOffer, freeTurnedOff, freeWritePattern, kindsUsed, namesFree } from './free.ts';
import { usedLayouts } from './layouts.ts';
import { partOf } from './parts.ts';
import {
  MOVED_RULE,
  PARTS_RULE,
  TURN_RULE,
  arrangeAfter,
  carryOver,
  carryParts,
  cut,
  freePhotos,
  heldParts,
  listFree,
  listParts,
  partsProperties,
  partsRoom,
  readParts,
  roomAsHeld,
  sameArrange,
  sameParts,
  settleParts,
  shownValue,
  statesArrange,
} from './reshape.ts';
import type { FreeChange, PartPlan } from './reshape.ts';
import { cleanText, describeParts, suppliesText, systemPrompt, writeSlides, writtenApart } from './write.ts';
import type { WriteJob } from './write.ts';

export interface EditOptions extends PlanOptions {
  /** Replaces the deck-level values, e.g. a translated header text. Defaults to the deck's own. */
  fields?: Record<string, string>;
  /** Reading direction after the edit. Defaults to the direction of the deck's language. */
  direction?: Direction;
  /** Whether slides added by the edit may use photo patterns. Defaults to `medium`. */
  imageDensity?: ImageDensity;
  /**
   * Sets the deck's colours, as {@link recolorDeck} does: `{ accent: '1E6B3A' }`. `null` removes
   * them, and the deck is back in the template's colours. A slide that has colours of its own
   * keeps them. Code sets the colours once the edit is planned; no model reads a colour out of
   * the instruction. Left out, colours stay as they are.
   */
  colors?: DeckColors | null;
  /**
   * Gives `colors` to these slides alone, in place of the whole deck. The numbers are those of
   * the deck as it is before the edit, from 1: a slide the edit moves takes its colours along,
   * and a slide the edit adds gets none. A slide the edit removes cannot be coloured: a warning
   * says so. An empty list names no slide, so the colours are the deck's, as if it were left out.
   */
  colorSlides?: number[];
  /**
   * Sets the only patterns the deck may use from now on (`DeckSpec.patterns`). `null` or an empty
   * list lifts the limit. Left out, the deck keeps the list it has. This is the person's wish,
   * kept in the deck; `patterns` is what the app allows, and is not kept. Slides that are there
   * already keep their patterns unless the instruction says otherwise. No id is an error here:
   * one the template does not have is dropped, and a list that names none of its patterns sets
   * nothing.
   */
  deckPatterns?: string[] | null;
  /**
   * Stop before writing when a slide the change writes would hold less text than was asked for
   * it: planEdit then throws a {@link TextCutError} that names those slides, and nothing is
   * written. Off: such a slide is cut to what it holds, and the result's warnings say so.
   */
  stopBeforeCut?: boolean;
}

/** Thrown by planEdit with `stopBeforeCut` when a slide would be cut: nothing was written. */
export class TextCutError extends PlanError {
  /** The slides that would hold less than they carry, numbered as in the deck after the change. */
  readonly fits: SlideTextFit[];
  /** Whether free slides are off, so that no slide's text is set smaller: the message then says so. */
  readonly freeOff: boolean;
  constructor(message: string, fits: SlideTextFit[], freeOff = false) {
    super(message);
    this.name = 'TextCutError';
    this.fits = fits;
    this.freeOff = freeOff;
  }

  /**
   * The cut as a person reads it: a note for each slide (`change-would-cut`), for `formatIssue`. A
   * slide says its size where its text is set small, and that free slides are off where it is not.
   */
  get issues(): Issue[] {
    return this.fits.map((fit) => ({
      severity: 'warning',
      ...said('change-would-cut', { slide: fit.slide, words: fit.words, holdsWords: fit.holds, size: fit.size, freeOff: this.freeOff && fit.size === undefined }),
    }));
  }
}

/** Sets or removes the colours of the slides with these numbers (from 1), and of no other. */
function recolorSlides(deck: DeckSpec, stored: DeckColors | undefined, numbers: number[]): DeckSpec {
  const chosen = new Set(numbers);
  const slides = deck.slides.map((slide, index) => {
    if (!chosen.has(index + 1) || (!stored && slide.colors === undefined)) return slide;
    const { colors: _before, ...rest } = slide;
    return stored ? { ...rest, colors: structuredClone(stored) } : rest;
  });
  return { ...deck, slides };
}

/**
 * Sets or removes colours on a deck, and returns the new deck. No model is asked, and the deck
 * given is not changed. Without `slides` they are the deck's colours (`DeckSpec.colors`); a slide
 * that has colours of its own keeps them. With `slides`, those slides get them as their own
 * (numbers from 1; a number the deck does not have is passed over) and the deck's stay as they
 * are. An empty list names no slide, and is read as if it were left out. `null` removes: the
 * slides named follow the deck again, or the deck its template.
 */
export function recolorDeck(deck: DeckSpec, colors: DeckColors | null, slides?: number[]): DeckSpec {
  const stored = storedColors(colors);
  if (slides?.length) return recolorSlides(deck, stored, slides);
  const { colors: _before, ...rest } = deck;
  return stored ? { ...rest, colors: stored } : rest;
}

export interface EditResult {
  deck: DeckSpec;
  /** One sentence from the model on what it changed, for the person who asked. */
  summary: string;
  /**
   * Slide numbers in the new deck that were written or rewritten, and those of free slides whose
   * parts or arrangement code changed without writing anything (a part taken off, three side by
   * side). Every other slide is copied as it was. A free slide of which only the title was
   * offered, and came back as it was, is not among them: `warnings` says that nothing on it changed.
   */
  rewritten: number[];
  /**
   * With `colorSlides`: the numbers those slides have in the new deck, where `colors` were set
   * or removed. A slide the edit removed is not among them.
   */
  recolored?: number[];
  warnings: Issue[];
}

interface PlanEntry {
  /** Number of the existing slide this one is based on; 0 for a new slide. */
  from: number;
  pattern: string;
  /** What to write or change in the slide's content. Empty when the content stays as it is. */
  change: string;
  /** Whether the pictures and icons of the slide should change. */
  images: boolean;
  /** For a pattern with a row of items: how many items the slide has after the change. */
  items?: number;
  /** For a free slide: its parts after the change, and how they are set out. */
  free?: FreeChange;
  /** The words of running text the request asks for on the slide, where it asks for an amount. */
  words?: number;
  /** The slide is a free text alone set dense to hold its words (`FreeArrange.dense`). Set by code. */
  dense?: true;
  /**
   * In a deck made from a document: the passages the slide carries after the change. As the plan
   * gives them, `undefined` keeps those of the slide it comes from; once settled by code, the list
   * itself, left out where it is empty.
   */
  passages?: number[];
  /**
   * In a deck made from a document: the request asks for the slide's text in other words. Its text
   * is then written by the writer, also where its passages change, and not set by code.
   */
  reword?: true;
}

/** A slide to write, with what the writer is shown of the slide it comes from. */
type EditJob = WriteJob & {
  source?: SlideSpec;
  /** The slide takes another layout, or nothing of its parts is kept: its content is shown as material. */
  relaid: boolean;
  /** For a free slide: lines on what stays of it, and on what its new parts take the place of. */
  notes?: string[];
  /** For a free slide that takes the place of another slide: what that slide held, less what was carried over. */
  material?: Record<string, unknown>;
  /**
   * For a free slide the plan has words for while it names every part as it is, none to be
   * written again: the slide as it was. Its title is all the writer is offered. Where that comes
   * back as it was, nothing on the slide changed, and the deck is told (see {@link nothingChanged}).
   */
  asItWas?: SlideSpec;
};

const SHORTEN = 'Shorten the text to fit the new text size. Keep the meaning.';
/** What a slide whose passages of the document changed is written for, where the plan says nothing. */
const FROM_DOCUMENT = 'Write the slide from its text from the document.';
/** The same for a slide that gave up all its passages. */
const GAVE_UP = 'Its text from the document is on another slide now: write the slide without it.';
/**
 * What `reword` says of a slide of a deck made from a document. A slide whose passages change has
 * its text set by code as the document has it: a wish for other words on that slide needs saying.
 */
const REWORD_NOTE = 'true only when the request asks for the text of this slide in other words: simpler, shorter sentences, another tone. Leave it out otherwise: a slide whose passages change has its text set as the document has it.';

/**
 * Said of a free slide the plan of a change had words for and named no part of to write again,
 * when its title came back as it was: the request was not met on that slide. It is a note on the
 * deck, with no slide number in the issue: writing the slide again would give the same.
 */
const nothingChanged = (slide: number): Issue => ({ severity: 'warning', ...said('nothing-changed', { slide }) });

/** A title as one line, for telling whether a change left it as it was. */
const titleLine = (value: SlotValue | undefined): string => (typeof value === 'string' ? cleanText(value) : Array.isArray(value) ? cleanText(value.join(' ')) : '');

function describeValue(value: SlotValue | undefined): string {
  if (value === undefined || value === null) return '';
  if (typeof value === 'string') return JSON.stringify(value);
  if (Array.isArray(value)) return JSON.stringify(value);
  if (!isImageRef(value)) return `${isChartValue(value) ? 'chart' : 'table'} ${JSON.stringify(value)}`;
  return value.prompt ? `picture of ${JSON.stringify(value.prompt)}` : `picture "${value.asset}"`;
}

/**
 * The length of a free slide's running texts and lists, as they are written, in words: their
 * characters at the characters a word takes in the language, as an asked length is counted. A
 * length kept so takes no more characters than the text has; counted between spaces, a text of
 * short words would ask for more than it has, and more than a slide that holds it.
 */
function wordsIn(slide: SlideSpec, language: string): number {
  const chars = Object.entries(slide.slots)
    .filter(([name, value]) => /^b\d+(?:Text|Items)$/.test(name) && (typeof value === 'string' || Array.isArray(value)))
    .reduce((sum, [, value]) => sum + toParagraphs(value as string | string[]).join('').length, 0);
  return wordsOfChars(chars, language);
}

/**
 * A value of a slide that carries passages of a document: a long text is cut short, its paragraphs
 * joined by " / ". The document is shown whole before the deck.
 */
function describeCarried(value: SlotValue | undefined): string {
  if (typeof value === 'string' || (Array.isArray(value) && value.every((item) => typeof item === 'string'))) {
    const text = toParagraphs(value as string | string[])
      .filter(Boolean)
      .join(' / ');
    const shown = cut(text);
    if (shown !== text) return JSON.stringify(shown);
  }
  return describeValue(value);
}

/**
 * The deck as the model sees it: one line per slide with every slot's content. A free slide is
 * listed by its parts, numbered and without their data: it may hold two hundred charts, and the
 * plan of a change names its parts by these numbers. `marks` says more of some free slides, by
 * slide: the words of a text set dense, and the most it holds. `carried` names the passages of a
 * document a slide carries; such a slide's long texts are cut short.
 */
function describeDeck(deck: DeckSpec, isFree: (slide: SlideSpec) => boolean, marks?: ReadonlyMap<SlideSpec, readonly string[]>, carried?: ReadonlyMap<SlideSpec, string>): string {
  return deck.slides
    .map((slide, index) => {
      const own = carried?.get(slide);
      if (isFree(slide)) return `${index + 1}. ${listFree(slide, [...(own ? [own] : []), ...(marks?.get(slide) ?? [])])}`;
      const slots = Object.entries(slide.slots)
        .map(([name, value]) => `${name}: ${own ? describeCarried(value) : describeValue(value)}`)
        .filter((line) => !line.endsWith(': '))
        .join('; ');
      return `${index + 1}. [${slide.pattern}${slide.items !== undefined ? `, ${slide.items} items` : ''}${own ? `, ${own}` : ''}] ${slots}`;
    })
    .join('\n');
}

/**
 * The passages of the document each slide carries after a change, set on the plan's entries: a
 * list the plan gives comes first, each passage on the first slide that names it; a slide the plan
 * gives none keeps those of the slide it comes from, less any the plan names for another slide,
 * each passage once. Returns the slides whose passages are not those of the slide they come from:
 * they are written again.
 */
function resolveDeckPassages(entries: PlanEntry[], deck: DeckSpec, passages: number): Set<PlanEntry> {
  const had = (entry: PlanEntry): number[] => readPassageNumbers(entry.from > 0 ? deck.slides[entry.from - 1]?.passages : undefined, passages);
  const placed = new Set<number>();
  for (const entry of entries) {
    if (entry.passages === undefined) continue;
    entry.passages = entry.passages.filter((number) => !placed.has(number));
    for (const number of entry.passages) placed.add(number);
  }
  for (const entry of entries) {
    if (entry.passages !== undefined) continue;
    entry.passages = had(entry).filter((number) => !placed.has(number));
    for (const number of entry.passages) placed.add(number);
  }
  const changed = new Set<PlanEntry>();
  for (const entry of entries) {
    const before = had(entry);
    const now = entry.passages ?? [];
    if (now.length !== before.length || now.some((number, at) => number !== before[at])) changed.add(entry);
    if (now.length === 0) delete entry.passages;
  }
  return changed;
}

/** A short label for a slide that is not being rewritten, so the model knows what surrounds its work. */
function slideLabel(slide: SlideSpec): string {
  const title = slide.slots.title;
  const text = typeof title === 'string' ? title : Object.values(slide.slots).find((value): value is string => typeof value === 'string');
  return text ? JSON.stringify(text) : '';
}

/**
 * The schema of the plan of a change. `freeKinds` says the deck may hold free slides after it,
 * and of which kinds of part: an entry then says what such a slide is made of. Without it the
 * schema is what it was before there were free slides. `documentPassages` says the deck was made
 * from a document of so many passages: an entry may then say which it carries after the change.
 */
function editPlanSchema(patternIds: string[], slideCount: number, fieldNames: string[], mostItems: number, freeKinds?: readonly FreeBlockKind[], documentPassages?: number): JsonSchema {
  const schema = planSchemaBase(patternIds, slideCount);
  const top = schema.properties as Record<string, JsonSchema>;
  top.text_size = {
    type: 'string',
    enum: Object.keys(TEXT_SIZES),
    description: 'Size of body text after the change. Repeat the current size unless the request is about text size.',
  };
  (schema.required as string[]).push('text_size');
  const slide = (top.slides as JsonSchema).items as JsonSchema;
  const entry = slide.properties as Record<string, JsonSchema>;
  if (mostItems > 0) {
    entry.items = {
      type: 'integer',
      minimum: 1,
      maximum: mostItems,
      description: 'For a pattern with a row of items: how many items the slide has after the change. Leave it out for other patterns.',
    };
  }
  if (freeKinds) {
    // Asked of every slide, as in an outline, so that a model cannot choose `free` and leave its parts unsaid.
    const parts = partsProperties(freeKinds);
    Object.assign(entry, parts);
    (slide.required as string[]).push(...Object.keys(parts));
  }
  // Left out, a slide keeps the passages it has: most slides of a change keep theirs.
  if (documentPassages !== undefined) {
    entry.passages = passagesChangeProperty(documentPassages);
    entry.reword = { type: 'boolean', description: REWORD_NOTE };
  }
  if (fieldNames.length > 0) {
    const properties = schema.properties as Record<string, JsonSchema>;
    properties.fields = {
      type: 'object',
      description: 'The deck-level texts after the change. Repeat a value exactly unless the request affects it.',
      properties: Object.fromEntries(fieldNames.map((name) => [name, { type: 'string' }])),
      required: fieldNames,
    };
    (schema.required as string[]).push('fields');
  }
  return schema;
}

function planSchemaBase(patternIds: string[], slideCount: number): JsonSchema {
  return {
    type: 'object',
    properties: {
      summary: { type: 'string', description: 'One sentence saying what you changed, written for the person who asked, in their language' },
      language: { type: 'string', description: 'BCP 47 tag of the deck language after the change' },
      title: { type: 'string', description: 'Title of the deck after the change' },
      slides: {
        type: 'array',
        minItems: 1,
        items: {
          type: 'object',
          properties: {
            from: { type: 'integer', minimum: 0, maximum: slideCount, description: 'Number of the existing slide this slide is based on; 0 for a new slide' },
            pattern: { type: 'string', enum: patternIds },
            change: {
              type: 'string',
              description:
                'What to write or change in the content of this slide: its text, numbers, pictures or parts. Empty when the content stays as it is; a new colour or a new place in the deck is not a change of content.',
            },
            images: { type: 'boolean', description: 'true only when the pictures or icons of this slide should change' },
            words: { type: 'integer', minimum: 1, maximum: MOST_WORDS, description: WORDS_NOTE },
          },
          required: ['from', 'pattern', 'change', 'images'],
        },
      },
    },
    required: ['summary', 'language', 'title', 'slides'],
  };
}

/**
 * Applies an instruction such as "shorten slide 3" or "add a slide about pricing" to a deck.
 *
 * Two model calls: the first decides what the deck looks like after the change (which slides
 * stay, move, go, get rewritten or are new); the second writes only the slides that change.
 * Untouched slides are copied by code, so they cannot drift, and pictures are kept unless the
 * instruction is about them.
 *
 * A free slide is changed part by part. The plan names the parts the slide has after the change;
 * a part that stays is carried over by code under its new number, and only the parts that are new
 * or change go to the writer. Parts taken off or moved, and a new arrangement, need no writing;
 * where the plan has words for such a slide all the same, its title alone is offered to the
 * writer, since a new title is all they can still ask for. When the plan named every part as it
 * is and that title comes back as it was, a warning without a slide number says that nothing on
 * the slide was changed. An arrangement the plan states (a side, a number side by side) is always
 * set; `auto` and 0 hand a slide back to the software only where the plan changes the slide, and
 * state nothing on a slide it leaves alone. A part that was to be written again
 * and comes back with nothing keeps what it had, and a warning says that the change was not made.
 * An edit can make a free slide too, new or out of a slide of a pattern, where free slides are on
 * offer (`free`); one the deck has is changed and kept whatever is on offer.
 *
 * Colours and the list of patterns the deck is limited to are not read from the instruction:
 * give them in the options (`colors`, `colorSlides`, `deckPatterns`) and code sets them.
 */
export async function planEdit(template: Template, deck: DeckSpec, instruction: string, options: EditOptions): Promise<EditResult> {
  // A deck's document is read as it is, whatever the limit for a new one; a wrong limit is still the app's error, here as in every stage.
  documentLimit(options.maxDocumentChars);
  // A slide stored without its slots, or with slots that are no object, is read as a slide whose slots are empty.
  const slotless = (slide: SlideSpec) => typeof slide.slots !== 'object' || slide.slots === null;
  if (deck.slides.some(slotless)) deck = { ...deck, slides: deck.slides.map((slide) => (slotless(slide) ? { ...slide, slots: {} } : slide)) };
  // A list of patterns may name `free` where the template, as it was loaded, can carry free slides.
  const freeNamed = (namesFree(options.deckPatterns) || namesFree(deck.patterns)) && freeKnown(template);
  // Free slides are laid out and checked on the template as it was loaded.
  const loaded = template;
  // The deck's own layouts are built onto the template, so its slides can be rewritten like any other.
  template = withLayouts(template, deck.layouts);
  const { manifest } = template;
  const measure = measurerFor(template);
  const layout = options.layout ?? 'strict';
  const fields = options.fields ?? deck.fields ?? {};
  const currentLanguage = deck.language ?? manifest.language ?? 'en';
  const definitions = new Map<string, PatternDef>(manifest.patterns.map((pattern) => [pattern.id, pattern]));
  // The document the deck was made from, where it has one: read into passages, and in the language
  // the planner said it is written in. A deck without one is changed as before.
  const writtenIn = readLanguageTag(deck.documentLanguage);
  const given = typeof deck.document === 'string' && deck.document.trim() ? readDocument(deck.document, currentLanguage, writtenIn) : undefined;
  const docNow = given && given.passages.length > 0 ? given : undefined;
  const passageCount = docNow?.passages.length ?? 0;
  /** The passages a slide of the deck carries. */
  const passagesOf = (slide: SlideSpec | undefined): number[] => (docNow && slide ? readPassageNumbers(slide.passages, passageCount) : []);

  // What new slides may use. Number patterns need figures from the instruction or the original brief.
  const numbers = options.numberPatterns ?? 'from-brief';
  const colorsGiven = [options.colors, deck.colors, ...deck.slides.map((slide) => slide.colors)];
  // The figures of the document the deck was made from are the brief's too.
  const figures = numbers === 'always' || (numbers === 'from-brief' && hasFigures(withoutColorCodes(`${instruction} ${deck.brief ?? ''}${docNow ? ` ${docNow.text}` : ''}`, ...colorsGiven)));
  const context: CatalogContext = {
    layout,
    fields,
    icons: options.icons ?? [],
    assets: options.assets ?? [],
    imageDensity: options.imageDensity ?? 'medium',
    figures,
    textScale: deck.textScale ?? 1,
    templateScale: manifest.textScale ?? 1,
  };
  // What a writer is told counts the template's own spacing, on its layouts and the deck's.
  if (options.countSpacing !== false) context.told = toldCounter(template);
  const currentSize = (Object.entries(TEXT_SIZES).find(([, factor]) => factor === context.textScale)?.[0] ?? 'normal') as TextSize;
  // The list the deck is limited to: the one given with this edit, or else the deck's own. Both
  // are read leniently, since only a new brief is held to its ids. It narrows nothing here: the
  // model is told to keep to it.
  const relimited = givenPatterns(manifest, options.deckPatterns, freeNamed);
  const lists = allowedPatterns(manifest, { app: options.patterns, person: relimited ?? deck.patterns }, false, freeNamed);
  if (lists.allow) context.allow = lists.allow;
  if (deck.layouts?.length) {
    context.own = deck.layouts.map((design) => design.id);
    if (context.allow) context.allow = [...context.allow, ...context.own];
  }
  const offer = buildCatalog(manifest, context);
  const offered = new Map(offer.patterns.map((pattern) => [pattern.def.id, pattern]));
  // An existing slide keeps its pattern even if that pattern would not be offered for a new slide.
  const existing = [...new Set(deck.slides.map((slide) => slide.pattern))].filter((id) => definitions.has(id) && !offered.has(id));
  // Free slides. A slide that lists its parts is one, unless the template has a pattern of its own under that id.
  const isFree = (slide: SlideSpec | undefined): slide is SlideSpec => slide !== undefined && isFreeSlide(slide) && !definitions.has(slide.pattern);
  const freeSlides = deck.slides.filter(isFree);
  // Whether a new slide, or one that changes its layout, may be a free slide: as for a new plan,
  // with the list the deck is limited to in the place of the brief's.
  const free = freeOnOffer(loaded, context, { direction: deck.direction ?? manifest.direction, textScale: context.textScale }, { option: options.free, app: lists.allow, limit: lists.limit }).offer;
  // A setting turns free slides off: no slide's text is set smaller, which a slide that is cut says.
  const freeOff = !free && freeTurnedOff(manifest, { option: options.free, app: lists.allow, limit: lists.limit });
  // A free slide the deck has stays one whatever is on offer, so the plan may name the kinds of
  // part the deck uses beside the kinds on offer.
  const freeIn = free !== undefined || freeSlides.length > 0;
  const used = kindsUsed(freeSlides, free?.kinds);
  const kinds: FreeBlockKind[] = used.length > 0 ? used : ['text'];
  const mostPhotos = MOST_PHOTOS + context.assets.filter((asset) => asset.kind === 'photo').length;
  const patternIds = [...offered.keys(), ...existing, ...(freeIn ? [FREE_PATTERN] : [])];
  if (patternIds.length === 0) throw new PlanError(`No pattern of template "${manifest.name}" can be used`);
  // Of the list the deck is limited to, what can be chosen here. The rest would only mislead.
  const limit = (lists.limit ?? []).filter((id) => patternIds.includes(id));

  // 1. Plan the edit.
  // Deck-level texts can change too ("translate the deck"), unless the caller fixes them.
  const fieldNames = options.fields ? [] : Object.keys(fields);
  // A free slide set dense carries a text that was asked for. Its line says how many words it has,
  // its size where that is small, and the most it holds, so that the plan keeps the text whole or
  // grows it within that. A deck without such a slide is described as before.
  const dense = deck.slides.filter((slide) => isFree(slide) && slide.arrange?.dense === true);
  let marks: Map<SlideSpec, string[]> | undefined;
  if (dense.length > 0) {
    const kit = freeKit(loaded);
    const templateSize = kit ? freeBodySize(kit, context.textScale) : undefined;
    const room = freeTextRoom(loaded, { direction: deck.direction ?? manifest.direction, textScale: context.textScale });
    const most = room ? tensDown(wordsOfChars(room.densest, currentLanguage)) : undefined;
    marks = new Map(
      dense.map((slide): [SlideSpec, string[]] => {
        const resolved = freeFor(loaded, deck, slide);
        const size = typeof resolved === 'string' ? undefined : resolved.plan.smallest;
        const small = size !== undefined && templateSize !== undefined && setSmall(size, templateSize);
        return [slide, [`${wordsIn(slide, currentLanguage)} words`, ...(small ? [`text at ${size} pt`] : []), ...(most !== undefined ? [`at most about ${most} words`] : [])]];
      }),
    );
  }
  // A slide that carries passages of the document says which, and the share of their words it
  // keeps where that is under nine tenths. The document itself comes after the brief, with the
  // passages no slide carries marked.
  const onDeck = new Set(deck.slides.flatMap(passagesOf));
  const carriedNow = docNow
    ? new Map(
        deck.slides.flatMap((slide): Array<[SlideSpec, string]> => {
          const own = passagesOf(slide);
          if (own.length === 0) return [];
          const share = docNow.same ? keptShare(docNow, own, slide.slots) : undefined;
          return [[slide, `${passageLabel(own)}${share !== undefined && share < 0.9 ? `, ${Math.round(100 * share)}% of its text kept` : ''}`]];
        }),
      )
    : undefined;
  options.onProgress?.({ stage: 'outline' });
  const planRequest = [
    deck.brief ? `The deck was written from this brief:\n${deck.brief}` : '',
    docNow ? documentBlock(docNow, { change: true, onNoSlide: new Set(docNow.passages.map((passage) => passage.number).filter((number) => !onDeck.has(number))) }) : '',
    `The deck as it is now${deck.title ? ` ("${deck.title}")` : ''}, in the language with BCP 47 tag "${currentLanguage}", body text size "${currentSize}":\n${describeDeck(deck, isFree, marks, carriedNow)}`,
    fieldNames.length > 0
      ? `Deck-level texts, shown on several slides:\n${fieldNames.map((name) => `${name}: ${JSON.stringify(fields[name])}`).join('\n')}`
      : '',
    `Requested change:\n${instruction.trim()}`,
    [
      'Return the slides of the deck after the change, in their new order.',
      `- ${wishesRule('request')}`,
      '- from: the number of the existing slide a slide is based on, or 0 for a new slide. Leave a slide out to delete it.',
      '- pattern: keep the pattern of an existing slide unless the request calls for a different layout.',
      `  ${SLIDE_KIND_RULE}`,
      // The phrases the rule quotes are in the list of patterns, so it is given only with that list.
      ...(offer.patterns.length > 0 ? [`  ${columnsRule('request')}`] : []),
      ...(limit.length > 0 ? [`- ${limitRule(limit)}`] : []),
      '- change: what to write or change in the content of that slide. Leave it empty for every slide whose content the request leaves alone.',
      '- Colours are set by the software, not by you: for a wish about colour leave change empty.',
      `- ${colorNote(options.colors !== undefined)}`,
      '- A wish about how the deck looks (more modern, less empty, less text) is met with another pattern for a slide, another number of items, or the text size. It changes layouts and settings, not wording: leave change empty unless the slide gets another pattern.',
      // The rule above is about the look of a deck. A person who asks for other words on a slide still gets them.
      '- A request to shorten, cut or rewrite the text of a slide is a change of its content: say it in change.',
      '- images: true only when the pictures or icons of that slide should change.',
      `- words: ${wordsRule('request')}`,
      ...(docNow ? [`- ${DOCUMENT_EDIT_RULE}`, `- reword: ${REWORD_NOTE}`] : []),
      ...(marks ? [`- ${DENSE_EDIT_RULE}`] : []),
      '- items: for a pattern with a row of items, the number the slide has after the change.',
      ...(freeIn ? [`- ${PARTS_RULE}`, `- ${MOVED_RULE}`] : []),
      // Only where a slide may become a free slide: its chart is then handed on by code, if the plan lets it.
      ...(free ? [`- ${TURN_RULE}`] : []),
    ].join('\n'),
    offer.patterns.length > 0 || freeIn
      ? `Patterns, for new slides and for a slide that changes its layout ("?" marks an optional slot):\n${[
          describeCatalog(offer, budgetScale(context)),
          // Free slides have a line of their own, after the template's patterns. Where none is on
          // offer for a new slide, the line is there for the slides the deck has.
          ...(freeIn ? [freeLine(kinds, free ? { comfortable: free.comfortable, alone: offer.patterns.length === 0, asked: 'request', changes: 'deck' } : { closed: true, changes: 'deck' })] : []),
        ]
          .filter(Boolean)
          .join('\n')}`
      : '',
  ]
    .filter(Boolean)
    .join('\n\n');

  // `final` says no further answer will be asked for: what is wrong with the parts of a free
  // slide is then put right by code, and is no problem any more.
  const readPlan = (
    answer: unknown,
    final = false,
  ): { summary: string; language: string; title: string; textSize: TextSize; fields: Record<string, string>; entries: PlanEntry[]; problem?: string } => {
    const raw = (answer ?? {}) as { summary?: unknown; language?: unknown; title?: unknown; slides?: unknown; fields?: unknown; text_size?: unknown };
    const rawFields = (typeof raw.fields === 'object' && raw.fields !== null ? raw.fields : {}) as Record<string, unknown>;
    // A field the model leaves empty keeps its value: a header must not silently disappear.
    const nextFields = Object.fromEntries(fieldNames.map((name) => [name, cleanText(rawFields[name]) || (fields[name] as string)]));
    const countProblems: string[] = [];
    const entries: PlanEntry[] = (Array.isArray(raw.slides) ? raw.slides : []).map((item, index) => {
      const given = (item ?? {}) as { from?: unknown; pattern?: unknown; change?: unknown; images?: unknown; items?: unknown; words?: unknown; passages?: unknown; reword?: unknown };
      const entry: PlanEntry = { from: Number(given.from) || 0, pattern: cleanText(given.pattern), change: cleanText(given.change), images: given.images === true };
      const words = readWords(given.words);
      if (words !== undefined) entry.words = words;
      // The passages a slide carries after the change, where the plan names them: a list that is no list keeps those it has.
      if (docNow && Array.isArray(given.passages)) entry.passages = readPassageNumbers(given.passages, passageCount);
      if (docNow && given.reword === true) entry.reword = true;
      // What an entry says of parts counts for a free slide alone; on any other slide it is passed over.
      if (freeIn && entry.pattern === FREE_PATTERN) {
        const based = Number.isInteger(entry.from) && entry.from > 0 ? deck.slides[entry.from - 1] : undefined;
        entry.free = readParts(given, kinds, isFree(based) ? ((based.blocks ?? []) as FreeBlockKind[]) : undefined);
      }
      const definition = definitions.get(entry.pattern);
      if (definition?.repeat) {
        // A slide that keeps its layout keeps its number of items unless the plan says otherwise.
        const source = entry.from > 0 ? deck.slides[entry.from - 1] : undefined;
        const keeps = source?.pattern === entry.pattern;
        // A slide that keeps its layout and states no number, in the plan or on the slide, shows
        // what the template draws. It is given no number here, so that it is copied as it is.
        if (!(keeps && (given.items === undefined || given.items === null) && source?.items === undefined)) {
          const count = readItems(definition, given.items ?? (keeps ? source?.items : undefined));
          if (count.problem) countProblems.push(`slide ${index + 1}: ${count.problem}`);
          if (count.items !== undefined) entry.items = count.items;
        }
      }
      return entry;
    });
    const plan = {
      summary: cleanText(raw.summary),
      language: cleanText(raw.language) || currentLanguage,
      title: cleanText(raw.title),
      // "constructor" and "toString" are in every object and no sizes: only a size of the list is taken.
      textSize: (typeof raw.text_size === 'string' && Object.hasOwn(TEXT_SIZES, raw.text_size) ? raw.text_size : currentSize) as TextSize,
      fields: { ...fields, ...nextFields },
      entries,
    };
    if (entries.length === 0) return { ...plan, problem: 'the answer contained no slides' };
    if (countProblems.length > 0) return { ...plan, problem: countProblems.join('; ') };
    for (const [index, entry] of entries.entries()) {
      const where = `slide ${index + 1}`;
      if (!Number.isInteger(entry.from) || entry.from < 0 || entry.from > deck.slides.length) {
        return { ...plan, problem: `${where}: "from" must be 0 or the number of an existing slide (1 to ${deck.slides.length})` };
      }
      if (!definitions.has(entry.pattern) && !entry.free) return { ...plan, problem: `${where}: unknown pattern "${entry.pattern}"` };
      const source = entry.from > 0 ? deck.slides[entry.from - 1] : undefined;
      // A slide keeps the layout it has, whatever is on offer. A free slide is one by its parts, not by the id alone.
      const stays = entry.free ? isFree(source) : source?.pattern === entry.pattern;
      if (!stays && !(entry.free ? free !== undefined : offered.has(entry.pattern))) {
        return { ...plan, problem: `${where}: pattern "${entry.pattern}" cannot be used for a new or re-laid-out slide here` };
      }
      if (!source && !entry.change) return { ...plan, problem: `${where}: a new slide needs "change" to say what it shows` };
    }
    // The parts of the free slides: checked against the slides they come from, and put right where no answer will follow.
    const partProblems: string[] = [];
    entries.forEach((entry, index) => {
      if (!entry.free) return;
      const source = entry.from > 0 ? deck.slides[entry.from - 1] : undefined;
      const settled = settleParts(entry.free, index + 1, isFree(source) ? ((source.blocks ?? []) as FreeBlockKind[]) : undefined, mostPhotos, final);
      entry.free = { ...entry.free, parts: settled.parts };
      if (settled.problem) partProblems.push(settled.problem);
    });
    return partProblems.length > 0 ? { ...plan, problem: partProblems.join('; ') } : plan;
  };

  // Rows that take one number are counted in when some row varies: the list names their number, so it must be allowed.
  const itemRanges = patternIds.flatMap((id) => (definitions.has(id) ? [itemRange(definitions.get(id) as PatternDef)] : []));
  const mostItems = itemRanges.some((range) => range && range.max > range.min) ? Math.max(0, ...itemRanges.map((range) => range?.max ?? 0)) : 0;
  const planSchema = editPlanSchema(patternIds, deck.slides.length, fieldNames, mostItems, freeIn ? kinds : undefined, docNow ? passageCount : undefined);
  // The edit plan says a little more for each slide than an outline does, and may add slides.
  // Where slides may be free slides, each says how it is set out, and one whose parts change names
  // every one of them. Of a deck made from a document, a slide may name its passages.
  const partsNow = freeSlides.reduce((sum, slide) => sum + (slide.blocks?.length ?? 0), 0);
  const room =
    planRoom(deck.slides.length + 10) * 1.5 + (freeIn ? FREE_ENTRY_ROOM * (deck.slides.length + 10) + partsRoom(partsNow, instruction) : 0) + (docNow ? 8 * (deck.slides.length + 10) : 0);
  const planSystem = systemPrompt(currentLanguage, 'deck', docNow !== undefined, options.delivery, options.emphasis !== false);
  let plan = readPlan(await asker(options.llm, planSystem, options.signal, 'low')(planRequest, planSchema, room));
  if (plan.problem) {
    const retry = `${planRequest}\n\nYour previous answer could not be used: ${plan.problem}. Answer again.`;
    plan = readPlan(await asker(options.llm, planSystem, options.signal, 'low')(retry, planSchema, room), true);
    if (plan.problem) throw new PlanError(`The model did not produce a usable edit plan: ${plan.problem}`);
  }

  const language = plan.language;
  const direction = options.direction ?? (language === currentLanguage ? (deck.direction ?? directionOf(language)) : directionOf(language));
  // Of a deck made from a document: the document as the deck's language after the change has it,
  // the passages each slide carries, and the slides whose passages changed, which are written again.
  const doc = docNow ? inLanguage(docNow, language, writtenIn) : undefined;
  const passagesChanged = doc ? resolveDeckPassages(plan.entries, deck, passageCount) : new Set<PlanEntry>();
  for (const entry of passagesChanged) if (!entry.change) entry.change = entry.passages ? FROM_DOCUMENT : GAVE_UP;
  // Whether a slide was titled with the heading its text opened under, or had no title: a slide
  // that is new has none.
  const headed = (before: SlideSpec | undefined): boolean => {
    const title = before?.slots.title;
    if (!doc || !before || title === undefined) return true;
    return typeof title === 'string' && title.trim() === openingSection(doc.passages, before.passages ?? [])?.trim();
  };
  // Whether code sets the title of a slide whose text it sets again from its passages to the heading
  // that text opens under (`headingTitle`): where the change is to its passages alone and the slide
  // was so titled. A change in words may speak of the title, and a title the person gave stays: the
  // writer then titles the slide, offered the heading where the slide had it.
  const retitles = (entry: PlanEntry, before: SlideSpec | undefined): boolean => !!doc && entry.change === FROM_DOCUMENT && headed(before);
  // Words for a slide with passages ask for a shorter text only under four fifths of their words.
  // Otherwise the slide carries its passages whole.
  if (doc) {
    for (const entry of plan.entries) {
      if (entry.passages && entry.words !== undefined && slideDocument(doc, entry.passages, entry.words, language)?.asked === undefined) delete entry.words;
    }
  }
  // The slides the plan itself has words for. A change of language gives every other slide words too, by code.
  const worded = new Set(plan.entries.filter((entry) => entry.change !== ''));
  // A change of language touches every slide, whatever the plan says.
  const translating = language !== currentLanguage;
  if (translating) {
    for (const entry of plan.entries) if (!entry.change) entry.change = 'Translate into the new deck language.';
  }

  // Larger text fits less: from here on, limits are those of the size the deck will have.
  const textScale = TEXT_SIZES[plan.textSize];
  const resized = textScale !== context.textScale;
  context.textScale = textScale;

  // A slide the request asks an amount of text for gets the room for it, on that slide: it keeps
  // its layout when that holds the text, or moves to a layout of plain text or a free slide that
  // does, its text alone set dense where it needs it. Code adds no slide: where nothing holds the
  // text, the slide holds what it can, and the deck is told.
  const freeKinds = new Set<FreeBlockKind>(free?.kinds ?? []);
  const plainIds = offer.patterns.filter((pattern) => isPlainText(pattern.def)).map((pattern) => pattern.def.id);
  const lengthPlanning: LengthPlanning = {
    context,
    language,
    layout: (id, items) => {
      const definition = definitions.get(id);
      // A slide that stays in its layout can always be written again, even where that layout is not on offer for new slides.
      const planned = definition ? planPattern(definition, offered.has(id) ? context : { ...context, figures: true }, undefined, items) : undefined;
      return planned && 'pattern' in planned ? planned.pattern : undefined;
    },
    plainText: (id) => plainIds.includes(id),
    plain: plainIds,
    free: freeKinds.has('text')
      ? (blocks, arrange, lengths, paragraphs) => {
          const written = freeWritePattern(loaded, { direction, textScale }, { blocks, arrange, lengths, ...(paragraphs ? { paragraphs } : {}) }, context);
          return 'pattern' in written ? written.pattern : undefined;
        }
      : undefined,
    freeParts: (id) => freePartsOf(definitions.get(id), freeKinds),
  };
  // A slide set dense keeps the length of its text when that text is written again and the plan
  // gives it no words: "simpler words on slide 2" must not plan it for a usual length and cut the
  // text that was asked for. A change that keeps its text as it is (a better title) leaves it so.
  // No one asked for that length now: where the slide holds less (larger text), its text is
  // written to what it holds, and the change is not stopped for it (see toldLength).
  const keptLength = new Set<PlanEntry>();
  // A slide set dense, or one that carries passages, that keeps its text while a part is added
  // beside it has less room for that text. It is given the room for the text it has, which may be
  // shorter than its passages: a text that still holds beside the new part stays as it is, unless
  // the plan writes it again (`asIs`), and one that does not is written for the room it has, and is
  // a cut. In another language, or with other passages, the text is written anew all the same.
  const besides = new Map<PlanEntry, { from: number; chars: number; asIs: boolean }>();
  for (const entry of plan.entries) {
    const source = entry.from > 0 ? deck.slides[entry.from - 1] : undefined;
    if (entry.words !== undefined || !isFree(source)) continue;
    // A slide with passages has the length of its passages, and one that gives them up has no text
    // of that length to keep.
    const carried = entry.passages !== undefined;
    if (!carried && (passagesOf(source).length > 0 || source.arrange?.dense !== true)) continue;
    const named = entry.free?.parts ?? [];
    const old = (source.blocks ?? []) as FreeBlockKind[];
    const text = named.find((part) => part.kind === 'text' && old[part.from - 1] === 'text');
    const had = text ? source.slots[`b${text.from}Text`] : undefined;
    const chars = typeof had === 'string' || Array.isArray(had) ? toParagraphs(had).join('').length : 0;
    if (text && chars > 0 && !translating && !passagesChanged.has(entry) && named.some((part) => part.from === 0)) {
      besides.set(entry, { from: text.from, chars, asIs: !text.rewrite });
      if (!carried) entry.words = Math.max(1, wordsOfChars(chars, language));
      continue;
    }
    if (carried) continue;
    const words = readWords(wordsIn(source, currentLanguage));
    const again = translating || (worded.has(entry) && (named.length === 0 || named.some((part) => part.kind === 'text' && part.rewrite)));
    if (!again || words === undefined) continue;
    entry.words = words;
    keptLength.add(entry);
  }
  /** Whether a slide of a layout has a photo: a filled image slot of a photo. A free slide has no layout of the template. */
  const filledPhoto = (slide: SlideSpec | undefined): boolean => {
    const definition = slide ? definitions.get(slide.pattern) : undefined;
    return slide !== undefined && (definition?.slots ?? []).some((slot) => slot.kind === 'image' && slot.imageKind === 'photo' && isImageRef(slide.slots[slot.name]));
  };
  /**
   * Whether the running text of a slide is written again: its passages changed, the deck is
   * translated, words are asked for it, a part is added beside it, or the plan changes it and names
   * no part of it, or names its text to be written again. A slide of a layout, or one that becomes
   * a free slide, is written whole.
   */
  const textAgain = (entry: PlanEntry): boolean => {
    if (passagesChanged.has(entry) || translating || entry.words !== undefined || besides.has(entry)) return true;
    if (!entry.change) return false;
    const source = entry.from > 0 ? deck.slides[entry.from - 1] : undefined;
    if (!entry.free || !isFree(source)) return true;
    return entry.free.parts.length === 0 || entry.free.parts.some((part) => part.kind === 'text' && part.rewrite);
  };
  // The slides whose text from the document is written again, with what they carry of it. A slide
  // whose text stays (a better title) keeps it as it is.
  const carries = new Map<PlanEntry, SlideDocument>();
  if (doc) {
    for (const entry of plan.entries) {
      // A table of its passages goes into a table where the slide has one after the change, and is running text where it has none.
      const source = entry.from > 0 ? deck.slides[entry.from - 1] : undefined;
      const blocks = entry.free ? (entry.free.parts.length > 0 ? entry.free.parts.map((part) => part.kind) : isFree(source) ? ((source.blocks ?? []) as FreeBlockKind[]) : []) : undefined;
      const tables = blocks ? blocks.includes('table') : (definitions.get(entry.pattern)?.slots.some((slot) => slot.kind === 'table') ?? false);
      const carry = entry.passages && textAgain(entry) ? slideDocument(doc, entry.passages, entry.words, language, tables) : undefined;
      // A text a part is added beside carries what it has, where that is less than its passages: a text asked to be shorter, say.
      const kept = besides.get(entry);
      if (carry) carries.set(entry, kept && kept.chars < carry.need ? { ...carry, need: kept.chars, words: Math.max(1, wordsOfChars(kept.chars, language)) } : carry);
    }
  }
  /** The words a slide is to carry: those of its passages, or those asked for it. */
  const wordsFor = (entry: PlanEntry): number => carries.get(entry)?.words ?? (entry.words as number);
  /** The characters of text a slide is given the room for: its text from the document, the text a part is added beside, or the words asked for it. */
  const needFor = (entry: PlanEntry, carry: SlideDocument | undefined): number | undefined =>
    carry ? carry.need : (besides.get(entry)?.chars ?? (entry.words !== undefined ? charsOfWords(entry.words, language) : undefined));
  // The slides given the room for their text, and set dense where it needs that. A slide with words
  // always is. A slide with passages is where they changed, where it is asked to be shorter, where
  // a part is added beside its text, or where the deck is translated: one whose passages stay was
  // fitted when it was planned, and a better title does not move it to another layout.
  const homed = new Set<PlanEntry>();
  // Slides whose photo gives way to a text set dense.
  const dropped: PlanEntry[] = [];
  /**
   * Whether a slide whose passages stay is written again in other words where its layout holds its
   * text only as code sets it (`holdsOnlyAsIs`): its writer is told less than the text, so it is
   * given the room for it, as a slide whose passages changed is.
   */
  const outgrown = (entry: PlanEntry, found: SlideDocument): boolean => {
    if (!entry.reword || passagesChanged.has(entry) || entry.free) return false;
    const asIs = asIsOf(found);
    const pattern = asIs ? lengthPlanning.layout(entry.pattern, entry.items) : undefined;
    return asIs !== undefined && pattern !== undefined && holdsOnlyAsIs(pattern, context, found.shape, found.need, asIs);
  };
  for (const entry of plan.entries) {
    const found = carries.get(entry);
    const carry = found && (passagesChanged.has(entry) || translating || found.asked !== undefined || besides.has(entry) || outgrown(entry, found)) ? found : undefined;
    const need = needFor(entry, carry);
    if (need === undefined) continue;
    homed.add(entry);
    const source = entry.from > 0 ? deck.slides[entry.from - 1] : undefined;
    // The parts a free slide has after the change: those the plan names, or else those it has. A
    // plan that names none ("[]") and has words for the slide writes them all again, as a change
    // that names no part does (see planFree): that stays so if a running text is added to them.
    const named = entry.free?.parts.length ? entry.free.parts : undefined;
    const partsNow: PartPlan[] = named ?? (entry.free && isFree(source) ? ((source.blocks ?? []) as FreeBlockKind[]).map((kind, at) => ({ from: at + 1, kind, rewrite: true })) : []);
    const current: TextHome = entry.free ? { blocks: partsNow.map((part) => part.kind) } : { pattern: entry.pattern, items: entry.items };
    // A layout counts the text's own paragraph lengths only where code sets it there as it is (see `fill` below).
    const setByCode = carry !== undefined && passagesChanged.has(entry) && !entry.reword && asIsOf(carry) !== undefined;
    const shape = carry && (setByCode ? carry.shape : { ...carry.shape, layoutAsIs: false as const });
    const placed = homeForChars(current, need, lengthPlanning, { shape, dense: true });
    // A free text alone that is only set dense now keeps the parts the plan gave it.
    if (!sameHome(undense(placed.home), current)) {
      if (isDenseHome(placed.home)) {
        // A text alone, set dense. Such a home is found only for a slide of one running text and
        // photos, so no text or list of the slide is lost: the text is written again, the photos go.
        const text = partsNow.find((part) => part.kind === 'text');
        entry.free = { ...(entry.free ?? {}), parts: [{ from: text?.from ?? 0, kind: 'text', rewrite: true }] };
        entry.pattern = FREE_PATTERN;
        delete entry.items;
        if (partsNow.some((part) => part.kind === 'photo') || filledPhoto(source)) dropped.push(entry);
      } else if ('blocks' in placed.home) {
        // A free slide that gets a running text keeps its parts after it. A slide of a layout
        // becomes a free slide of a running text and its chart, table or photo, which it hands
        // on as they are (not written again: see carryOver); a new slide has none to hand on.
        const added = entry.free ? placed.home.blocks.length - partsNow.length : placed.home.blocks.length;
        const keeps = (kind: FreeBlockKind) => source !== undefined && !entry.free && (kind === 'chart' || kind === 'table' || kind === 'photo');
        const fresh = placed.home.blocks.slice(0, added).map((kind): PartPlan => ({ from: 0, kind, rewrite: !keeps(kind) }));
        entry.free = { ...(entry.free ?? {}), parts: [...fresh, ...(entry.free ? partsNow : [])] };
        entry.pattern = FREE_PATTERN;
        delete entry.items;
      } else {
        entry.pattern = placed.home.pattern;
        delete entry.free;
        if (placed.home.items !== undefined) entry.items = placed.home.items;
        else delete entry.items;
      }
    }
    if (isDenseHome(placed.home)) entry.dense = true;
    // The text is written whatever else the plan says of the slide. A part added beside a text has
    // nothing said of it but the request, where the plan says nothing.
    if (!entry.change && !besides.has(entry)) entry.change = carry ? FROM_DOCUMENT : 'Write the text at the length asked for.';
  }
  // The deck is told of a photo that gave way: the summary of the change was written before, and says nothing of it.
  // Beside a text no one asked for now, it gives way to the text the slide has.
  const leftOut: Issue[] = dropped.map((entry) => photoLeftOut(plan.entries.indexOf(entry) + 1, wordsFor(entry), !besides.has(entry)));
  // A deck that holds a free slide is checked with the template as it was loaded: the slide is
  // then laid out, and its text checked against the room it gets. Any other deck is checked as before.
  const check = (given: DeckSpec): Issue[] =>
    given.slides.some((slide) => isFreeSlide(slide)) ? validateDeck(loaded, given, { layout }) : validateDeck(manifest, given, measure ? { layout, measure } : { layout });
  /** What no longer fits on a slide at the size the deck will have. */
  const tight = (slide: SlideSpec): Issue[] => check({ ...deck, textScale, slides: [slide] }).filter((issue) => /characters|paragraphs|lines/.test(issue.message));
  const overflows = (slide: SlideSpec) => tight(slide).length > 0;

  // Slides that hold less text than the request asked for them, said with the deck.
  const short: Issue[] = [];
  // Slots told the whole of the length asked for their slide, held to it once they are written.
  const told: ToldLength[] = [];
  // Slides written from their passages of the document, held to its words.
  const keptChecks: KeptCheck[] = [];
  // Slides given the room for their text that hold less than they carry, as they are written: with
  // the heading or caption they keep. Where the caller asks, the change stops on them before writing.
  const cuts: SlideTextFit[] = [];
  const cutAt = (entry: PlanEntry, number: number, words: number, need: number, room: number, size: number | undefined): void => {
    if (!homed.has(entry)) return;
    const label = carries.get(entry)?.label;
    cuts.push({
      slide: number,
      words,
      holds: Math.floor((words * room) / need),
      ...(size !== undefined ? { size } : {}),
      ...(entry.dense ? { dense: true as const } : {}),
      ...(label !== undefined ? { passages: label } : {}),
    });
  };
  /**
   * The pattern a slide is written against, told its text from the document where its text from
   * the document is written again, or else the length asked for it where there was one. A slide set
   * dense (`dense`) that still holds less says the size its text is set at.
   */
  const toldLength = (pattern: UsablePattern, number: number, entry: PlanEntry, names?: ReadonlySet<string>, dense = false): UsablePattern => {
    const carry = carries.get(entry);
    const size = dense ? runningSize(pattern, names) : undefined;
    if (carry && doc) {
      // Its writer writes the text (code sets it, where it can, before this): it is held to what the writer is told.
      const room = documentRoom(pattern, context, carry.shape, names);
      // A slide written again only because the deck is translated, whose passages and parts stay and
      // of which no length was asked, held its text: it is not stopped as a cut. In another script
      // than the document's, its length is a guess from the words of its passages, and it is
      // written to what it holds; in the document's own, a slide that holds less is said after.
      const translated = translating && !passagesChanged.has(entry) && carry.asked === undefined;
      const need = translated && !doc.same ? Math.min(carry.need, room) : carry.need;
      if (room < need) {
        short.push(size !== undefined ? textCutDense(number, carry.label, need, room, size) : textCut(number, carry.label, need, room, freeOff));
        if (!translated) cutAt(entry, number, carry.words, need, room, size);
      }
      // A slide whose passages changed is held to their words. One the request has written otherwise is not.
      const check = passagesChanged.has(entry) && !entry.reword ? keptCheck(doc, number, carry, room, runningNames(pattern, names)[0]) : undefined;
      if (check) keptChecks.push(check);
      return withPassages(pattern, need, context, { names, language, label: carry.label, shape: carry.shape, words: carry.asked });
    }
    if (entry.words === undefined) return pattern;
    const room = heldBy(pattern, context, { names, language });
    // A length kept by code, which no one asked for now, is what the slide holds where that is less.
    const kept = keptLength.has(entry);
    const words = kept ? Math.min(entry.words, Math.max(1, wordsOfChars(room, language))) : entry.words;
    if (!kept && room < charsOfWords(words, language)) {
      short.push(size !== undefined ? shortOfWordsDense(number, words, room, size, language) : shortOfWords(number, words, room, language, freeOff));
      cutAt(entry, number, words, charsOfWords(words, language), room, size);
    }
    const whole = toldAsAsked(pattern, words, context, { names, language });
    if (whole) told.push({ slide: number, words, ...whole });
    return withLength(pattern, words, context, { names, language });
  };

  // 2. Work out, per slide, whether it is copied or written, and what a written slide carries over.
  const jobs: EditJob[] = [];
  const kept = new Map<number, SlideSpec>();
  // Free slides whose parts or arrangement code changes, with nothing to write.
  const reshaped: number[] = [];
  // Free slides the plan had words for and on which nothing changed, with nothing to write.
  const idle: number[] = [];

  /**
   * A slide that is a free slide after the change. One that was a free slide keeps the parts
   * the plan does not touch, by code; only what the plan says is new or to be written again goes
   * to the writer. One that was a slide of a pattern, or is new, is written whole.
   */
  const planFree = (entry: PlanEntry, number: number, source: SlideSpec | undefined): void => {
    const before = isFree(source) ? source : undefined;
    const stated: FreeChange = entry.free ?? { parts: [] };
    // Its text from the document, where that is written again.
    const carry = carries.get(entry);
    // A length asked for the slide is the length of its running texts, or of its lists where it
    // has no running text (as `freeLengths` shares it out): those are written again. So are they
    // where its text from the document is written again.
    const holder: FreeBlockKind = stated.parts.some((part) => part.kind === 'text') ? 'text' : 'list';
    let parts: PartPlan[] = entry.words !== undefined || carry ? stated.parts.map((part) => (part.kind === holder ? { ...part, rewrite: true } : part)) : stated.parts;
    // The title is written with the parts, and whenever the plan has words for the slide: where
    // nothing is written and nothing is said, it stays.
    let title: boolean | undefined;
    // Whether the parts that are written again were named for it outright, by the plan or by a
    // change of language. What such a part must have is then asked for in so many words.
    let outright = false;
    // Whether the plan changes the slide: it has words for it, or its parts change. Only then do
    // `auto` and 0 take an arrangement away. On a slide the plan leaves alone they state nothing,
    // also where code writes the slide again (another language, text that no longer fits).
    let changed = true;
    // The plan has words for the slide and names every part as it is, none to be written again.
    let wordsOnly = false;

    if (before) {
      const old = (before.blocks ?? []) as FreeBlockKind[];
      const whole = (rewrite: (part: number) => boolean): PartPlan[] => old.map((kind, at) => ({ from: at + 1, kind, rewrite: rewrite(at + 1) }));
      changed = worded.has(entry) || !sameParts(parts, old);
      if (sameParts(parts, old) && !entry.change) {
        const long = resized ? tight(before) : [];
        if (long.length === 0) {
          // Left alone: copied with its parts and its frame. An arrangement the plan states is set by code, and nothing is written.
          const copy = structuredClone(before);
          const arrange = statesArrange(stated, false) ? arrangeAfter(before.arrange, stated, old.length, false) : before.arrange;
          if (!sameArrange(arrange, before.arrange)) {
            if (arrange) copy.arrange = arrange;
            else delete copy.arrange;
            reshaped.push(number);
          }
          kept.set(number, copy);
          return;
        }
        // Text that no longer fits at the new size: the parts it is in are written again, and no other.
        entry.change = SHORTEN;
        const slots = new Set(long.map((issue) => issue.slot));
        const blocks = new Set([...slots].map((slot) => partOf(slot)?.block));
        parts = whole((part) => blocks.has(part));
        title = slots.has('title');
      } else if (translating) {
        // Another language: every part the slide keeps is written again, whatever the plan says of it, and the title with them.
        parts = parts.length === 0 ? whole(() => true) : parts.map((part) => ({ ...part, rewrite: true }));
        title = true;
        outright = true;
      } else if (parts.length === 0) {
        // The content changes and the plan names no part: each is offered with what it holds, and the title with them.
        parts = whole(() => true);
        title = true;
      } else {
        // The plan names the parts. Those it marks are written, and no other: a part named as it
        // is, not to be written again, is kept by code also where the slide gets a new title.
        outright = true;
        // With every part named as it is, the words can only be about the title. Should that
        // come back as it was, nothing on the slide changed: the deck is told, and no call is added.
        wordsOnly = worded.has(entry) && sameParts(parts, old);
      }
    }

    // A slide that names a part of no kind cannot be laid out. Left alone it was copied above, and the check of the deck speaks of it.
    if (before && !(before.blocks ?? []).every(isFreeBlockKind)) throw new PlanError(`Slide ${number}: "blocks" names a kind of part that does not exist, so the slide cannot be changed`);

    let arrange = arrangeAfter(before?.arrange, stated, parts.length, changed);
    // A slide given the room for its text is set dense where its text needed that, and at the
    // template's own sizes otherwise. Any other slide keeps what it had.
    if (homed.has(entry)) {
      const { dense: _dense, ...rest }: FreeArrange = arrange ?? {};
      const set: FreeArrange = entry.dense ? { ...rest, dense: true } : rest;
      arrange = Object.keys(set).length > 0 ? set : undefined;
    }
    // A text that a part is added beside: its length is its own, and a new text beside it is written as the request says.
    const beside = besides.get(entry);
    const besideAt = beside ? parts.findIndex((part) => part.from === beside.from && part.kind === 'text') : -1;
    // `filled` names a running text that code sets from the document: the writer is not asked for it.
    const prepare = (images: boolean, filled?: string, titled?: string, offered?: string) => {
      const moved = new Set<string>();
      const shape = before ? carryParts(before, parts, images) : carryOver(source, source ? definitions.get(source.pattern) : undefined, parts, images, moved);
      const skip = new Set(shape.skip);
      if (filled !== undefined) skip.add(filled);
      const carried = { ...shape.carried };
      // In the order the writer is asked for them: the title, then the parts.
      const current: Record<string, SlotValue> = {};
      const had = before?.slots.title;
      // A plan that has words for the slide and names no part to write can only mean the title:
      // it is what the plan has no other way to name. "Take the note off and call the slide
      // otherwise" then takes one small call, where the note alone takes none.
      if (titled !== undefined) {
        // The heading its text opens under, set by code with the text (`headingTitle`).
        skip.add('title');
        carried.title = titled;
      } else if (title ?? (shape.written.length > 0 || entry.change !== '')) {
        // The heading its text now opens under, where the slide had its heading for a title.
        const shown = offered ?? had;
        if (shown !== undefined && shown !== null) current.title = shown;
      } else {
        skip.add('title');
        if (had !== undefined) carried.title = had;
      }
      Object.assign(current, shape.current);
      // A chart or a table is asked of the writer here as for a new deck: required, and empty
      // where there are no numbers for it. One the slide has and that comes back empty keeps
      // what it had, and one the plan named to be written again is then asked for once more (`due`).
      // The slide is planned with the optional parts it has (a caption under its one photo, a
      // heading over its one text), so that they have their room and can be written again.
      // Its running texts are set out for the length asked, where one was, or for its text from the document.
      const need = needFor(entry, carry);
      const lengths = need === undefined ? undefined : besideAt >= 0 ? { [`b${besideAt + 1}Text`]: need } : freeLengths(shape.blocks, need);
      const names = new Set(Object.keys(lengths ?? {}));
      const planned = freeWritePattern(loaded, { direction, textScale }, { blocks: shape.blocks, arrange, base: before?.base, lengths, ...paragraphsFor(carry, names) }, context, skip, heldParts(shape));
      // A chart or a table that is offered with what it holds is asked for at the size it has.
      const written = 'reason' in planned ? planned : { ...planned, pattern: roomAsHeld(planned.pattern, current) };
      return { shape, moved, carried, current, written, names };
    };
    let images = entry.images;
    let made = prepare(images);
    // Its pictures cannot be replaced with what the caller supplied (say, images are turned off): keep them all.
    if ('reason' in made.written && source && entry.images) {
      images = false;
      made = prepare(false);
    }
    // A text whose passages changed, and that holds them, is set by code as the document has it:
    // the writer writes the title and the rest. One the request asks for in other words, it writes.
    let fill = doc && carry && passagesChanged.has(entry) && !entry.reword && 'pattern' in made.written ? filledText(doc, carry, made.written.pattern, context, made.names) : undefined;
    // A text that a part is added beside, which the plan keeps as it is, stays so where it still
    // holds as the slide is written: it is not written again.
    const had = beside?.asIs && besideAt >= 0 ? before?.slots[`b${beside.from}Text`] : undefined;
    if (beside && (typeof had === 'string' || Array.isArray(had)) && 'pattern' in made.written && heldBy(made.written.pattern, context, { names: made.names }) >= beside.chars) {
      fill = { slot: `b${besideAt + 1}Text`, value: had };
    }
    if (fill) {
      const heading = carry && doc && 'pattern' in made.written ? headingTitle(doc, carry, made.written.pattern, context, freeKit(loaded, made.written.base)) : undefined;
      const titled = heading !== undefined && retitles(entry, source) ? heading : undefined;
      const again = prepare(images, fill.slot, titled, titled === undefined && heading !== undefined && headed(source) ? heading : undefined);
      if ('pattern' in again.written) made = again;
      else fill = undefined;
    }
    const { shape, carried, current, written, names } = made;
    if ('reason' in written) throw new PlanError(`Slide ${number}: the free slide ${written.reason}`);

    // What the writer is not asked for stays as it is: an icon where no icon set came with this change.
    const asked = new Set(written.pattern.slots.map((planned) => planned.slot.name));
    for (const [name, value] of Object.entries(current)) {
      if (asked.has(name)) continue;
      carried[name] = value;
      delete current[name];
    }
    if (fill) carried[fill.slot] = fill.value;
    if (before && asked.size === 0) {
      // Parts taken off or moved, none to write, and no words from the plan: the slide is made by code from what it keeps, its title too.
      const had = before.slots.title;
      const slide: SlideSpec = {
        pattern: FREE_PATTERN,
        blocks: shape.blocks,
        ...(arrange ? { arrange } : {}),
        ...(before.base !== undefined ? { base: before.base } : {}),
        slots: { ...(had !== undefined ? { title: had } : {}), ...carried },
      };
      kept.set(number, slide);
      // Words for a slide whose every part stays, and no title to offer: unless it is set out anew, nothing on it changed.
      if (wordsOnly && sameArrange(arrange, before.arrange)) idle.push(number);
      else reshaped.push(number);
      return;
    }

    // Nothing of the slide as it was is kept: it is written like a slide that takes another layout, from what the old one held.
    const anew = before !== undefined && parts.every((part) => part.from === 0);
    const job: EditJob = {
      number,
      pattern: (entry.words !== undefined || carry) && !fill ? toldLength(written.pattern, number, entry, names, arrange?.dense === true) : written.pattern,
      // A plan that names new parts and says nothing of them leaves the request itself to go by.
      purpose: entry.change || instruction.trim(),
      carried,
      relaid: (!!source && !before) || anew,
      free: { blocks: shape.blocks, ...(arrange ? { arrange } : {}), base: written.base },
      // The purpose is what to do, which no slide should show: a slide left without a part says its title.
      task: true,
    };
    const deckTitle = plan.title || deck.title;
    if (deckTitle) job.fallback = deckTitle;
    if (source) job.source = source;
    if (fill) job.filled = fill.slot;
    // The writer is shown the slide's text from the document, to write it, or its title, from it.
    if (doc && entry.passages?.length) job.passage = { label: passageLabel(entry.passages), text: passageText(doc.passages, entry.passages) };
    if (Object.keys(current).length > 0) job.current = current;
    if (wordsOnly && before) job.asItWas = before;
    if (outright) {
      // What was named to be written again has to come back: the values such a part must have.
      // Left unsaid they stay as they are, and that is then asked about and told.
      const due = written.pattern.slots.filter((planned) => planned.required && partOf(planned.slot.name) !== undefined && current[planned.slot.name] !== undefined).map((planned) => planned.slot.name);
      if (due.length > 0) job.due = new Set(due);
    }
    if (job.relaid && source) {
      job.material = Object.fromEntries(
        Object.entries(source.slots)
          .filter(([name, value]) => value !== null && !made.moved.has(name))
          .map(([name, value]) => [name, shownValue(value)]),
      );
    }
    const told: string[] = [];
    // The writer is told what stays, by the same short lines the plan of the change was made from.
    // A part of which something is asked is not among "what is not to be written": what it keeps
    // (its icon, its caption) is said in a line of its own, or the two would contradict each other.
    const partsOf = (names: Iterable<string>) => new Set([...names].flatMap((name) => partOf(name)?.block ?? []));
    const writing = partsOf(asked);
    const staying = [...partsOf(Object.keys(carried))].sort((a, b) => a - b);
    const whole = staying.filter((part) => !writing.has(part));
    const partly = staying.filter((part) => writing.has(part));
    if (whole.length > 0) told.push(`What stays as it is and is not to be written: ${listParts(shape.blocks, carried, whole).join('; ')}`);
    if (partly.length > 0) told.push(`Of the parts that are written, this stays as it is: ${listParts(shape.blocks, carried, partly).join('; ')}`);
    if (Object.keys(shape.replaced).length > 0) told.push(`The new parts with these numbers take the place of parts that held: ${JSON.stringify(shape.replaced)}`);
    if (told.length > 0) job.notes = told;
    jobs.push(job);
  };

  plan.entries.forEach((entry, index) => {
    const number = index + 1;
    const source = entry.from > 0 ? deck.slides[entry.from - 1] : undefined;
    if (entry.free) {
      planFree(entry, number, source);
      return;
    }
    const definition = definitions.get(entry.pattern) as PatternDef;

    if (source && source.pattern === entry.pattern && !entry.change) {
      // Two things rewrite a slide the plan left alone: a new number of items, and text that no longer fits.
      if (entry.items !== undefined && entry.items !== itemCount(definition, source.items)) {
        entry.change = `Use ${entry.items} ${definition.repeat?.item ?? 'item'}s.`;
      } else if (resized && overflows(source)) {
        entry.change = SHORTEN;
      }
    }

    if (source && source.pattern === entry.pattern && !entry.change) {
      kept.set(number, structuredClone(source));
      return;
    }

    const carried: Record<string, SlotValue> = {};
    if (source && source.pattern === entry.pattern) {
      // Same layout: everything the model is not asked about stays, pictures included unless they are the point.
      for (const slot of effectiveSlots(definition, entry.items)) {
        const value = source.slots[slot.name];
        if (value === undefined) continue;
        const policy = slot.empty ?? 'error';
        if ((slot.kind === 'image' && !entry.images) || policy === 'keep' || policy === 'drop') carried[slot.name] = value;
      }
    } else if (source && !entry.images) {
      // New layout for an existing slide: its photos move over, in order, so they are not generated again.
      const oldDefinition = definitions.get(source.pattern);
      const wasFree = isFreeSlide(source) && !oldDefinition;
      const photos = wasFree
        ? freePhotos(source)
        : (oldDefinition?.slots ?? [])
            .filter((slot) => slot.kind === 'image' && slot.imageKind === 'photo')
            .map((slot) => source.slots[slot.name])
            .filter(isImageRef);
      for (const slot of effectiveSlots(definition, entry.items)) {
        if (slot.kind !== 'image' || slot.imageKind !== 'photo') continue;
        const photo = photos.shift();
        if (photo) carried[slot.name] = photo;
      }
    }

    // A slide that stays in its layout can always be rewritten, even if that layout is not on offer for new slides.
    const own: CatalogContext = source && source.pattern === entry.pattern ? { ...context, figures: true, imageDensity: 'high' } : context;
    let planned = planPattern(definition, own, new Set(Object.keys(carried)), entry.items);
    if ('reason' in planned && source && source.pattern === entry.pattern) {
      // Its pictures cannot be replaced with what the caller supplied (say, no icon set): keep them all.
      for (const slot of effectiveSlots(definition, entry.items)) {
        const value = source.slots[slot.name];
        if (slot.kind === 'image' && value !== undefined) carried[slot.name] = value;
      }
      planned = planPattern(definition, own, new Set(Object.keys(carried)), entry.items);
    }
    if ('reason' in planned) throw new PlanError(`Slide ${number}: pattern "${entry.pattern}" ${planned.reason}`);

    // A text whose passages changed, and that holds them, is set by code as the document has it,
    // unless the request asks for it in other words. So is one whose passages stay, that the slide
    // has as code set it, where its layout holds it only so (`holdsOnlyAsIs`): a change of its
    // title, say, does not hand it to a writer that would be told less than the text.
    const carry = carries.get(entry);
    const set = doc && carry && !entry.reword ? filledText(doc, carry, planned.pattern, context) : undefined;
    const asIs = carry ? asIsOf(carry) : undefined;
    const stays =
      !!set && !!carry && !!asIs && !!source && holdsOnlyAsIs(planned.pattern, context, carry.shape, carry.need, asIs) && Object.values(source.slots).some((value) => JSON.stringify(value) === JSON.stringify(set.value));
    const fill = set && (passagesChanged.has(entry) || stays) ? set : undefined;
    if (fill) carried[fill.slot] = fill.value;
    // Its title then is the heading its text opens under, where the title holds it and the change leaves it to the document.
    const titled = fill && !stays && doc && carry && retitles(entry, source) ? headingTitle(doc, carry, planned.pattern, context) : undefined;
    if (titled !== undefined) carried.title = titled;
    const job: EditJob = {
      number,
      pattern: fill ? { ...planned.pattern, slots: planned.pattern.slots.filter((one) => one.slot.name !== fill.slot && !(titled !== undefined && one.slot.name === 'title')) } : toldLength(planned.pattern, number, entry),
      purpose: entry.change,
      carried,
      relaid: !!source && source.pattern !== entry.pattern,
    };
    if (source) job.source = source;
    if (fill) job.filled = fill.slot;
    if (doc && entry.passages?.length) job.passage = { label: passageLabel(entry.passages), text: passageText(doc.passages, entry.passages) };
    jobs.push(job);
  });

  // A cut is said before anything is written, where the caller asks for that: the person can then
  // ask for more slides or a shorter text, or accept the cut. It is found in the patterns the slides
  // are written against, which count the heading or caption a slide keeps: no model was asked yet.
  if (options.stopBeforeCut && cuts.length > 0) {
    const slides = cuts.map((fit) => `slide ${fit.slide} would hold about ${fit.holds} of its ${fit.words} words${fit.size !== undefined ? ` (text at ${fit.size} pt)` : ''}`).join('; ');
    const off = freeOff && cuts.some((fit) => fit.size === undefined) ? ' Free slides are off, so the text is not set smaller.' : '';
    throw new TextCutError(`Nothing was changed: ${slides}. Ask for more slides or a shorter text, or accept that the rest is cut.${off}`, cuts, freeOff);
  }

  // The slides the caller's colours are for, when some were named. The caller numbered the deck
  // as it was; the plan says where each of those slides is now. A slide the edit adds (from 0) is
  // none of them, and one the edit removed cannot be coloured, which is said.
  const named = options.colors !== undefined && options.colorSlides?.length ? [...new Set(options.colorSlides)] : undefined;
  const recolored = named ? plan.entries.flatMap((entry, index) => (entry.from > 0 && named.includes(entry.from) ? [index + 1] : [])) : undefined;
  const removed = (named ?? []).filter((number) => Number.isInteger(number) && number >= 1 && number <= deck.slides.length && !plan.entries.some((entry) => entry.from === number));
  const notes: Issue[] = [...idle.map(nothingChanged), ...removed.map((number): Issue => ({ severity: 'warning', ...said('recolor-slide-removed', { slide: number }) })), ...leftOut, ...short];

  const assemble = (written: Map<number, SlideSpec>): DeckSpec => {
    const slides = plan.entries.map((entry, index) => {
      const slide = kept.get(index + 1) ?? written.get(index + 1) ?? { pattern: entry.pattern, slots: {} };
      // A slide's own colours travel with it: a slide that is rewritten or moved keeps them.
      const own = entry.from > 0 ? deck.slides[entry.from - 1]?.colors : undefined;
      const coloured = own && slide.colors === undefined ? { ...slide, colors: structuredClone(own) } : slide;
      if (!doc) return coloured;
      // Of a deck made from a document, each slide carries the passages the change gave it. A slide
      // written in this change is settled with them: their headings in its text are marked, a
      // heading in brackets is set plain, and a text does not open with its title. A slide the
      // change leaves alone keeps its text, and its headings, as they were.
      if (kept.has(index + 1)) {
        const { passages: _passages, ...rest } = coloured;
        return entry.passages?.length ? { ...rest, passages: [...entry.passages] } : rest;
      }
      const { passages: _passages, headings: _headings, ...rest } = coloured;
      return entry.passages?.length ? settleDocumentSlide({ ...rest, passages: [...entry.passages] }, doc.passages, entry.passages, { translated: !doc.same }) : rest;
    });
    // The request joins the brief: later edits, and the check of chart numbers, see what was asked for.
    const brief = [deck.brief, instruction.trim()].filter(Boolean).join('\n');
    const next: DeckSpec = { ...deck, version: 1, language, direction, brief, slides };
    if (textScale === 1) delete next.textScale;
    else next.textScale = textScale;
    if (plan.title) next.title = plan.title;
    if (Object.keys(plan.fields).length > 0) next.fields = plan.fields;
    // A layout made for the deck goes when no slide uses it any more.
    const layouts = usedLayouts(deck.layouts, slides);
    if (layouts) next.layouts = layouts;
    else delete next.layouts;
    // What the caller set outright is done by code, now that the plan says where each slide went.
    if (relimited?.length) next.patterns = relimited;
    else if (relimited) delete next.patterns;
    if (options.colors === undefined) return next;
    return recolored ? recolorSlides(next, storedColors(options.colors), recolored) : recolorDeck(next, options.colors);
  };
  const result = (edited: DeckSpec, rewritten: number[], warnings: Issue[]): EditResult => ({
    deck: edited,
    summary: plan.summary,
    rewritten,
    ...(recolored ? { recolored } : {}),
    warnings: [...warnings, ...notes],
  });

  if (jobs.length === 0) {
    // Reordering and deleting need no writing. What the check says of the deck as a whole is still
    // passed on: a colour set with this edit that is no colour would otherwise go unsaid. So is
    // what it says of a free slide whose parts or arrangement code changed.
    const edited = assemble(new Map());
    const whole = check(edited).filter((issue) => issue.severity === 'warning' && (issue.slide === undefined || reshaped.includes(issue.slide)));
    return result(edited, reshaped, whole);
  }

  // 3. Write the slides that change.
  options.onProgress?.({ stage: 'write', slides: jobs.length });
  const overview = plan.entries
    .map((entry, index) => {
      const number = index + 1;
      const keptSlide = kept.get(number);
      return keptSlide ? `${number}. ${entry.pattern}, unchanged ${slideLabel(keptSlide)}` : `${number}. ${entry.pattern}, to write`;
    })
    .join('\n');
  const materialLine = (job: EditJob): string[] => (job.material ? [`Content of the slide it replaces: ${JSON.stringify(job.material)}`] : []);
  /** The slide's text from the document, under its first line, where it carries passages. */
  const textOf = (job: EditJob): string[] => (job.passage ? [slideText(job.number, job.passage)] : []);
  const taskOf = (job: EditJob): string => {
    if (job.free) {
      // A free slide's slots are named by the numbers of its parts, so the line says which parts are what.
      const lines = [`s${job.number} (${job.pattern.def.id}: ${describeParts(job.free.blocks)}). ${job.source ? 'Change' : 'New slide'}: ${job.purpose}`, ...textOf(job), ...(job.notes ?? [])];
      // What the slots to write hold now, under the names they are written by. Nothing of the parts that stay is shown.
      if (job.current) lines.push(`Current content: ${JSON.stringify(Object.fromEntries(Object.entries(job.current).map(([name, value]) => [name, shownValue(value)])))}`);
      return [...lines, ...materialLine(job)].join('\n');
    }
    const offeredSlots = new Set(job.pattern.slots.map((planned) => planned.slot.name));
    const lines = [`s${job.number} (${job.pattern.def.id}). ${job.source ? 'Change' : 'New slide'}: ${job.purpose}`, ...textOf(job)];
    if (job.source) {
      // Show what is there now. After a layout change the slot names differ, so everything is shown as material.
      const current = Object.entries(job.source.slots).filter(([name, value]) => (job.relaid || offeredSlots.has(name)) && value !== null);
      const shown = Object.fromEntries(current.map(([name, value]) => [name, shownValue(value)]));
      lines.push(`${job.relaid ? 'Content of the slide it replaces' : 'Current content'}: ${JSON.stringify(shown)}`);
    }
    return lines.join('\n');
  };
  // A free slide of many parts is written by calls of its own. Every other slide is in one call, as before.
  const together = jobs.filter((job) => !writtenApart(job));
  const head = [deck.brief ? `The deck was written from this brief:\n${deck.brief}` : '', `Requested change:\n${instruction.trim()}`, `The deck after the change:\n${overview}`];
  // Slides with text from the document are written in its words: the rules for that come once, after the slides.
  const fromDocument = together.some((job) => job.passage) ? [documentWriteRule(together.some((job) => job.filled !== undefined))] : [];
  const request = [
    ...head,
    'Write only the slides below. When a slide has current content, start from it and change only what is asked: keep every other text word for word.',
    together.map(taskOf).join('\n\n'),
    ...fromDocument,
    ...suppliesText(together, context),
    'Fill every required property. Leave an optional property out when the slide is better without it.',
  ]
    .filter(Boolean)
    .join('\n\n');
  /** The prompt of a call that writes some parts of one slide by themselves: what was asked, the deck around the slide, then the task. */
  const requestFor = (job: EditJob, task: string): string =>
    [
      ...head,
      'Write only the parts asked for below. Where they have current content, start from it and change only what is asked: keep every other text word for word.',
      [task, ...(job.notes ?? []), ...materialLine(job)].join('\n'),
      ...(job.passage ? [slideText(job.number, job.passage), documentWriteRule(job.filled !== undefined)] : []),
      ...suppliesText([job], context),
    ]
      .filter(Boolean)
      .join('\n\n');

  const system = systemPrompt(language, 'deck', doc !== undefined, options.delivery, options.emphasis !== false);
  // A text written well short of the length asked for it goes back once, with the first answer's
  // problems, and so does a slide that keeps too few of the words of its text from the document.
  // The issues that send them are the writer's alone: the deck is told after the last answer.
  const sent = new WeakSet<Issue>();
  const lengthNotes = (written: DeckSpec, issues: Issue[], round: number): Issue[] =>
    round > 0
      ? []
      : [...writtenUnder(written, told).map(writtenLength), ...(doc ? keptIssues(doc, keptChecks, written, issues) : [])].map((issue) => {
          sent.add(issue);
          return issue;
        });
  const draft = await writeSlides({
    ask: asker(options.llm, system, options.signal),
    askLight: asker(options.llm, system, options.signal, 'low'),
    request,
    requestFor: (job, task) => requestFor(job as EditJob, task),
    jobs,
    context,
    assemble,
    manifest,
    template: loaded,
    measure,
    review: reviewWith(options),
    ...(told.length > 0 || keptChecks.length > 0 ? { notes: lengthNotes } : {}),
    layout,
    maxRepairs: options.maxRepairs ?? 2,
    onRepair: (attempt, issues) => options.onProgress?.({ stage: 'repair', attempt, issues }),
  });

  const errors = draft.issues.filter((issue) => issue.severity === 'error');
  if (errors.length > 0) throw new DeckValidationError(errors);
  // A chart's values that were not known stay so where the request gives no number for them.
  const statedNow = statedNumbers([instruction, deck.brief ?? '', deck.document ?? ''].join('\n'));
  for (const job of jobs) {
    const now = draft.deck.slides[job.number - 1];
    if (now && job.source) draft.deck.slides[job.number - 1] = keptGaps(job.source, now, statedNow);
  }
  if (doc) notes.push(...keptWarnings(doc, keptChecks, draft.deck));
  notes.push(...writtenUnder(draft.deck, told).map((under) => writtenShort(under, language)));
  // A slide the plan had words for and named no part of to write: its title was offered. Where
  // that came back as it was and the slide is set out as before, the deck is told that nothing
  // changed, and the slide is not among those that were rewritten.
  const same = new Set<number>();
  for (const job of jobs) {
    const now = draft.deck.slides[job.number - 1];
    if (!job.asItWas || !now || titleLine(now.slots.title) !== titleLine(job.asItWas.slots.title) || !sameArrange(now.arrange, job.asItWas.arrange)) continue;
    same.add(job.number);
    notes.push(nothingChanged(job.number));
  }
  // The numbers of a chart are checked against the brief where the chart was written now. A chart
  // that a free slide kept as it was is not looked at again: it is known by its content, since a
  // slide that lost a part has numbered the others anew.
  const keptCharts = new Map(jobs.map((job) => [job.number, new Set(job.free ? Object.values(job.carried ?? {}).filter(isChartValue).map((chart) => JSON.stringify(chart)) : [])]));
  const unstated = unstatedChartValues(manifest, draft.deck).filter((issue) => {
    const kept = issue.slide === undefined ? undefined : keptCharts.get(issue.slide);
    return kept !== undefined && !kept.has(JSON.stringify(draft.deck.slides[(issue.slide as number) - 1]?.slots[issue.slot ?? '']));
  });
  const rewritten = [...jobs.map((job) => job.number).filter((number) => !same.has(number)), ...reshaped].sort((a, b) => a - b);
  // The details of the slides written now, as their chart numbers.
  const details = unstatedDetails(draft.deck, new Set(rewritten));
  return result(draft.deck, rewritten, [...draft.issues.filter((issue) => !sent.has(issue)), ...unstated, ...details]);
}
