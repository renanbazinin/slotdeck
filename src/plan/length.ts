// An amount of text a person asks for on a slide ("200 words on the expansion"): where that text
// has the room, and what the writer is told. Without such a wish nothing here is used, and a plan
// and its writing are what they were.
//
// A slide's slots say how much to write, and their limits are hard. A wish for a length comes
// before what a slot says of length, but not before its limits: the text goes where it fits. That
// is the layout planned, another layout of plain text that holds it, a free slide set out for that
// length, or a free slide of that text alone set smaller where it needs it. A text asked for a
// slide stays on that slide: where nothing holds it, the slide holds what it can, and the plan and
// the deck say so.
import { toParagraphs } from '../deck/spec.ts';
import type { HeadingOver } from '../deck/headings.ts';
import type { DeckSpec, FreeArrange, FreeBlockKind } from '../deck/spec.ts';
import type { Issue } from '../deck/validate.ts';
import { freeSlotName } from '../free/blocks.ts';
import { said } from '../messages/en.ts';
import { scalesWithText, textBudget } from '../template/manifest.ts';
import type { TextSlotDef } from '../template/manifest.ts';
import { budgetScale, toldOf } from './catalog.ts';
import { mostHeld, toldFor } from '../template/told.ts';
import type { SlotCount } from '../template/told.ts';
import type { CatalogContext, PlannedSlot, UsablePattern } from './catalog.ts';
import type { TextShape } from './document.ts';

/**
 * Characters a word takes with the space after it, in running text: about 6 in English. Budgets
 * are in characters and a person counts in words; this is how the one becomes the other.
 */
export const WORD_CHARS = 6;
/**
 * The same in Hebrew: 6.2 to 6.6 in the texts the production model wrote for a slide (its words
 * carry their prefixes), so a text of 200 words written to 6 characters a word came out at 168.
 */
const HEBREW_WORD_CHARS = 6.5;
/** The most words a plan takes for one slide. More is a slip of a model, not a wish. */
export const MOST_WORDS = 3000;

/** Characters a word takes in a language, by its BCP 47 tag. */
export const wordChars = (language?: string): number => (/^(?:he|iw)(?:-|$)/i.test(language ?? '') ? HEBREW_WORD_CHARS : WORD_CHARS);
/** How many characters so many words take, in a language. */
export const charsOfWords = (words: number, language?: string): number => Math.round(words * wordChars(language));
/** About how many words so many characters hold, in a language. */
export const wordsOfChars = (chars: number, language?: string): number => Math.floor(chars / wordChars(language));

/**
 * Reads the words asked for a slide, as an answer or a stored plan gives them: a number from 1 to
 * {@link MOST_WORDS}, rounded to a whole one, or a string of digits that is one. Anything else is no
 * wish: a value over the most is a slip, not a wish to be held to the most, and `"1e9"` or `"0x20"`
 * is no number a person wrote.
 */
export function readWords(value: unknown): number | undefined {
  const words = typeof value === 'number' ? value : typeof value === 'string' && /^\s*\d+\s*$/.test(value) ? Number(value) : Number.NaN;
  if (!Number.isFinite(words) || words < 1 || words > MOST_WORDS) return undefined;
  return Math.round(words);
}

/**
 * The rule a plan is given: when a slide has words, and what they ask of its pattern. `keep`
 * says the slides listed may have words already (a plan being changed), which stay.
 */
export const wordsRule = (asked: 'brief' | 'request', keep = false): string =>
  `When the ${asked} asks for an amount of text on a slide ("200 words", "a paragraph of about 80 words"), set "words" on that slide to that number of words, ` +
  `and choose a pattern whose slots hold that much: a word takes about ${WORD_CHARS} characters with its space. Leave "words" out of every other slide${keep ? ', and keep the words a slide has unless the request changes them' : ''}.`;

/**
 * The rule a change to a plan that has words is given, after {@link wordsRule}: a slide given
 * words holds them, and code adds no slide for a text. The plan's lines mark a slide set small
 * ("text at 7 pt") and the most a dense slide holds ("at most about 570 words").
 */
export const WORDS_CHANGE_RULE =
  'A slide given words holds them all: the software sets its text smaller where they need it ("text at 7 pt"), so do not add slides to make room for a text. ' +
  'A slide marked "at most about N words" holds no more than N: to make it longer or shorter when the request names no number, set its "words" between the words it has and N. ' +
  'Give a text more slides only when the request asks for more slides or for larger text, and then share its words over them, each slide with "words" for its part. ' +
  'Where the request asks for fewer slides, a slide that takes the text of others takes their words too.';

/**
 * The rule a change to a deck that has a slide set dense is given, after the rule on words: such
 * a slide keeps its text whole when it is written again, and grows within the most it holds.
 */
export const DENSE_EDIT_RULE =
  'A slide marked "at most about N words" carries a text that was asked for, set smaller where it needs it. Keep that text whole when the slide is written again: leave "words" out to keep its length. ' +
  'To make it longer or shorter when the request names no number, set "words" between the words it has and N: it holds no more than N. ' +
  'Only when the request asks for more slides or for larger text, give its text a new slide after it, and set "words" on both for their parts.';

/** What the `words` of a slide say, in a plan or a change to one. */
export const WORDS_NOTE = 'Only when an amount of text is asked for on this slide: the number of words. Leave it out otherwise.';

/**
 * Whether a slot holds running text: paragraphs or a list, or one text of 200 characters or more.
 * A title, a heading, a caption or a figure does not, and the length asked for a slide is not
 * theirs. `names`, where given, are the only slots that are.
 */
function isRunning(planned: PlannedSlot, names?: ReadonlySet<string>): planned is PlannedSlot & { slot: TextSlotDef } {
  if (planned.slot.kind !== 'text') return false;
  if (names) return names.has(planned.slot.name);
  return planned.slot.name !== 'title' && (planned.mode === 'paragraphs' || (planned.slot.maxChars ?? 0) >= 200);
}

/** How many characters of running text a pattern holds at the deck's text size, as its writers are told (`toldOf`): in all its running slots, or in those named. */
export function textRoom(pattern: UsablePattern, context: CatalogContext, names?: ReadonlySet<string>): number {
  return pattern.slots.filter((planned) => isRunning(planned, names)).reduce((sum, planned) => sum + (toldOf(planned, context).maxChars ?? 0), 0);
}

/** The count of the template's spacing for a slot (`CatalogContext.told`), where it speaks for it. */
const countOf = (planned: PlannedSlot): SlotCount | undefined => (planned as { told?: SlotCount }).told;

/** Whether the count of the template's spacing speaks for a slot: what it is told is safe in any text of its paragraphs, and no share of it is kept back. */
const counted = (planned: PlannedSlot): boolean => countOf(planned) !== undefined;

/**
 * The running slots' room in all: of a slot the count speaks for, what `held` says it holds; of any
 * other, `share` of its stored budget.
 */
function sharedRoom(pattern: UsablePattern, context: CatalogContext, share: number, held: (planned: PlannedSlot, count: SlotCount) => number): number {
  let told = 0;
  let stored = 0;
  for (const planned of pattern.slots.filter((one) => isRunning(one))) {
    const count = countOf(planned);
    if (count) told += held(planned, count);
    else stored += toldOf(planned, context).maxChars ?? 0;
  }
  return Math.floor(told + share * stored);
}

/**
 * The texts of a free slide that a length asked for goes to, with each one's share in
 * characters: its running texts, or its lists where it has no running text. A slide of neither
 * has no place for it (see {@link withText}).
 */
export function freeLengths(blocks: readonly FreeBlockKind[], chars: number): Record<string, number> {
  const texts = blocks.flatMap((kind, at) => (kind === 'text' ? [freeSlotName(at + 1, 'Text')] : []));
  const lists = blocks.flatMap((kind, at) => (kind === 'list' ? [freeSlotName(at + 1, 'Items')] : []));
  const names = texts.length > 0 ? texts : lists;
  return Object.fromEntries(names.map((name, at) => [name, Math.floor(chars / names.length) + (at < chars % names.length ? 1 : 0)]));
}

/** The parts of a free slide with a running text among them: a slide of no text or list gets one, first. */
export const withText = (blocks: readonly FreeBlockKind[]): FreeBlockKind[] => (blocks.some((kind) => kind === 'text' || kind === 'list') ? [...blocks] : ['text', ...blocks]);

/**
 * What a description says of length, as a clause of its own: "About 420 characters in all",
 * "Write 260 to 500 characters", "Between 100 and 200 characters". Taken out where a length was
 * asked for, which comes first, when it opens the description before a colon or a full stop, ends
 * it after one, or is all of it. A length said inside a sentence ("sentences of 15 to 25 words
 * each", "50 to 70 characters, or more when there are fewer") is part of what the sentence says,
 * and stays: taken out, it would leave a sentence that says something else.
 */
const AMOUNT = String.raw`(?:write\s+)?(?:(?:about|around|roughly|up to|at most)\s+)?(?:between\s+\d[\d,]*\s+and\s+\d[\d,]*|\d[\d,]*(?:\s*(?:to|-|\u2013)\s*\d[\d,]*)?)\s+(?:characters?|words?)(?:\s+(?:in all|in total|altogether|each))?`;
const LEADING = new RegExp(String.raw`^\s*${AMOUNT}\s*(?:[:.]\s*|$)`, 'i');
const TRAILING = new RegExp(String.raw`(?:^|[.;]\s*)${AMOUNT}\s*\.?\s*$`, 'i');

/** A description without what it says of length as a clause of its own: "Four or five paragraphs of full sentences". Empty when that was all it said. */
export function withoutLength(description: string | undefined): string {
  const rest = (description ?? '').replace(LEADING, '').replace(TRAILING, '').trim();
  return rest.charAt(0).toUpperCase() + rest.slice(1);
}

/** A number of characters in round tens, as a writer is told it: rounded down, so that it is never over the limit stated after it. */
const tens = (chars: number): number => Math.max(10, Math.floor(chars / 10) * 10);
/** A number of words in round tens, rounded down: "at most about 570 words" is never more than a slide holds. */
export const tensDown = (words: number): number => Math.floor(words / 10) * 10;

/**
 * A running text longer than this, in characters, that has nothing else to say of itself is asked
 * for in paragraphs: written to a length alone, a model gives one block of a thousand characters.
 */
const IN_PARAGRAPHS = 500;

/**
 * The share of a layout's limits that an asked length may fill. A layout's limits were found once,
 * for the template's own text; a text that fills one to its limit in many paragraphs is the one most
 * likely to run over its box, where the template keeps space between paragraphs in points (532
 * characters in five paragraphs ran over the 556 of a text layout of the right-to-left template,
 * in PowerPoint). A free slide is set out for the text itself, counts that space, and has no such share.
 */
const LAYOUT_SHARE = 0.9;

/**
 * The share of a layout's limits that a text of three or more paragraphs may fill: a layout's
 * budget leaves out the space the template keeps between paragraphs, and many paragraphs take
 * more of it.
 */
const PARAGRAPH_SHARE = 0.75;

/**
 * A slot of paragraphs of a layout that takes this many or more takes one fewer of a document's
 * text of three or more paragraphs (see {@link documentParagraphs}). Every paragraph ends on a line
 * of its own: the 417 characters a text layout of the right-to-left template holds, in five
 * paragraphs each just over a line, took nine lines and ran 19 pt past its box in PowerPoint. In
 * four they take at most eight, which the box holds.
 */
const FEWER_FROM = 4;

/**
 * How many paragraphs of a document's text of `paragraphs` paragraphs a slot of paragraphs of a
 * layout takes, by its limit at the deck's text size (`textBudget`): one fewer than its limit where
 * that is `FEWER_FROM` or more and the text has three or more. One rule for what the writer is
 * told ({@link withPassages}), what code sets as it is (`filledText`), and where a text is held
 * (`homeForChars`). `undefined` where the slot has no limit.
 */
export function documentParagraphs(planned: PlannedSlot, context: CatalogContext, paragraphs: number): number | undefined {
  // A slot the count speaks for takes what the count tells a text of that many paragraphs.
  if (counted(planned)) return toldOf(planned, context, paragraphs).maxParagraphs;
  const most = textBudget(planned.slot as TextSlotDef, budgetScale(context)).maxParagraphs;
  if (most === undefined) return undefined;
  return paragraphs >= 3 && most >= FEWER_FROM ? most - 1 : most;
}

/** Which texts of a pattern take an asked length, and in which language it is counted. */
export interface LengthOptions {
  /** The texts of a free slide that were set out for the length (see {@link freeLengths}). Left out, the pattern is a layout's, and its running texts take it. */
  names?: ReadonlySet<string> | undefined;
  /** The deck's language, for the characters a word takes. */
  language?: string | undefined;
}

/**
 * How many characters of running text a pattern holds for an asked length: of a layout, the share
 * it may fill, and of a slot the count speaks for, the most it holds in the paragraphs its writer
 * may choose ({@link toldFor}).
 */
export function heldBy(pattern: UsablePattern, context: CatalogContext, options: LengthOptions = {}): number {
  if (options.names) return Math.floor(textRoom(pattern, context, options.names));
  return sharedRoom(pattern, context, LAYOUT_SHARE, (_planned, count) => mostHeld(count));
}

/**
 * How many characters of a document's text a layout holds before its paragraphs are known: as
 * {@link heldBy}, but of a slot the count speaks for, the pair it is told. A document's text keeps
 * its paragraphs, so what a slot holds in fewer of them is not its room. What a deck's slides are
 * sized by, where its document is shared out.
 */
export function documentHeld(pattern: UsablePattern, context: CatalogContext): number {
  return sharedRoom(pattern, context, LAYOUT_SHARE, (planned) => toldOf(planned, context).maxChars ?? 0);
}

/**
 * How many paragraphs of a document's text of `paragraphs` paragraphs the running slots of a
 * pattern take in all (see {@link documentParagraphs}): a slot of one line takes one, and a slot
 * with no limit any number.
 */
function slotParagraphs(pattern: UsablePattern, context: CatalogContext, paragraphs: number): number {
  return pattern.slots
    .filter((planned) => isRunning(planned))
    .reduce((sum, planned) => sum + (planned.mode === 'text' ? 1 : (documentParagraphs(planned, context, paragraphs) ?? Number.POSITIVE_INFINITY)), 0);
}

/**
 * How many characters of a text of so many paragraphs a pattern holds: of a free slide, what its
 * texts named were set out for; of a layout, the share an asked length may fill, three quarters
 * for a text of three or more paragraphs. A text with more paragraphs than its slots take joins
 * some of them, and loses no characters by that: a slot's limits are its characters and its
 * paragraphs, and none is a limit on one paragraph. So this is the room the writer is told (see
 * {@link withPassages}), and the room a slide is said to hold.
 */
export function documentRoom(pattern: UsablePattern, context: CatalogContext, shape?: { paragraphs: number }, names?: ReadonlySet<string>, asIs?: { chars: number; each: readonly number[] }): number {
  if (names) return heldBy(pattern, context, { names });
  if (!shape) return heldBy(pattern, context);
  const room = sharedRoom(pattern, context, shape.paragraphs >= 3 ? PARAGRAPH_SHARE : LAYOUT_SHARE, (planned) => toldOf(planned, context, shape.paragraphs).maxChars ?? 0);
  // A text code sets as it is (`TextShape.asIs`) in the one running slot is held where its own
  // paragraphs fit, each counted in whole lines: the same count, of lengths that are known.
  return asIs && holdsAsIs(pattern, asIs) ? Math.max(room, asIs.chars) : room;
}

/**
 * Whether a pattern holds a document's text only where code sets it as it is: its writer would be
 * told less than the text (`documentRoom` without `asIs`), and its own paragraphs fit
 * ({@link holdsAsIs}). Only where the count of the template's spacing speaks for its slot.
 */
export function holdsOnlyAsIs(pattern: UsablePattern, context: CatalogContext, shape: { paragraphs: number }, need: number, asIs: { chars: number; each: readonly number[] }): boolean {
  return documentRoom(pattern, context, shape) < need && documentRoom(pattern, context, shape, undefined, asIs) >= need;
}

/** Whether the one running slot of a pattern, counted, holds a text of paragraphs of these lengths as it is. */
export function holdsAsIs(pattern: UsablePattern, asIs: { each: readonly number[] }, names?: ReadonlySet<string>): boolean {
  const running = pattern.slots.filter((planned) => isRunning(planned, names));
  const [only] = running;
  const count = running.length === 1 && only ? (only as { told?: { holds(each: readonly number[]): boolean } }).told : undefined;
  return count !== undefined && count.holds(asIs.each);
}

/** One running slot's part of an asked length: its share in characters, and the range it is told. */
interface LengthShare {
  planned: PlannedSlot;
  share: number;
  least: number;
  most: number;
  /** Of a slot the count speaks for, the pair it is told for its share ({@link toldFor}). */
  pair?: { maxChars: number; maxParagraphs: number };
}

/**
 * How an asked length is shared over the running slots of a pattern (or those named), and whether
 * they hold less than it (`short`). `undefined` where the pattern has no running slot with room.
 */
function sharesOf(pattern: UsablePattern, words: number, context: CatalogContext, options: LengthOptions): { shares: LengthShare[]; short: boolean } | undefined {
  const { names, language } = options;
  const fill = names ? 1 : LAYOUT_SHARE;
  const running = pattern.slots.filter((planned) => isRunning(planned, names));
  // A slot the count speaks for holds what it holds in the paragraphs its share takes (`toldFor`).
  const rooms = running.map((planned) => {
    const count = countOf(planned);
    return count ? mostHeld(count) : fill * (toldOf(planned, context).maxChars ?? 0);
  });
  const total = rooms.reduce((sum, room) => sum + room, 0);
  if (running.length === 0 || total <= 0) return undefined;
  const need = charsOfWords(words, language);
  const short = total < need;
  const shares = running.map((planned, at): LengthShare => {
    const share = Math.min(rooms[at] as number, Math.round((need * (rooms[at] as number)) / total));
    const count = countOf(planned);
    const pair = count ? toldFor(count, share) : undefined;
    const room = pair ? pair.maxChars : (rooms[at] as number);
    // From the share to a little over it, never over what the slot may hold. A slot that holds
    // less than its share is asked for what it holds, from nine tenths of it.
    const least = tens(short ? 0.9 * share : Math.min(share, 0.95 * room));
    const most = Math.max(least, Math.min(tens(room), tens(1.15 * share)));
    return { planned, share, least, most, ...(pair ? { pair } : {}) };
  });
  return { shares, short };
}

/**
 * The pattern a slide is written against when a length was asked for it. Each running slot (or
 * each one named) is told its share of the length, in words and as a range of characters, before
 * the rest of its description, from which what it said of length is taken out. A model writes
 * short of a length: told "about" it, four fifths (164 words of 200 asked, live); told a range
 * from nine tenths of it, the bottom of that range (168). So the range starts at the length. A
 * slot's share follows its room, of which a layout's may be filled to its share (`LAYOUT_SHARE`).
 * Its limits are not changed: where the slide holds less than was asked, each slot is told what
 * it holds, and that it is the most this slide holds of what was asked.
 */
export function withLength(pattern: UsablePattern, words: number, context: CatalogContext, options: LengthOptions = {}): UsablePattern {
  const { language } = options;
  const shared = sharesOf(pattern, words, context, options);
  if (!shared) return pattern;
  const { shares, short } = shared;
  const several = shares.length > 1;
  const told = new Map<string, string>();
  const pairs = new Map<string, { maxChars: number; maxParagraphs: number }>();
  shares.forEach(({ planned, share, least, most, pair }) => {
    if (pair) pairs.set(planned.slot.name, pair);
    const amount = `Write ${Math.max(1, wordsOfChars(share, language))} words`;
    const range = `${least} to ${most} characters`;
    const said = short
      ? `${amount}${several ? ' here' : ''}, ${range}: of the ${words} words asked for this slide, the most it holds`
      : several
        ? `${amount} here, ${range}, of the ${words} words asked for this slide`
        : `${amount}, as was asked: ${range} in all`;
    const rest = withoutLength(planned.slot.description);
    const paragraphs = !rest && planned.mode === 'paragraphs' && share > IN_PARAGRAPHS ? 'In paragraphs of two to four sentences, one string each' : '';
    told.set(planned.slot.name, [said, rest || paragraphs].filter(Boolean).join('. '));
  });
  return {
    ...pattern,
    slots: pattern.slots.map((planned) => {
      const description = told.get(planned.slot.name);
      const count = countOf(planned);
      const pair = pairs.get(planned.slot.name);
      // A slot the count speaks for is told the pair for its share: its limits and its number of items say it.
      const counting = count && pair ? { told: { ...count, ...pair } } : {};
      return description === undefined ? planned : ({ ...planned, ...counting, slot: { ...planned.slot, description } } as PlannedSlot);
    }),
  };
}

/** What a slide made from a document is told of its text: which passages, how many paragraphs, and any words asked for it. */
export interface PassageOptions extends LengthOptions {
  /** The passages, as they are named: "p4-p6". */
  label: string;
  /** How many paragraphs the text has. */
  shape?: { paragraphs: number } | undefined;
  /** The words asked for the slide, where it was asked to be shorter than its passages. */
  words?: number | undefined;
}

/**
 * The pattern a slide is written against when it carries passages of a document. Each running slot
 * (or each one named) is told its share of the text, in characters, before the rest of its
 * description: from nine tenths of the share to a little over it, never over what the slot holds.
 * Where the slide holds less than the text, it is told to shorten it to what it holds, in its own
 * words; where fewer words were asked for, it is told that many. A slot of paragraphs is told to
 * give one string for each paragraph of the text. Its limits are not changed, but for one: a slot
 * of paragraphs of a layout that takes four or more is told one fewer for a text of three or more
 * paragraphs ({@link documentParagraphs}), and its writer joins them.
 */
export function withPassages(pattern: UsablePattern, need: number, context: CatalogContext, options: PassageOptions): UsablePattern {
  const { names, language, label, shape, words } = options;
  const fill = names ? 1 : (shape?.paragraphs ?? 1) >= 3 ? PARAGRAPH_SHARE : LAYOUT_SHARE;
  const scale = budgetScale(context);
  const running = pattern.slots.filter((planned) => isRunning(planned, names));
  const paragraphs = names ? undefined : shape?.paragraphs;
  const rooms = running.map((planned) => (counted(planned) ? (toldOf(planned, context, paragraphs).maxChars ?? 0) : fill * (textBudget(planned.slot as TextSlotDef, scale).maxChars ?? 0)));
  const total = rooms.reduce((sum, room) => sum + room, 0);
  if (running.length === 0 || total <= 0) return pattern;
  const holds = documentRoom(pattern, context, shape, names) >= need;
  const several = running.length > 1;
  const told = new Map<string, string>();
  running.forEach((planned, at) => {
    const room = rooms[at] as number;
    const share = Math.min(room, Math.round((need * room) / total));
    const least = tens(0.9 * share);
    const most = holds ? Math.max(least, Math.min(tens(room), tens(1.05 * share))) : Math.max(least, tens(room));
    const text = several ? `This slot's share of the slide's text from the document (${label})` : `The slide's text from the document (${label})`;
    const said =
      words !== undefined
        ? `${several ? Math.max(1, wordsOfChars(share, language)) : words} words of the slide's text from the document (${label}), ${least} to ${most} characters: keep its words and leave out what matters least`
        : !holds
          ? `${text} shortened to ${least} to ${most} characters, the most this slide holds: keep its words and leave out what matters least`
          : several
            ? `${text}, in order and in its own words: ${least} to ${most} characters`
            : `${text}, in its own words: ${least} to ${most} characters in all`;
    const hint = planned.mode === 'paragraphs' ? 'One string for each paragraph of the text; join paragraphs where there are more than the slot takes' : '';
    told.set(planned.slot.name, [said, hint, withoutLength(planned.slot.description)].filter(Boolean).join('. '));
  });
  /** The paragraphs a slot of a layout takes of the text ({@link documentParagraphs}), as its limit is counted (`textBudget`), where that is fewer than its limit. */
  const fewer = (planned: PlannedSlot): { maxParagraphs: number } | undefined => {
    if (names || planned.mode !== 'paragraphs' || counted(planned)) return undefined;
    const slot = planned.slot as TextSlotDef;
    const most = textBudget(slot, scale).maxParagraphs;
    const taken = documentParagraphs(planned, context, shape?.paragraphs ?? 1);
    if (most === undefined || taken === undefined || taken >= most) return undefined;
    return { maxParagraphs: Math.ceil(taken * (scalesWithText(slot) ? scale : 1)) };
  };
  /** A slot the count speaks for is told the pair for the text's paragraphs: its writer writes the text, so it is held to the count. */
  const pairFor = (planned: PlannedSlot): { told: SlotCount } | undefined => {
    const count = countOf(planned);
    if (names || !count || planned.mode !== 'paragraphs') return undefined;
    return { told: { ...count, ...toldOf(planned, context, shape?.paragraphs ?? 1) } };
  };
  return {
    ...pattern,
    slots: pattern.slots.map((planned) => {
      const description = told.get(planned.slot.name);
      return description === undefined ? planned : ({ ...planned, ...pairFor(planned), slot: { ...planned.slot, description, ...fewer(planned) } } as PlannedSlot);
    }),
  };
}

/** A slot told the whole of the length asked for its slide ("Write 500 words, as was asked"), and the range of characters it was told. */
export interface ToldLength {
  /** The slide's number, from 1. */
  slide: number;
  /** The slot, by the name the writer wrote it by. */
  slot: string;
  /** The words asked for the slide. */
  words: number;
  least: number;
  most: number;
}

/**
 * The slot a pattern tells the whole of an asked length, as {@link withLength} tells it: the one
 * running slot (or the one named) of a slide that holds the length, with the range of characters
 * it is told. `undefined` where the length is shared over several slots, or the slide holds less.
 */
export function toldAsAsked(pattern: UsablePattern, words: number, context: CatalogContext, options: LengthOptions = {}): Omit<ToldLength, 'slide' | 'words'> | undefined {
  const shared = sharesOf(pattern, words, context, options);
  const only = shared && !shared.short && shared.shares.length === 1 ? shared.shares[0] : undefined;
  return only ? { slot: only.planned.slot.name, least: only.least, most: only.most } : undefined;
}

/**
 * The share of the least it was told that a text written to an asked length must have. A text
 * under it was written well short of what the person asked for: it is sent back once, and the
 * deck is told when it is still short.
 */
export const WRITTEN_SHARE = 0.85;

/** The slots of a deck that were told an asked length and were written well short of it, with the characters they have. */
export function writtenUnder(deck: Pick<DeckSpec, 'slides'>, told: readonly ToldLength[]): Array<ToldLength & { chars: number }> {
  return told.flatMap((one) => {
    const value = deck.slides[one.slide - 1]?.slots[one.slot];
    if (typeof value !== 'string' && !Array.isArray(value)) return [];
    const chars = toParagraphs(value).reduce((sum, paragraph) => sum + paragraph.length, 0);
    return chars < WRITTEN_SHARE * one.least ? [{ ...one, chars }] : [];
  });
}

/** Sent back to the writer for a slot written well short of the length asked for its slide. */
export const writtenLength = (under: ToldLength & { chars: number }): Issue => ({
  severity: 'warning',
  slide: under.slide,
  slot: under.slot,
  ...said('written-under-retry', { chars: under.chars, words: under.words, least: under.least, most: under.most }),
});

/** Said of a slide whose text is still well short of the length asked for it once the writer has been asked again. */
export const writtenShort = (under: ToldLength & { chars: number }, language?: string): Issue => ({
  severity: 'warning',
  slide: under.slide,
  ...said('written-under', { words: under.words, hasWords: wordsOfChars(under.chars, language) }),
});

/** The names of the running slots of a pattern, in its order: those that hold running text, or those named among them. */
export const runningNames = (pattern: UsablePattern, names?: ReadonlySet<string>): string[] => pattern.slots.filter((planned) => isRunning(planned, names)).map((planned) => planned.slot.name);

/** The size the running text of a pattern is set at, in points: the smallest of its running slots, or of those named. `undefined` where none says. */
export function runningSize(pattern: UsablePattern, names?: ReadonlySet<string>): number | undefined {
  const sizes = pattern.slots.filter((planned) => isRunning(planned, names)).flatMap((planned) => {
    const size = (planned.slot as TextSlotDef).fontSize ?? (planned.slot as TextSlotDef).size;
    return size === undefined ? [] : [size];
  });
  return sizes.length > 0 ? Math.min(...sizes) : undefined;
}

/**
 * Whether running text set at `size` is set small: under four fifths of the size it has on a free
 * slide of the template when nothing makes it smaller, in the half points sizes are set in (9 pt
 * where that size is 11.5). Text down to four fifths was set so before there were dense slides,
 * and is not remarked on.
 */
export const setSmall = (size: number, templateSize: number): boolean => size < Math.round(0.8 * templateSize * 2) / 2 - 0.01;

/**
 * Said of a slide that holds less text than was asked for it: what was asked, and what it holds.
 * `freeOff` adds at its end that free slides are not on offer (the option, the template's
 * manifest, or a list of patterns that does not name them), so that no slide's text is set
 * smaller. What comes before it is said as it always was.
 */
export const shortOfWords = (slide: number, words: number, room: number, language?: string, freeOff = false): Issue => ({
  severity: 'warning',
  slide,
  ...said('words-cut', { words, holdsWords: wordsOfChars(room, language), freeOff }),
});

/** The same of a slide whose text is set small to hold what was asked (`FreeArrange.dense`), and still holds less. */
export const shortOfWordsDense = (slide: number, words: number, room: number, size: number, language?: string): Issue => ({
  severity: 'warning',
  slide,
  ...said('words-cut-dense', { words, holdsWords: wordsOfChars(room, language), size }),
});

/**
 * Said of a slide whose photo gave way to a text set dense: the text it was asked for, or, where
 * no length was asked now (`asked` false), the text it has, which a photo added beside it left too
 * little room.
 */
export const photoLeftOut = (slide: number, words: number, asked = true): Issue => ({
  severity: 'warning',
  slide,
  ...said('photo-left-out', { words, asked }),
});

/** Where the text of a slide goes: a layout of the template, with its number of items, or a free slide of these parts. */
export type TextHome = { pattern: string; items?: number | undefined } | { blocks: FreeBlockKind[]; arrange?: FreeArrange | undefined };

const isFreeHome = (home: TextHome): home is { blocks: FreeBlockKind[]; arrange?: FreeArrange | undefined } => 'blocks' in home;

/** What the planner has to find a home for a text with. */
export interface LengthPlanning {
  context: CatalogContext;
  /** The deck's language, for the characters a word takes. */
  language?: string | undefined;
  /** The pattern of a layout with so many items, as the writer gets it; `undefined` where it cannot be used. */
  layout: (id: string, items?: number) => UsablePattern | undefined;
  /** Whether a layout is plain text: a title with paragraphs or a list, and nothing else. */
  plainText: (id: string) => boolean;
  /** The layouts of plain text on offer, which a slide of plain text may move to. */
  plain: readonly string[];
  /**
   * The pattern of a free slide of these parts with these lengths, and where it is known the
   * number of paragraphs of some of them, where free slides are on offer; `undefined` where they
   * are not, or cannot be made.
   */
  free?: FreeMaker | undefined;
  /**
   * The same for a free slide a plan has, which is written whatever is on offer for a new slide:
   * what it holds where free slides are off (see {@link TextFirst}). Left out, as `free`.
   */
  written?: FreeMaker | undefined;
  /** The parts a free slide is made of when a slide of a layout becomes one. */
  freeParts: (id: string) => FreeBlockKind[];
}

/**
 * The pattern of a free slide of these parts with these lengths, and where it is known the number
 * of paragraphs of some of them, and of a slide set dense the lengths of the paragraphs of a text
 * code sets as it is (`asIs`).
 */
export type FreeMaker = (
  blocks: FreeBlockKind[],
  arrange: FreeArrange | undefined,
  lengths: Record<string, number>,
  paragraphs?: Record<string, number>,
  asIs?: Record<string, readonly number[]>,
  heads?: Record<string, readonly HeadingOver[]>,
) => UsablePattern | undefined;

/** The paragraphs of a text, where they are known: a document's text has those of its passages (see `shapeOf`). */
type Shape = TextShape | undefined;

/**
 * How many characters of running text a home holds when it is set out for so many: of a layout,
 * the share an asked length may fill, or for a text of known paragraphs the share it may fill of
 * them (see {@link documentRoom}). A free slide whose text goes to one slot is set out for that
 * many paragraphs.
 */
function roomOf(home: TextHome, chars: number, planning: LengthPlanning, shape?: Shape): number | undefined {
  if (!isFreeHome(home)) {
    const pattern = planning.layout(home.pattern, home.items);
    return pattern ? (shape ? documentRoom(pattern, planning.context, shape, undefined, setOnLayout(shape, chars)) : heldBy(pattern, planning.context)) : undefined;
  }
  const lengths = freeLengths(home.blocks, chars);
  const names = Object.keys(lengths);
  const [name] = names;
  // A document's text that code sets as it is (`TextShape.asIs`) is judged as it will be set: only
  // where these are its characters, not where words asked for the slide make it shorter.
  const asIs = shape?.asIs && shape.asIs.chars === chars && name !== undefined ? { [name]: shape.asIs.each } : undefined;
  const heads = asIs && shape?.asIs?.heads && name !== undefined ? { [name]: shape.asIs.heads } : undefined;
  const pattern = shape && name !== undefined && names.length === 1 ? planning.free?.(home.blocks, home.arrange, lengths, { [name]: shape.paragraphs }, asIs, heads) : planning.free?.(home.blocks, home.arrange, lengths);
  return pattern ? heldBy(pattern, planning.context, { names: new Set(names) }) : undefined;
}

/** A free home set dense (`FreeArrange.dense`). */
export const isDenseHome = (home: TextHome): boolean => isFreeHome(home) && home.arrange?.dense === true;
/** A free home of one running text and nothing else. */
const isOneText = (home: TextHome): boolean => isFreeHome(home) && home.blocks.length === 1 && home.blocks[0] === 'text';
/**
 * Whether a slide of these parts may become a text alone set dense: one running text, and photos,
 * which then give way. With `join`, a slide of several texts and photos may too: its texts become
 * one running text. That is for a plan, where nothing is written yet; a change to a deck keeps a
 * second text it has. A list, a chart, a table or any other part stays here: the slide keeps them,
 * and holds what it can of the text, unless its text comes first (`first` in {@link homeForChars}).
 */
export const densable = (blocks: readonly FreeBlockKind[], join: boolean): boolean => {
  const texts = blocks.filter((kind) => kind === 'text').length;
  return (texts === 1 || (join && texts > 1)) && blocks.every((kind) => kind === 'text' || kind === 'photo');
};

/** The same home at the template's own sizes: without `dense`, its arrangement left out when that was all it said. */
export function undense(home: TextHome): TextHome {
  if (!isFreeHome(home) || !isDenseHome(home)) return home;
  const { dense: _dense, ...rest } = home.arrange as FreeArrange;
  return Object.keys(rest).length > 0 ? { blocks: home.blocks, arrange: rest } : { blocks: home.blocks };
}

/** Whether two homes are one: the same layout with the same items, or the same parts set out alike. */
export function sameHome(a: TextHome, b: TextHome): boolean {
  if (!isFreeHome(a) || !isFreeHome(b)) return !isFreeHome(a) && !isFreeHome(b) && a.pattern === b.pattern && a.items === b.items;
  const set = (arrange: FreeArrange | undefined) => JSON.stringify([arrange?.columns ?? null, arrange?.lead ?? null, arrange?.dense === true]);
  return a.blocks.length === b.blocks.length && a.blocks.every((kind, at) => kind === b.blocks[at]) && set(a.arrange) === set(b.arrange);
}

/** Where a text goes: its home, what that holds in characters, and whether it holds the text. */
export interface HomeFound {
  home: TextHome;
  room: number;
  holds: boolean;
}

/**
 * The text of a slide comes first (see {@link homeForChars}): where the home found keeps less than
 * `below` of the text (1: wherever it is cut), the text alone, set dense, is tried, and taken where
 * it keeps more. With `plain`, where free slides are not on offer, the layout of plain text that
 * keeps most of it is tried the same way: a document's text on a layout of cards or figures that
 * holds a few characters of it, or on a free slide of a text and other parts that a plan made with
 * free slides on (a free slide of its text alone stays). There a layout of a table keeps the rows
 * of a table of the passages that it shows, as well as its running text. `need` and `shape` are
 * those of the text alone, where a table of a document's passages is running text.
 */
export interface TextFirst {
  need: number;
  shape?: TextShape | undefined;
  below: number;
  plain?: boolean;
}

/** What a home keeps of a text: the share of it that its room holds. */
const keptOf = (room: number, need: number): number => (need > 0 ? Math.min(room, need) / need : 1);

/**
 * Where a text of so many characters goes, in this order: the home it has, at the template's own
 * sizes; the smallest layout of plain text that holds it, for a slide of plain text; a free slide
 * set out for that length; and, where `dense` allows it and the slide shows one running text and
 * nothing else but photos, a free slide of that text alone set dense. With `join`, a slide of
 * several texts and photos may become that text alone too, last: a plan's slide of seven short
 * texts for a long text is one text of it. A text of known paragraphs (`shape`) that a layout's
 * slots take fewer of ({@link documentParagraphs}) goes on to a layout of plain text that takes
 * them, or to a free slide, where code sets them as they are; where none holds it, a layout that
 * holds its characters does, and its writer joins paragraphs. Where none holds it, the one that
 * holds most. With `first`, where that keeps more of the
 * text, the text alone set dense, whatever else the slide shows (cards, a chart, a table, a list):
 * those give way, as photos do; or, where free slides are not on offer, a layout of plain text
 * (`TextFirst.plain`). With `dense`, a free home of one text is returned set dense, also where it
 * holds the text at the template's size: its text is then set at that size, counted as PowerPoint
 * sets it, and smaller only where the paragraphs written need it. A home that comes out the same as the one given is
 * the one given, so that a caller can tell by identity that nothing moved.
 */
export function homeForChars(given: TextHome, need: number, planning: LengthPlanning, options: { shape?: TextShape | undefined; dense?: boolean; join?: boolean; first?: TextFirst | undefined } = {}): HomeFound {
  const dense = options.dense === true;
  const found = place(undense(given), need, planning, dense, options.shape, options.join === true);
  let result: HomeFound;
  if (!dense || !isOneText(found.home) || isDenseHome(found.home)) result = sameHome(found.home, given) ? { ...found, home: given } : found;
  else {
    const home: TextHome = { blocks: ['text'], arrange: { ...((found.home as { arrange?: FreeArrange }).arrange ?? {}), dense: true } };
    result = { ...found, home: sameHome(home, given) ? given : home };
  }
  const { first } = options;
  if (!first || result.holds || keptOf(result.room, need) >= first.below) return result;
  let other: HomeFound | undefined;
  // What the home found keeps of the text: of its running text, and, without free slides, the rows
  // of a table of the passages that it shows in its table too, which the other home has as text.
  let kept = keptOf(result.room, need);
  if (planning.free) {
    if (!dense || (isOneText(result.home) && isDenseHome(result.home))) return result;
    // The text alone, set dense: what else the slide shows gives way to it.
    const own = isOneText(given) ? (given as { arrange?: FreeArrange }).arrange : undefined;
    const alone: TextHome = { blocks: ['text'], arrange: { ...(own ?? {}), dense: true } };
    const room = roomOf(alone, first.need, planning, first.shape);
    if (room !== undefined) other = { home: alone, room, holds: room >= first.need };
  } else if (first.plain && !isOneText(result.home)) {
    // Without free slides, the layout of plain text that holds the text, the smallest, else the one
    // that holds most of it: better than a layout that holds a few characters of it, or than a free
    // slide a plan made with free slides on, of a text and cards, that holds next to none. One where
    // code sets a document's text as it is comes first (`setsAsIs`).
    const rooms = planning.plain
      .map((id) => ({ id, room: roomOf({ pattern: id }, first.need, planning, first.shape) ?? 0 }))
      .filter((one) => one.room > 0)
      .sort((a, b) => a.room - b.room);
    const pick = rooms.find((one) => one.room >= first.need && setsAsIs({ pattern: one.id }, planning, first.shape, first.need)) ?? rooms.find((one) => one.room >= first.need) ?? rooms.at(-1);
    if (pick) other = { home: { pattern: pick.id }, room: pick.room, holds: pick.room >= first.need };
    // A free slide the plan has is written all the same: it keeps what it holds as it is written.
    const room = isFreeHome(result.home) ? (roomOf(result.home, need, { ...planning, free: planning.written }, options.shape) ?? result.room) : result.room;
    kept = keptOf(room, need);
    if (first.need > need) kept = (Math.min(room, need) + first.need - need) / first.need;
  }
  if (!other || keptOf(other.room, first.need) <= kept) return result;
  return { ...other, home: sameHome(other.home, given) ? given : other.home };
}

/**
 * Where a text of so many words asked for a slide of a plan goes (see {@link homeForChars}), a free
 * text alone set dense where it needs it. A slide that keeps cards, a chart, a table or a list
 * beside the text and so holds less than half of it is the text alone, where that holds more: a
 * chart slide that holds most of its words keeps its chart. With `first` false its parts stay,
 * whatever they leave the text: parts a change to a plan asks for, or asked for before.
 */
export const homeFor = (current: TextHome, words: number, planning: LengthPlanning, first = true): HomeFound => {
  const need = charsOfWords(words, planning.language);
  return homeForChars(current, need, planning, { dense: true, join: true, ...(first ? { first: { need, below: 0.5 } } : {}) });
};

function place(current: TextHome, need: number, planning: LengthPlanning, dense: boolean, shape: Shape, join: boolean): HomeFound {
  const tried: Array<{ home: TextHome; room: number }> = [];
  /**
   * Whether a home with this room holds the text as it is: a layout whose slots take fewer
   * paragraphs than the text has ({@link documentParagraphs}) does not. The text goes on to a
   * layout of plain text that takes them, or to a free slide, where code sets them as they are.
   * Where none does, such a layout still holds it (see the end of `place`): its writer joins
   * paragraphs and loses no characters by that, so its room stays what it holds (see
   * {@link documentRoom}).
   */
  const holdsIn = (home: TextHome, room: number): boolean => room >= need && (need === 0 || takesParagraphs(home, planning, shape, need));
  const attempt = (home: TextHome): boolean => {
    const room = roomOf(home, need, planning, shape);
    if (room === undefined) return false;
    tried.push({ home, room });
    return holdsIn(home, room);
  };
  const last = (): number => (tried.at(-1) as { room: number }).room;
  if (attempt(current)) return { home: current, room: last(), holds: true };
  if (!isFreeHome(current) && planning.plainText(current.pattern)) {
    const others = planning.plain
      .filter((id) => id !== current.pattern)
      .map((id) => ({ id, room: roomOf({ pattern: id }, need, planning, shape) ?? 0 }))
      .filter((other) => other.room > 0)
      .sort((a, b) => a.room - b.room);
    for (const other of others) tried.push({ home: { pattern: other.id }, room: other.room });
    // One where code sets a document's text as it is comes first.
    const fits = others.find((other) => holdsIn({ pattern: other.id }, other.room) && (need === 0 || setsAsIs({ pattern: other.id }, planning, shape, need))) ?? others.find((other) => holdsIn({ pattern: other.id }, other.room));
    if (fits) return { home: { pattern: fits.id }, room: fits.room, holds: true };
  }
  if (planning.free) {
    // A free slide that has a running text was tried as it is. One that has none gets one, and a
    // slide of a layout becomes a free slide of a running text and what else it shows.
    const blocks = withText(isFreeHome(current) ? current.blocks : planning.freeParts(current.pattern));
    const asIs = isFreeHome(current) && blocks.length === current.blocks.length;
    if (!asIs && attempt({ blocks })) return { home: { blocks }, room: last(), holds: true };
    // Its text alone, set dense: smaller, down to the least size that reads, in the most room the
    // frame has. One that does not hold the text there is still the home that holds most of it.
    if (dense && densable(blocks, join)) {
      const own = isOneText(current) ? (current as { arrange?: FreeArrange }).arrange : undefined;
      const home: TextHome = { blocks: ['text'], arrange: { ...(own ?? {}), dense: true } };
      if (attempt(home)) return { home, room: last(), holds: true };
    }
  }
  // A layout that holds the characters of a text, though not its paragraphs as they are, holds it
  // where nothing else does: its writer joins some of them. The one that takes most of them, the
  // home it has before another, else the least.
  const taken = (home: TextHome): number => {
    const pattern = shape && !isFreeHome(home) ? planning.layout(home.pattern, home.items) : undefined;
    const asIs = setOnLayout(shape, need);
    if (pattern && asIs && holdsAsIs(pattern, asIs)) return shape?.paragraphs ?? Number.POSITIVE_INFINITY;
    return pattern && shape ? Math.min(shape.paragraphs, slotParagraphs(pattern, planning.context, shape.paragraphs)) : Number.POSITIVE_INFINITY;
  };
  const loose = tried
    .filter((entry) => entry.room >= need)
    .map((entry) => ({ entry, taken: taken(entry.home), own: entry.home === current ? 0 : 1 }))
    .sort((a, b) => b.taken - a.taken || a.own - b.own || a.entry.room - b.entry.room);
  const kept = loose[0]?.entry;
  if (kept) return { ...kept, holds: true };
  const best = tried.reduce<{ home: TextHome; room: number } | undefined>((most, entry) => (!most || entry.room > most.room ? entry : most), undefined);
  return best ? { ...best, holds: false } : { home: current, room: 0, holds: false };
}

/**
 * The lengths of a text's own paragraphs that a layout is judged by, where code sets the text there
 * as it is (`TextShape.asIs`): a text of `need` characters, that code sets on a layout
 * (`TextShape.layoutAsIs`).
 */
function setOnLayout(shape: Shape, need: number | undefined): { chars: number; each: readonly number[] } | undefined {
  return shape?.asIs && shape.layoutAsIs !== false && shape.asIs.chars === need ? shape.asIs : undefined;
}

/** Whether the running slots of a home take a text's paragraphs as they are ({@link documentParagraphs}): a free slide's, and any where they are not known, do. */
function takesParagraphs(home: TextHome, planning: LengthPlanning, shape: Shape, need?: number): boolean {
  if (!shape || isFreeHome(home)) return true;
  const pattern = planning.layout(home.pattern, home.items);
  // A text code sets as it is, that the one running slot holds as it is, keeps its paragraphs.
  const asIs = setOnLayout(shape, need);
  if (pattern && asIs && holdsAsIs(pattern, asIs)) return true;
  return !pattern || shape.paragraphs <= slotParagraphs(pattern, planning.context, shape.paragraphs);
}

/**
 * Whether code can set a document's text as it is in a home: a free slide, or a layout of one
 * running slot that takes its paragraphs (see `filledText`). Of two layouts of plain text that hold
 * a text, one of two columns that the writer fills each, and one of a single text that takes its
 * paragraphs, the second is chosen: there the person's words are set as they are.
 */
function setsAsIs(home: TextHome, planning: LengthPlanning, shape: Shape, need?: number): boolean {
  if (!shape || isFreeHome(home)) return true;
  const pattern = planning.layout(home.pattern, home.items);
  if (!pattern) return true;
  return pattern.slots.filter((planned) => isRunning(planned)).length === 1 && takesParagraphs(home, planning, shape, need);
}

/** The purpose of one of several slides that share a text: "The plan for 2026 (2/3)". */
export const partPurpose = (purpose: string, part: number, parts: number): string => (parts > 1 ? `${purpose} (${part}/${parts})` : purpose);
