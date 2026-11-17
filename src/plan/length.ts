// An amount of text a person asks for on a slide ("200 words on the expansion"): where that text
// has the room, and what the writer is told. Without such a wish nothing here is used, and a plan
// and its writing are what they were.
//
// A slide's slots say how much to write, and their limits are hard. A wish for a length comes
// before what a slot says of length, but not before its limits: the text goes where it fits. That
// is the layout planned, another layout of plain text that holds it, or a free slide set out for
// that length. Where none holds it, the text goes over more slides when the deck may have them,
// and otherwise the slide holds what it can, and the deck is told.
import type { FreeArrange, FreeBlockKind } from '../deck/spec.ts';
import type { Issue } from '../deck/validate.ts';
import { freeSlotName } from '../free/blocks.ts';
import { textBudget } from '../template/manifest.ts';
import type { TextSlotDef } from '../template/manifest.ts';
import { budgetScale } from './catalog.ts';
import type { CatalogContext, PlannedSlot, UsablePattern } from './catalog.ts';

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

/** How many characters of running text a pattern holds at the deck's text size: in all its running slots, or in those named. */
export function textRoom(pattern: UsablePattern, context: CatalogContext, names?: ReadonlySet<string>): number {
  const scale = budgetScale(context);
  return pattern.slots.filter((planned) => isRunning(planned, names)).reduce((sum, planned) => sum + (textBudget(planned.slot as TextSlotDef, scale).maxChars ?? 0), 0);
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

/** Which texts of a pattern take an asked length, and in which language it is counted. */
export interface LengthOptions {
  /** The texts of a free slide that were set out for the length (see {@link freeLengths}). Left out, the pattern is a layout's, and its running texts take it. */
  names?: ReadonlySet<string> | undefined;
  /** The deck's language, for the characters a word takes. */
  language?: string | undefined;
}

/** How many characters of running text a pattern holds for an asked length: of a layout, the share it may fill. */
export function heldBy(pattern: UsablePattern, context: CatalogContext, options: LengthOptions = {}): number {
  return Math.floor((options.names ? 1 : LAYOUT_SHARE) * textRoom(pattern, context, options.names));
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
  const { names, language } = options;
  const fill = names ? 1 : LAYOUT_SHARE;
  const scale = budgetScale(context);
  const running = pattern.slots.filter((planned) => isRunning(planned, names));
  const rooms = running.map((planned) => fill * (textBudget(planned.slot as TextSlotDef, scale).maxChars ?? 0));
  const total = rooms.reduce((sum, room) => sum + room, 0);
  if (running.length === 0 || total <= 0) return pattern;
  const need = charsOfWords(words, language);
  const several = running.length > 1;
  const short = total < need;
  const told = new Map<string, string>();
  running.forEach((planned, at) => {
    const room = rooms[at] as number;
    const share = Math.min(room, Math.round((need * room) / total));
    // From the share to a little over it, never over what the slot may hold. A slot that holds
    // less than its share is asked for what it holds, from nine tenths of it.
    const least = tens(short ? 0.9 * share : Math.min(share, 0.95 * room));
    const most = Math.max(least, Math.min(tens(room), tens(1.15 * share)));
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
      return description === undefined ? planned : ({ ...planned, slot: { ...planned.slot, description } } as PlannedSlot);
    }),
  };
}

/** Said of a slide that holds less text than was asked for it: what was asked, and what it holds. */
export const shortOfWords = (slide: number, words: number, room: number, language?: string): Issue => ({
  severity: 'warning',
  slide,
  message: `${words} words were asked for this slide, and it holds about ${wordsOfChars(room, language)}: ask for the text over more slides, or for a shorter one`,
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
  /** The pattern of a free slide of these parts with these lengths, where free slides are on offer; `undefined` where they are not, or cannot be made. */
  free?: ((blocks: FreeBlockKind[], arrange: FreeArrange | undefined, lengths: Record<string, number>) => UsablePattern | undefined) | undefined;
  /** The parts a free slide is made of when a slide of a layout becomes one. */
  freeParts: (id: string) => FreeBlockKind[];
}

/** How many characters of running text a home holds when it is set out for so many: of a layout, the share an asked length may fill. */
function roomOf(home: TextHome, chars: number, planning: LengthPlanning): number | undefined {
  if (!isFreeHome(home)) {
    const pattern = planning.layout(home.pattern, home.items);
    return pattern ? heldBy(pattern, planning.context) : undefined;
  }
  const lengths = freeLengths(home.blocks, chars);
  const pattern = planning.free?.(home.blocks, home.arrange, lengths);
  return pattern ? heldBy(pattern, planning.context, { names: new Set(Object.keys(lengths)) }) : undefined;
}

/**
 * Where a text of so many words goes: the home it has, when that holds it; else the smallest
 * layout of plain text that holds it, for a slide of plain text; else a free slide set out for
 * that length, where free slides are on offer. Where none holds it, the one that holds most, and
 * `holds` is false. `room` is what the home holds, in characters.
 */
export function homeFor(current: TextHome, words: number, planning: LengthPlanning): { home: TextHome; room: number; holds: boolean } {
  const need = charsOfWords(words, planning.language);
  const tried: Array<{ home: TextHome; room: number }> = [];
  const attempt = (home: TextHome): boolean => {
    const room = roomOf(home, need, planning);
    if (room === undefined) return false;
    tried.push({ home, room });
    return room >= need;
  };
  if (attempt(current)) return { home: current, room: (tried[0] as { room: number }).room, holds: true };
  if (!isFreeHome(current) && planning.plainText(current.pattern)) {
    const others = planning.plain
      .filter((id) => id !== current.pattern)
      .map((id) => ({ id, room: roomOf({ pattern: id }, need, planning) ?? 0 }))
      .filter((other) => other.room > 0)
      .sort((a, b) => a.room - b.room);
    for (const other of others) tried.push({ home: { pattern: other.id }, room: other.room });
    const fits = others.find((other) => other.room >= need);
    if (fits) return { home: { pattern: fits.id }, room: fits.room, holds: true };
  }
  if (planning.free) {
    // A free slide that has a running text was tried as it is. One that has none gets one, and a
    // slide of a layout becomes a free slide of a running text and what else it shows.
    const blocks = withText(isFreeHome(current) ? current.blocks : planning.freeParts(current.pattern));
    const asIs = isFreeHome(current) && blocks.length === current.blocks.length;
    if (!asIs && attempt({ blocks })) return { home: { blocks }, room: (tried.at(-1) as { room: number }).room, holds: true };
  }
  const best = tried.reduce<{ home: TextHome; room: number } | undefined>((most, entry) => (!most || entry.room > most.room ? entry : most), undefined);
  return best ? { ...best, holds: false } : { home: current, room: 0, holds: false };
}

/**
 * Whether a text may go over more slides of this home: one of plain text, a layout of plain text
 * or a free slide of running texts and lists alone. A slide of anything else (a chart, a table, a
 * photo, cards, figures) is not repeated for its text: every copy would ask for that content again,
 * which the brief has once.
 */
export const spreads = (home: TextHome, planning: LengthPlanning): boolean =>
  isFreeHome(home) ? home.blocks.length > 0 && home.blocks.every((kind) => kind === 'text' || kind === 'list') : planning.plainText(home.pattern);

/**
 * How many slides a text of so many words takes in a home that holds `held` words, when the deck
 * has `spare` slides to give: at least 1, and no more than one plus the spare ones. Counted in
 * whole words, so that the share of each slide is never over what it holds.
 */
export const partsFor = (words: number, held: number, spare: number): number => (held > 0 ? Math.max(1, Math.min(spare + 1, Math.ceil(words / held))) : 1);

/** The purpose of one of several slides that share a text: "The plan for 2026 (2/3)". */
export const partPurpose = (purpose: string, part: number, parts: number): string => (parts > 1 ? `${purpose} (${part}/${parts})` : purpose);
