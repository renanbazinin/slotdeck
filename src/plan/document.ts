// A document a deck is made from, word for word: a report, minutes, an article. Its text is cut
// into numbered passages, the outline puts each passage on a slide, and the slides carry the
// document's own text. A slide that holds its passages has them set into its running text by code;
// the writer writes the rest, and is held to the words of the text. Without a document nothing here
// is used, and a plan and its writing are what they were.
import { toParagraphs } from '../deck/spec.ts';
import { ITEM, leadsInto } from '../deck/headings.ts';
import type { DeckSpec, SlideSpec, SlotValue } from '../deck/spec.ts';
import type { Issue } from '../deck/validate.ts';
import type { JsonSchema } from '../llm/provider.ts';
import { said } from '../messages/en.ts';
import { charsOfWords, readWords } from './length.ts';

/** The most characters of a document a deck is made from: about 1,800 words, which one writing call carries. */
export const MOST_DOCUMENT_CHARS = 12000;
/** The most slides a deck made from a document has when no number of slides was asked for. */
export const MOST_DOCUMENT_SLIDES = 30;

/** A passage is cut where it is longer than this, in characters: about what a slide holds in comfort. */
const PASSAGE_CHARS = 500;
/** Lines of text shorter than this go together into one passage, up to {@link PASSAGE_CHARS}. */
const SHORT_PASSAGE = 120;
/** The most passages a document is cut into. More are joined, the shortest first. */
const MOST_PASSAGES = 150;
/** A line this long that ends in no stop goes on in the next line: a paragraph broken where a page was. */
const SOFT_WRAP = 60;
/** A heading has at most so many characters and words. */
const HEADING_MOST = 80;
const HEADING_WORDS = 12;
/** A text a person asked to be shorter: `words` under this share of the words of its passages. */
export const SHORTEN_SHARE = 0.8;
/** A slide written from its passages that keeps less than this share of their words is sent back once. */
export const KEPT_REPAIR = 0.85;
/** The check of kept words runs only on a slide whose room holds at least this share of its text. */
const KEPT_ROOM = 0.9;
/** Nor on a slide whose text has fewer words than this: a line or two may well be put another way. */
const KEPT_LEAST_TOKENS = 20;

/** One passage of a document, as the planner numbers them. */
export interface DocumentPassage {
  /** From 1, in the document's order: p1, p2. */
  number: number;
  /** A paragraph or a piece of one, a list, or a table. */
  kind: 'text' | 'list' | 'table';
  /** As the document has it: a list one item to a line, a table one row to a line. */
  text: string;
  /** The heading the passage stands under. Headings are not passages. */
  section?: string;
  /** Goes on with the passage before it: a long paragraph or list cut at a sentence or an item. */
  continues?: true;
  /** Ends with a colon, and so goes on a slide with the passage after it. */
  leads?: true;
  /** The characters a slide shows of it: without list markers, cells joined by " | ". */
  chars: number;
  /** The paragraphs a slide shows of it: its lines for text, its items for a list, its rows for a table. */
  paragraphs: number;
}

/** A document as the planning stages read it. */
export interface DocumentPlan {
  /** The document as it was given. */
  text: string;
  passages: DocumentPassage[];
  /** The characters of all its passages. */
  chars: number;
  /** The words of all its passages. */
  words: number;
  /**
   * Whether it is written in the deck's language: in the language the planner said it is written
   * in, or else in the script of the deck's language. Text in another is translated, and is never
   * set by code.
   */
  same: boolean;
}

// --- the text, cut into passages ---

/** Marks that set the direction of text: kept in the text, never counted. */
const BIDI = /[\u200E\u200F\u202A-\u202E\u2066-\u2069]/g;
/**
 * A line that is an item of a list: a bullet, or a number or a letter with a stop or a bracket. An
 * "o" or a "§" before a tab is a bullet too: a word processor writes the bullets of a list within a
 * list so, as plain text.
 */
/** The bullet of an item, where it has one rather than a number or a letter. */
const BULLET = /^\s*(?:[-*•‣◦▪–·]|[o§](?=\t))\s+/u;
/** The bullets a word processor writes as letters, as a slide shows them. */
const WORD_BULLETS: Readonly<Record<string, string>> = { o: '◦', '§': '▪' };
const MARKED = /^#{1,6}\s+/;
/** Closing quotes and brackets, which come after the stop that ends a line. */
const CLOSING = /["'”’»)\]\s]+$/u;
/** Where a sentence ends: after a stop and any closing quotes, before the space that follows. */
const SENTENCE_END = /(?<=[.!?…׃])["'”’)\]]*\s+/gu;
/** Words before a full stop that end no sentence. A single letter ends none either. */
const NO_END = new Set(['e.g', 'i.e', 'vs', 'cf', 'etc', 'mr', 'dr', 'st']);

/** The characters a text shows, without the marks that set its direction. */
const charsIn = (text: string): number => text.replace(BIDI, '').length;
/** The last character of a line before any closing quotes or brackets. */
const lastOf = (line: string): string => line.replace(CLOSING, '').at(-1) ?? '';

/** The working copy of a document: without a leading byte order mark, with one kind of line end, each line without the spaces at its end. */
function linesOf(document: string): string[] {
  return document
    .replace(/^\uFEFF/, '')
    .replace(/\r\n?/g, '\n')
    .split('\n')
    .map((line) => line.trimEnd());
}

/** The characters of a document as it is counted against {@link MOST_DOCUMENT_CHARS}. */
export const documentChars = (document: string): number => charsIn(linesOf(document).join('\n').trim());

type LineKind = 'row' | 'item' | 'heading' | 'text';

/**
 * A row of a table: a line with a tab between two of its cells, or with two bars or more. A tab
 * before a line's text only sets it in, as text pasted from a word processor has it.
 */
const isRow = (line: string): boolean => line.trim().includes('\t') || (line.match(/\|/g)?.length ?? 0) >= 2;
/** How many tabs a line has before its text. */
const leadOf = (line: string): number => (/^[ \t]*/.exec(line)?.[0].split('\t').length ?? 1) - 1;
/**
 * The cells of a row of a table, as a slide shows them. Of the tabs before its text, the first
 * `indent` set the row in; any more stand for empty cells, as the corner of a table pasted from a
 * spreadsheet has.
 */
function cellsOf(line: string, indent = 0): string[] {
  const tabs = line.trim().includes('\t');
  const cells = (tabs ? line.split('\t').slice(indent) : line.trim().split('|')).map((cell) => cell.trim());
  if (!tabs) {
    if (cells[0] === '') cells.shift();
    if (cells.at(-1) === '') cells.pop();
  }
  return cells;
}
/** A row of dashes under the head of a table written with bars. It shows nothing. */
const isRule = (line: string): boolean => cellsOf(line).every((cell) => /^:?-{2,}:?$/.test(cell));
/**
 * The rows of a table as a slide shows them: the cells joined by " | ". Tabs that set in every row
 * of the table only set it in.
 */
function rowsOf(lines: readonly string[]): string[] {
  const rows = lines.filter((line) => !isRule(line));
  const indent = Math.min(...rows.map(leadOf));
  return rows.map((line) => cellsOf(line, indent).join(' | '));
}
/** An item of a list without its marker. */
const itemText = (line: string): string => line.replace(ITEM, '').trim();
/** A line short enough to be a heading, with no stop at its end. */
const headingLike = (line: string): boolean => charsIn(line) <= HEADING_MOST && line.split(/\s+/).length <= HEADING_WORDS && !'.!?,;:…'.includes(lastOf(line));

interface Unit {
  kind: DocumentPassage['kind'];
  lines: string[];
  section?: string | undefined;
  continues?: true;
  leads?: true;
}

/** What a unit shows on a slide, one string for each paragraph. */
function shownOf(unit: Pick<Unit, 'kind' | 'lines'>): string[] {
  if (unit.kind === 'list') return unit.lines.map(itemText);
  if (unit.kind === 'table') return rowsOf(unit.lines);
  return unit.lines.map((line) => line.trim());
}
const unitChars = (unit: Pick<Unit, 'kind' | 'lines'>): number => shownOf(unit).reduce((sum, line) => sum + charsIn(line), 0);

/** Where a line of text may be cut: after each sentence it ends, with where the next one starts. */
function sentenceEnds(text: string): Array<{ end: number; next: number }> {
  const ends: Array<{ end: number; next: number }> = [];
  for (const match of text.matchAll(SENTENCE_END)) {
    const at = match.index;
    if (text[at - 1] === '.') {
      const word = (text.slice(0, at - 1).split(/\s/).pop() ?? '').replace(/^[(["'“‘«]+/u, '');
      if (NO_END.has(word.toLowerCase()) || /^\p{L}$/u.test(word)) continue;
    }
    ends.push({ end: at + match[0].replace(/\s+$/u, '').length, next: at + match[0].length });
  }
  return ends;
}

/**
 * Pieces of these sizes in `pieces` runs of about even length, as the indexes of the pieces each
 * run starts with: the runs whose longest is least (see {@link packPassages}).
 */
function evenRuns(sizes: readonly number[], pieces: number): number[] {
  const runs = packPassages(
    sizes.map((size, at) => ({ numbers: [at], size, start: false })),
    pieces,
  );
  return runs.map((run) => run[0] as number);
}

/**
 * A unit longer than {@link PASSAGE_CHARS}, cut into pieces of about even length: a paragraph at
 * the ends of its sentences, a list between its items. A sentence longer than that stays whole,
 * and a table is never cut.
 */
function cutUnit(unit: Unit): Unit[] {
  const chars = unitChars(unit);
  if (chars <= PASSAGE_CHARS || unit.kind === 'table') return [unit];
  const pieces = Math.ceil(chars / PASSAGE_CHARS);
  const parts: string[][] = [];
  if (unit.kind === 'text' && unit.lines.length === 1) {
    const text = unit.lines[0] as string;
    // The sentences, by where each starts and ends.
    const sentences: Array<{ start: number; end: number }> = [];
    let start = 0;
    for (const end of sentenceEnds(text)) {
      sentences.push({ start, end: end.end });
      start = end.next;
    }
    sentences.push({ start, end: text.length });
    const starts = evenRuns(
      sentences.map((sentence) => charsIn(text.slice(sentence.start, sentence.end))),
      pieces,
    );
    starts.forEach((first, at) => {
      const last = (starts[at + 1] ?? sentences.length) - 1;
      parts.push([text.slice((sentences[first] as { start: number }).start, (sentences[last] as { end: number }).end).trim()]);
    });
  } else {
    // A list, or text of several lines, between its lines.
    const starts = evenRuns(shownOf(unit).map(charsIn), pieces);
    starts.forEach((first, at) => parts.push(unit.lines.slice(first, starts[at + 1] ?? unit.lines.length)));
  }
  const kept = parts.filter((lines) => lines.length > 0 && lines.some((line) => line.trim() !== ''));
  return kept.map((lines, at) => ({
    kind: unit.kind,
    lines,
    section: unit.section,
    ...(at > 0 || unit.continues ? { continues: true as const } : {}),
    ...(at === kept.length - 1 && unit.leads ? { leads: true as const } : {}),
  }));
}

/** Two passages next to each other as one: a piece that goes on joins the last line of the one before. */
function joined(a: Unit, b: Unit, kind: Unit['kind']): Unit {
  const lines = b.continues && a.kind === 'text' && b.kind === 'text' ? [...a.lines.slice(0, -1), `${a.lines.at(-1) ?? ''} ${b.lines[0] ?? ''}`, ...b.lines.slice(1)] : [...a.lines, ...b.lines];
  return { kind, lines, section: a.section, ...(a.continues ? { continues: true as const } : {}), ...(b.leads ? { leads: true as const } : {}) };
}

/**
 * The passages of a document, as the planner numbers them. A line ends a paragraph, and a
 * paragraph is a passage; a paragraph longer than about 500 characters is cut at the ends of its
 * sentences into pieces of about even length, and short lines of one section go together. A run
 * of lines that start with a bullet or a number is a list, and a run of lines with a tab between
 * cells or with bars, with no blank line between them, a table. A short line on its own, or one
 * that starts with "#", is a heading: the passages after it stand under it. The text of each
 * passage is as the document has it. Pure: the same text always gives the same passages.
 */
export function documentPassages(document: string): DocumentPassage[] {
  const lines = linesOf(document);
  const filled = (at: number): boolean => (lines[at] ?? '').trim() !== '';
  // Each line is one kind, worked out from the end: a heading is one only over a line that is none.
  const kinds: Array<LineKind | undefined> = lines.map(() => undefined);
  for (let at = lines.length - 1; at >= 0; at--) {
    const line = lines[at] as string;
    if (!filled(at)) continue;
    // A word processor sets a tab after the marker of an item: a tab after its text makes a row.
    if (ITEM.test(line) && !isRow(itemText(line))) kinds[at] = 'item';
    else if (isRow(line)) kinds[at] = 'row';
    else if (MARKED.test(line.trim())) kinds[at] = 'heading';
    else {
      let next = at + 1;
      while (next < lines.length && !filled(next)) next++;
      const bare = line.trim();
      // A line right under a line of text with no stop at its end is no heading, however short: the
      // end of a paragraph broken where a page was, or a line of a list typed without bullets.
      const above = filled(at - 1) ? (lines[at - 1] as string) : undefined;
      const goesOn = above !== undefined && !isRow(above) && !ITEM.test(above) && !MARKED.test(above.trim()) && !'.!?…:;׃'.includes(lastOf(above));
      // Nor is the first line of such a list: a short line with no stop, and two more right under it.
      const listed = (step: number): boolean => filled(at + step) && kinds[at + step] === 'text' && headingLike((lines[at + step] as string).trim());
      const heading = !goesOn && !(listed(1) && listed(2)) && next < lines.length && kinds[next] !== 'heading' && headingLike(bare) && (charsIn(bare) < SOFT_WRAP || !filled(at + 1));
      kinds[at] = heading ? 'heading' : 'text';
    }
  }

  // A long line that ends in no stop goes on in the next line of text, where no blank line comes between.
  // `apart` marks a line with a blank line above it.
  const read: Array<{ kind: LineKind; text: string; apart: boolean }> = [];
  for (let at = 0; at < lines.length; at++) {
    const kind = kinds[at];
    if (kind === undefined) continue;
    const text = lines[at] as string;
    const before = read.at(-1);
    if (kind === 'text' && before?.kind === 'text' && kinds[at - 1] === 'text' && charsIn(before.text) >= SOFT_WRAP && !'.!?…:;׃'.includes(lastOf(before.text))) {
      before.text = `${before.text} ${text.trimStart()}`;
      continue;
    }
    read.push({ kind, text, apart: at > 0 && !filled(at - 1) });
  }

  // Units: a line of text, a run of items, a run of rows with no blank line between them. Headings
  // set the section of what follows.
  const units: Unit[] = [];
  let section: string | undefined;
  let headings: string[] = [];
  let last: LineKind | undefined;
  for (const { kind, text, apart } of read) {
    if (kind === 'heading') {
      headings.push(text.trim().replace(MARKED, ''));
      last = kind;
      continue;
    }
    if (headings.length > 0) {
      section = headings.join(' / ');
      headings = [];
    }
    const unit = units.at(-1);
    if (kind === 'item' && last === 'item' && unit?.kind === 'list') unit.lines.push(text);
    else if (kind === 'row' && last === 'row' && !apart && unit?.kind === 'table') unit.lines.push(text);
    else units.push({ kind: kind === 'item' ? 'list' : kind === 'row' ? 'table' : 'text', lines: [text], section });
    last = kind;
  }
  // A heading with nothing after it is text.
  for (const heading of headings) units.push({ kind: 'text', lines: [heading], section });
  for (const unit of units) {
    if (unit.kind === 'table' && unit.lines.length < 2) unit.kind = 'text';
    if (unit.kind !== 'table' && lastOf(unit.lines.at(-1) ?? '') === ':') unit.leads = true;
  }

  // Short lines of text in one section go together, up to a passage's length.
  const merged: Unit[] = [];
  const short = (unit: Unit) => unit.kind === 'text' && !unit.leads && unitChars(unit) < SHORT_PASSAGE;
  for (const unit of units) {
    const before = merged.at(-1);
    if (before && short(unit) && before.kind === 'text' && !before.leads && before.lines.every((line) => charsIn(line.trim()) < SHORT_PASSAGE) && before.section === unit.section && unitChars(before) + unitChars(unit) <= PASSAGE_CHARS) {
      before.lines.push(...unit.lines);
      continue;
    }
    merged.push({ ...unit, lines: [...unit.lines] });
  }

  // Long ones are cut, and too many are joined, the shortest pair of one section first.
  const passages = merged.flatMap(cutUnit);
  while (passages.length > MOST_PASSAGES) {
    let best: { at: number; kind: Unit['kind']; size: number } | undefined;
    for (const like of [true, false]) {
      for (let at = 0; at + 1 < passages.length; at++) {
        const a = passages[at] as Unit;
        const b = passages[at + 1] as Unit;
        if (a.section !== b.section || (like && a.kind !== b.kind)) continue;
        const size = unitChars(a) + unitChars(b);
        if (!best || size < best.size) best = { at, kind: like ? a.kind : 'text', size };
      }
      if (best) break;
    }
    if (!best) break;
    passages.splice(best.at, 2, joined(passages[best.at] as Unit, passages[best.at + 1] as Unit, best.kind));
  }

  return passages.map((unit, at): DocumentPassage => ({
    number: at + 1,
    kind: unit.kind,
    text: unit.lines.join('\n'),
    ...(unit.section !== undefined ? { section: unit.section } : {}),
    ...(unit.continues ? { continues: true as const } : {}),
    ...(unit.leads ? { leads: true as const } : {}),
    chars: unitChars(unit),
    paragraphs: shownOf(unit).length,
  }));
}

// --- counting words ---

/** The prefixes a Hebrew word may carry: "and", "the", "in", "to", "from", "that", "as". */
const HEBREW_PREFIXES = 'והבלמשכ';

/**
 * The words of a text, as the check of kept words counts them: runs of letters and digits, in one
 * form and lower case, without the marks that set direction.
 */
export function tokensOf(text: string): string[] {
  return text.replace(BIDI, '').normalize('NFC').toLowerCase().match(/[\p{L}\p{N}][\p{L}\p{M}\p{N}]*/gu) ?? [];
}

const countsOf = (tokens: readonly string[]): Map<string, number> => {
  const counts = new Map<string, number>();
  for (const token of tokens) counts.set(token, (counts.get(token) ?? 0) + 1);
  return counts;
};

/** The forms a word may have in Hebrew with one prefix more or one less: "לקוחות", "הלקוחות". */
function prefixed(word: string): string[] {
  if (!/^\p{Script=Hebrew}/u.test(word)) return [];
  const less = word.length >= 4 && HEBREW_PREFIXES.includes(word[0] as string) ? [word.slice(1)] : [];
  return [...less, ...[...HEBREW_PREFIXES].map((prefix) => `${prefix}${word}`)];
}

/**
 * How many words of a source a written text keeps, each counted once: a word the text has as it is,
 * or in Hebrew with one prefix more or less ("לקוחות" for "הלקוחות").
 */
function keptCount(source: readonly string[], written: readonly string[]): number {
  const left = countsOf(written);
  const take = (word: string): boolean => {
    const count = left.get(word) ?? 0;
    if (count === 0) return false;
    left.set(word, count - 1);
    return true;
  };
  // Words as they are first, so that a word's own form is not taken for another's with a prefix.
  const missed = source.filter((word) => !take(word));
  return source.length - missed.length + missed.filter((word) => prefixed(word).some(take)).length;
}

/** The words a passage shows on a slide, without list markers. */
const passageTokens = (passage: DocumentPassage): string[] => tokensOf(shownOf({ kind: passage.kind, lines: passage.text.split('\n') }).join('\n'));

// --- scripts ---

type Script = 'hebrew' | 'arabic' | 'cyrillic' | 'greek' | 'latin';
const SCRIPTS: Array<[Script, RegExp]> = [
  ['hebrew', /\p{Script=Hebrew}/u],
  ['arabic', /\p{Script=Arabic}/u],
  ['cyrillic', /\p{Script=Cyrillic}/u],
  ['greek', /\p{Script=Greek}/u],
  ['latin', /\p{Script=Latin}/u],
];

/** How many letters of a text are in each script known here. */
function lettersOf(text: string): Map<Script, number> {
  const counts = new Map<Script, number>();
  for (const letter of text.match(/\p{L}/gu) ?? []) {
    const script = SCRIPTS.find(([, pattern]) => pattern.test(letter))?.[0];
    if (script) counts.set(script, (counts.get(script) ?? 0) + 1);
  }
  return counts;
}

/** The script most letters of a text are written in. `undefined` for a text of no letters, or of another script. */
export function scriptOf(text: string): Script | undefined {
  const counts = lettersOf(text);
  let best: Script | undefined;
  for (const [script, count] of counts) if (best === undefined || count > (counts.get(best) ?? 0)) best = script;
  return best;
}

const LATIN = new Set(['en', 'fr', 'de', 'es', 'it', 'pt', 'nl', 'sv', 'da', 'no', 'nb', 'nn', 'fi', 'is', 'et', 'lv', 'lt', 'pl', 'cs', 'sk', 'sl', 'hr', 'bs', 'hu', 'ro', 'tr', 'id', 'ms', 'tl', 'fil', 'vi', 'sq', 'ca', 'eu', 'gl', 'ga', 'cy', 'mt', 'sw', 'af', 'zu', 'xh', 'lb', 'fo', 'gd', 'br', 'oc', 'rm', 'eo', 'la']);
/** The script a deck's language is written in, by its BCP 47 tag. `undefined` where it is not known. */
export function deckScript(language: string): Script | undefined {
  const tag = language.toLowerCase().split('-')[0] ?? '';
  if (['he', 'iw', 'yi'].includes(tag)) return 'hebrew';
  if (['ar', 'fa', 'ur'].includes(tag)) return 'arabic';
  if (['ru', 'uk', 'bg', 'be', 'mk', 'kk'].includes(tag)) return 'cyrillic';
  if (tag === 'el') return 'greek';
  return LATIN.has(tag) ? 'latin' : undefined;
}

/** Whether a text is written in the script of a deck's language: only when both are known and the same. */
export const sameScript = (text: string, language: string): boolean => {
  const script = scriptOf(text);
  return script !== undefined && script === deckScript(language);
};

// --- languages ---

/** Tags that name no one language: undetermined, several, none, and one with no tag of its own. */
const NO_LANGUAGE = new Set(['und', 'mul', 'zxx', 'mis']);
/** Old tags of some languages, read as the tags they have now. */
const OLD_TAGS: Readonly<Record<string, string>> = { iw: 'he', in: 'id', ji: 'yi' };
/** The primary subtag of a language tag, in lower case: "en" of "en-GB". */
const primaryOf = (tag: string): string => {
  const primary = tag.trim().toLowerCase().split(/[-_]/)[0] ?? '';
  return OLD_TAGS[primary] ?? primary;
};

/** A language said of a text is belied where fewer than this share of its letters are in the language's script. */
const LEAST_OWN_LETTERS = 0.1;

/**
 * Whether a language cannot be the one a text is written in: the script of the language is known,
 * and hardly any of the text's letters are in it (a Hebrew text said to be "en"). Not the script
 * most of them are in: a Hebrew memo full of English names and terms may have more Latin letters
 * than Hebrew ones, and is Hebrew.
 */
const otherScript = (text: string, language: string): boolean => {
  const own = deckScript(language);
  const counts = lettersOf(text);
  const letters = [...counts.values()].reduce((sum, count) => sum + count, 0);
  return own !== undefined && letters > 0 && (counts.get(own) ?? 0) < LEAST_OWN_LETTERS * letters;
};

/**
 * The language an answer or a stored plan says a document is written in: a BCP 47 tag such as
 * "en" or "he-IL", as given. `undefined` for anything else, and for a tag that names no one
 * language ("und", "mul"). Given the document (`text`), also for a tag its own script belies: a
 * slip of the planner's is not kept with the plan.
 */
export function readLanguageTag(value: unknown, text?: string): string | undefined {
  if (typeof value !== 'string') return undefined;
  const tag = value.trim();
  if (!/^[a-z]{2,3}(?:[-_][a-z0-9]{1,8})*$/i.test(tag)) return undefined;
  if (text !== undefined && otherScript(text, tag)) return undefined;
  return NO_LANGUAGE.has(primaryOf(tag)) ? undefined : tag;
}

/** Whether two language tags name one language: the same primary subtag ("en" and "en-GB"). */
export const sameLanguage = (a: string, b: string): boolean => primaryOf(a) === primaryOf(b);

/**
 * Whether passages are in the deck's language, so that code may set them into its slides as they
 * are. Where the planner said which language the document is written in (`written`), when that is
 * the deck's language: an English document in a French deck is in another language, though in the
 * same script. Where it did not (a plan stored before it was asked), or said a language whose
 * script is not the text's, when they are written in the script of the deck's language.
 */
function inDeckLanguage(passages: readonly DocumentPassage[], language: string, written: string | undefined): boolean {
  const text = passages.map((passage) => passage.text).join('\n');
  return written !== undefined && !otherScript(text, written) ? sameLanguage(written, language) : sameScript(text, language);
}

/**
 * A document as the planning stages read it: its passages, its length and whether it is in the
 * deck's language. `written` is the language it is written in, where the planner said.
 */
export function readDocument(document: string, language: string, written?: string): DocumentPlan {
  const passages = documentPassages(document);
  return {
    text: document,
    passages,
    chars: passages.reduce((sum, passage) => sum + passage.chars, 0),
    words: passages.reduce((sum, passage) => sum + passageTokens(passage).length, 0),
    same: inDeckLanguage(passages, language, written),
  };
}

/** The same document for a deck in `language`, written in `written` where the planner said. */
export const inLanguage = (doc: DocumentPlan, language: string, written?: string): DocumentPlan => ({ ...doc, same: inDeckLanguage(doc.passages, language, written) });

// --- passages of a slide ---

/**
 * The numbers of passages an answer or a stored plan gives a slide: whole numbers, or strings of
 * digits ("3", "p3"), from 1 to the number of passages, each once and in order. Anything else is
 * passed over, and a value that is no list gives none.
 */
export function readPassageNumbers(value: unknown, passages: number): number[] {
  if (!Array.isArray(value)) return [];
  const numbers = value.flatMap((item): number[] => {
    const number = typeof item === 'number' ? item : typeof item === 'string' && /^\s*p?\d+\s*$/i.test(item) ? Number(item.trim().replace(/^p/i, '')) : Number.NaN;
    return Number.isInteger(number) && number >= 1 && number <= passages ? [number] : [];
  });
  return [...new Set(numbers)].sort((a, b) => a - b);
}

/** The passages a slide carries, as they are named: "p4-p6, p9". */
export function passageLabel(numbers: readonly number[]): string {
  const runs: Array<[number, number]> = [];
  for (const number of [...numbers].sort((a, b) => a - b)) {
    const run = runs.at(-1);
    if (run && number === run[1] + 1) run[1] = number;
    else runs.push([number, number]);
  }
  return runs.map(([first, last]) => (first === last ? `p${first}` : `p${first}-p${last}`)).join(', ');
}

/**
 * The text of some passages as the writer is shown it: a heading in brackets wherever the section
 * changes, a piece of a paragraph that goes on from the one before joined to it, a list one item
 * to a line with its markers, a table one row to a line with its cells joined by " | ".
 */
export function passageText(all: readonly DocumentPassage[], numbers: readonly number[]): string {
  const lines: string[] = [];
  let section: string | undefined;
  let previous: number | undefined;
  for (const number of numbers) {
    const passage = all[number - 1];
    if (!passage) continue;
    if (passage.section !== undefined && (previous === undefined || passage.section !== section)) lines.push(`[${passage.section}]`);
    section = passage.section;
    const own = passage.kind === 'table' ? rowsOf(passage.text.split('\n')) : passage.text.split('\n');
    if (passage.continues && passage.kind === 'text' && previous === number - 1 && lines.length > 0) {
      lines[lines.length - 1] = `${lines.at(-1)} ${(own.shift() ?? '').trim()}`;
    }
    lines.push(...own);
    previous = number;
  }
  return lines.join('\n');
}

/**
 * The paragraphs a slide's running text has when code sets it from its passages, as given: each
 * line of a text, each item of a list with its marker, a piece that goes on from the one before
 * joined to it, and a heading that changes inside the slide as a paragraph of its own. The slide's
 * first heading is not repeated: its title is taken from it. With `bullets: false`, for a slot
 * that draws bullets of its own, an item is set without its bullet; a number stays, as the text's.
 */
export function fillParagraphs(all: readonly DocumentPassage[], numbers: readonly number[], options: { bullets?: boolean } = {}): string[] {
  const paragraphs: string[] = [];
  let section: string | undefined;
  let previous: number | undefined;
  for (const number of numbers) {
    const passage = all[number - 1];
    if (!passage) continue;
    if (previous !== undefined && passage.section !== undefined && passage.section !== section) paragraphs.push(passage.section);
    section = passage.section;
    const lines = passage.kind === 'table' ? rowsOf(passage.text.split('\n')) : passage.text.split('\n');
    // An item's marker is followed by one space, where a word processor set a tab.
    const marked = (marker: string): string => (options.bullets === false && BULLET.test(marker) ? '' : `${WORD_BULLETS[marker.trim()] ?? marker.trim()} `);
    const shown = (line: string): string => (passage.kind === 'list' ? line.replace(ITEM, marked) : line);
    const own = lines.map((line) => shown(line).trim()).filter(Boolean);
    if (passage.continues && passage.kind === 'text' && previous === number - 1 && paragraphs.length > 0) {
      paragraphs[paragraphs.length - 1] = `${paragraphs.at(-1)} ${own.shift() ?? ''}`.trimEnd();
    }
    paragraphs.push(...own);
    previous = number;
  }
  return paragraphs;
}

/** How many paragraphs a slide's text from these passages has, as code would set it (see {@link fillParagraphs}). */
export const paragraphsOf = (all: readonly DocumentPassage[], numbers: readonly number[]): number => Math.max(1, fillParagraphs(all, numbers).length);

/**
 * The headings a slide's text from these passages shows as paragraphs of their own, as
 * {@link fillParagraphs} sets them: each section that starts inside the slide. They take room on
 * the slide like any line of its text.
 */
function headingsIn(all: readonly DocumentPassage[], numbers: readonly number[]): string[] {
  const headings: string[] = [];
  let section: string | undefined;
  let previous: number | undefined;
  for (const number of numbers) {
    const passage = all[number - 1];
    if (!passage) continue;
    if (previous !== undefined && passage.section !== undefined && passage.section !== section) headings.push(passage.section);
    section = passage.section;
    previous = number;
  }
  return headings;
}

/** The headings some passages stand under, each once, in their order. */
export function sectionsOf(all: readonly DocumentPassage[], numbers: readonly number[]): string[] {
  const sections: string[] = [];
  for (const number of numbers) {
    const section = all[number - 1]?.section;
    if (section !== undefined && !sections.includes(section)) sections.push(section);
  }
  return sections;
}

/** The heading the first of some passages stands under: what a slide of them is titled, where code sets its text. */
export const openingSection = (all: readonly DocumentPassage[], numbers: readonly number[]): string | undefined => (numbers[0] !== undefined ? all[numbers[0] - 1]?.section : undefined);

/** A text as a title and a paragraph are compared: its letters and words, not its case, brackets, spaces or a stop at its end. */
const comparable = (text: string): string =>
  text
    .normalize('NFC')
    .toLocaleLowerCase()
    .replace(/^\s*\[(.*)\]\s*$/, '$1')
    .replace(/[\s:.]+$/, '')
    .replace(/\s+/g, ' ')
    .trim();

/**
 * A written slide of a deck made from a document, settled with its passages:
 * - a paragraph a writer wrote as a heading in brackets ("[Riders]"), the heading of one of the
 *   slide's passages, loses its brackets;
 * - a text of several paragraphs whose first says the slide's title again ("Background:" under
 *   the title "Background"), and is no longer than a heading, loses it: its words are on the slide;
 * - `headings` lists the paragraphs of its texts, the last of each aside, that are headings of its
 *   passages, in their order, and is left out where there are none. A heading over a list whose
 *   items lost their markers (a layout that draws bullets of its own) is not: set run in, it would
 *   lead the first item alone.
 * In a deck written in another language than its document (`translated`), the writer writes every
 * text, and a paragraph it writes in brackets, of a heading's length, is a heading it translated:
 * it loses its brackets and is marked.
 * The slide's other fields and slots are as they were.
 */
export function settleDocumentSlide(slide: SlideSpec, all: readonly DocumentPassage[], numbers: readonly number[], options: { translated?: boolean } = {}): SlideSpec {
  const sections = sectionsOf(all, numbers);
  const { headings: _headings, ...rest } = slide;
  if (sections.length === 0) return rest;
  const known = new Set(sections.map((section) => section.trim()));
  // The headings whose first passage on the slide is a list or a table: what follows them is an item or a row.
  const listed = new Set(
    sections.filter((section) => {
      const first = numbers.map((number) => all[number - 1]).find((passage) => passage?.section === section);
      return first !== undefined && first.kind !== 'text';
    }),
  );
  const title = typeof slide.slots.title === 'string' ? comparable(slide.slots.title) : '';
  const slots: Record<string, SlotValue> = {};
  const headings: string[] = [];
  for (const [name, value] of Object.entries(slide.slots)) {
    if (name === 'title' || !Array.isArray(value) || !value.every((item) => typeof item === 'string')) {
      slots[name] = value;
      continue;
    }
    const translated = new Set<string>();
    let paragraphs = (value as string[]).map((paragraph, index, all) => {
      const inner = /^\s*\[(.+)\]\s*$/.exec(paragraph)?.[1]?.trim();
      if (inner === undefined) return paragraph;
      if (known.has(inner)) return inner;
      if (!options.translated || index === all.length - 1 || [...inner].length > HEADING_MOST) return paragraph;
      translated.add(inner);
      return inner;
    });
    const first = paragraphs[0];
    if (title && first !== undefined && paragraphs.length > 1 && [...first.trim()].length <= HEADING_MOST && comparable(first) === title) paragraphs = paragraphs.slice(1);
    paragraphs.forEach((paragraph, index) => {
      const text = paragraph.trim();
      const next = paragraphs[index + 1];
      if (next === undefined || !(known.has(text) || translated.has(text))) return;
      // Over a list or a table, the next paragraph says it is an item or a row, or the heading is left plain.
      if (listed.has(text) && leadsInto(next)) return;
      headings.push(text);
    });
    slots[name] = paragraphs;
  }
  return headings.length > 0 ? { ...rest, slots, headings } : { ...rest, slots };
}

/** The words of some passages. */
export const passageWords = (all: readonly DocumentPassage[], numbers: readonly number[]): number => numbers.reduce((sum, number) => sum + (all[number - 1] ? passageTokens(all[number - 1] as DocumentPassage).length : 0), 0);

/**
 * Of some passages, those a slide shows as running text: all but its tables where the slide has a
 * table to put them in (`tables`). On a slide with none, a table's rows are paragraphs of its text.
 */
const runningOf = (doc: DocumentPlan, numbers: readonly number[], tables: boolean): number[] =>
  numbers.filter((number) => doc.passages[number - 1] !== undefined && !(tables && doc.passages[number - 1]?.kind === 'table'));

/**
 * How many characters of running text some passages take on a slide. In the deck's script, the
 * characters of the paragraphs code sets on a slide of free parts (`fillParagraphs`): their own,
 * and those of the headings that start inside the slide, each a paragraph of its own. In another,
 * the characters their words and those of the headings take once they are translated. A table
 * takes none where the slide has a table (`tables`), which it goes into.
 */
export function needOf(doc: DocumentPlan, numbers: readonly number[], language: string | undefined, tables = true): number {
  const running = runningOf(doc, numbers, tables);
  if (doc.same) return fillParagraphs(doc.passages, running, { bullets: true }).reduce((sum, paragraph) => sum + paragraph.length, 0);
  const headings = headingsIn(doc.passages, running);
  return charsOfWords(passageWords(doc.passages, running) + headings.reduce((sum, heading) => sum + tokensOf(heading).length, 0), language);
}

/**
 * The paragraphs of running text some passages give a slide (`paragraphs`): of all but its tables,
 * where it has a table (`tables`), its headings among them. In the deck's script also how long
 * each of them is where code sets them as they are (`asIs`, see {@link fillParagraphs}), with the
 * characters the slide needs for them (`needOf`), which say that it is that text.
 */
export function shapeOf(doc: DocumentPlan, numbers: readonly number[], tables = true): TextShape {
  const running = runningOf(doc, numbers, tables);
  const shape: TextShape = { paragraphs: paragraphsOf(doc.passages, running) };
  if (doc.same && running.length > 0) shape.asIs = { chars: needOf(doc, numbers, undefined, tables), each: fillParagraphs(doc.passages, running, { bullets: true }).map((paragraph) => paragraph.length) };
  return shape;
}

/** The paragraphs of a slide's text from a document: how many (see {@link shapeOf}), and, where code may set them as they are, how long each is. */
export interface TextShape {
  paragraphs: number;
  /** The lengths of the paragraphs code sets on a free slide, of a text of `chars` characters (`needOf`). Only a text of those characters is that text. */
  asIs?: { chars: number; each: number[] };
}

/** What a slide of a plan carries of its document, worked out from its passages and its words. */
export interface SlideDocument {
  /** Its passages, by number. */
  own: number[];
  /** As they are named: "p4-p6". */
  label: string;
  /** The characters of text it is to hold. */
  need: number;
  /** The paragraphs that text has. */
  shape: TextShape;
  /** The words it is to carry: those of its passages, or those asked for it. */
  words: number;
  /** The words asked for it, where it was asked to be shorter than its passages. */
  asked?: number;
  /** One of its passages is a table, which goes into a table of the slide. */
  table: boolean;
}

/**
 * What a slide carries of its document, or `undefined` for a slide without passages. Words asked
 * for it are a wish for a shorter text only when they are under four fifths of the words of its
 * passages; otherwise the slide carries its passages whole. A stored plan is not taken at its word:
 * numbers that are no passage are passed over. `tables` says the slide has a table for a table of
 * its passages; on one without, a table's rows are part of its running text, and need its room.
 */
export function slideDocument(doc: DocumentPlan, passages: unknown, words: unknown, language: string | undefined, tables = true): SlideDocument | undefined {
  const own = readPassageNumbers(passages, doc.passages.length);
  if (own.length === 0) return undefined;
  const whole = passageWords(doc.passages, own);
  const wish = readWords(words);
  const asked = wish !== undefined && wish < SHORTEN_SHARE * whole ? wish : undefined;
  const need = needOf(doc, own, language, tables);
  return {
    own,
    label: passageLabel(own),
    need: asked !== undefined ? Math.min(need, charsOfWords(asked, language)) : need,
    shape: shapeOf(doc, own, tables),
    words: asked ?? whole,
    ...(asked !== undefined ? { asked } : {}),
    table: tables && own.some((number) => doc.passages[number - 1]?.kind === 'table'),
  };
}

// --- what the planner is told ---

/** A passage as a prompt shows it: its lines on one line, joined by " / ". */
const shownLine = (passage: DocumentPassage): string => (passage.kind === 'table' ? rowsOf(passage.text.split('\n')) : passage.text.split('\n').map((line) => line.trim())).join(' / ');

/**
 * The document as a plan is shown it: its passages by number, under their headings, each with its
 * length. `change` says the deck is made from it already; `onNoSlide` names the passages no slide
 * of it carries.
 */
export function documentBlock(doc: DocumentPlan, options: { change?: boolean; onNoSlide?: ReadonlySet<number> } = {}): string {
  const lines = [
    `The document ${options.change ? 'the deck is made from' : 'to make the deck from'}, in ${doc.passages.length} numbered passages (${doc.chars} characters). A line in brackets is a heading of the document, over the passages after it. Each passage gives its length in characters:`,
  ];
  let section: string | undefined;
  for (const passage of doc.passages) {
    if (passage.section !== undefined && passage.section !== section) lines.push(`[${passage.section}]`);
    section = passage.section;
    const what = passage.kind === 'list' ? `list of ${passage.paragraphs} ${passage.paragraphs === 1 ? 'item' : 'items'}, ` : passage.kind === 'table' ? `table of ${passage.paragraphs} ${passage.paragraphs === 1 ? 'row' : 'rows'}, ` : '';
    const more = [passage.continues ? `goes on from p${passage.number - 1}` : '', options.onNoSlide?.has(passage.number) ? 'on no slide' : ''].filter(Boolean);
    lines.push(`p${passage.number} (${what}${passage.chars}${more.map((part) => `, ${part}`).join('')}): ${shownLine(passage)}`);
  }
  return lines.join('\n');
}

/** Said to an outline in place of the line on opening and closing slides, when the deck is made from a document. */
export const DOCUMENT_OPENING =
  'A deck of four or more slides opens with an opening pattern and ends with a closing pattern when the list has them, unless the brief asks for a deck without them. ' +
  'A deck of three slides or fewer carries the document on every slide: it has no opening or closing slide unless the brief asks for one.';

/** Said to an outline in place of the rule on variety, when the deck is made from a document. */
export const DOCUMENT_VARY = "Otherwise vary the patterns among those whose slots hold each slide's passages.";

/**
 * The rules of an outline made from a document, after the rule on words. `slides` is how many
 * slides the document takes of its own, and `chars` what one of them holds of it at the template's
 * text size. `given` is the number of slides the brief fixed, or the range it gave.
 */
export function documentRules(count: { slides: number; chars: number; given?: { min: number; max: number } | undefined }): string[] {
  const { given } = count;
  const share = 'A slide that carries more than it holds at the template\'s size has its text set smaller, which is better than leaving any of it out.';
  const sizes = given
    ? [`The deck has ${given.min === given.max ? given.min : `${given.min} to ${given.max}`} ${given.max === 1 ? 'slide' : 'slides'}: share the passages over those that carry them, in the document's order. ${share}`]
    : [
        `The document takes about ${count.slides} ${count.slides === 1 ? 'slide' : 'slides'} of its own: one holds about ${count.chars} characters of it at the template's text size. Give a slide the passages of one topic, up to about that much and no more than its pattern's slots hold. Use more slides rather than fuller ones, unless the brief asks for a number of slides.`,
        'When the brief, or a passage that only asks for the presentation, names a number of slides ("2 slides", "a single slide"), make exactly that many and set "asked_slides" to that number. ' +
          'When it says which parts go on a slide of their own ("each part on its own slide"), do so and set "asked_slides" to the number of slides of your answer. ' +
          `Share the passages over those slides: ${share.charAt(0).toLowerCase()}${share.slice(1)} Otherwise set "asked_slides" to 0.`,
      ];
  return [
    'The slides carry the document\'s own text. Set "passages" on each slide to the numbers of the passages it carries. Every passage goes on one slide, in the document\'s order unless the brief asks for another. ' +
      'Put a passage in "left_out" only when the brief asks to leave that part out, or when it only asks for the presentation ("make two slides from this:") and is not part of the text. ' +
      'Such a passage is part of the brief: do what it asks (a number of slides, a summary, all of the text).',
    ...sizes,
    'A heading in brackets is a good place to start a slide. A passage that ends with a colon goes on the slide of the passage after it. ' +
      'A list suits a pattern with a list or a row of items, with "items" set to its number of items; a table suits a pattern with a table.',
    'An opening, a closing or a section slide carries no passage: give it []. A slide without passages adds no facts of its own.',
    'When the brief asks for less text than the document has (a summary, short slides), set "words" on each slide that carries passages to the words it should have. ' +
      'Otherwise leave "words" out of slides that carry passages: their length comes from the passages.',
  ];
}

/** The `passages` of an entry of an outline. */
export const passagesProperty = (passages: number): JsonSchema => ({
  type: 'array',
  items: { type: 'integer', minimum: 1, maximum: passages },
  description: 'The numbers of the passages of the document this slide carries, in order. [] for a slide that carries none.',
});

/** The `left_out` of an outline, or of a change to one. */
export const leftOutProperty = (passages: number, asked: 'brief' | 'request'): JsonSchema => ({
  type: 'array',
  items: { type: 'integer', minimum: 1, maximum: passages },
  description:
    asked === 'brief'
      ? 'The numbers of the passages the brief asks to leave out of the deck, or that only ask for the presentation. [] when there are none.'
      : 'The numbers of the passages the request asks to leave out of the deck. [] when it asks for none.',
});

/** The `asked_slides` of an outline without a number of slides, or of a change to a plan. */
export const askedSlidesProperty = (asked: 'brief' | 'request'): JsonSchema => ({
  type: 'integer',
  minimum: 0,
  maximum: MOST_DOCUMENT_SLIDES,
  description:
    asked === 'brief'
      ? 'The number of slides the brief, or a passage that only asks for the presentation, names ("2 slides", "a single slide"). When it says which parts go on a slide of their own ("each part on its own slide"), the number of slides of your answer. 0 otherwise.'
      : 'The number of slides the request asks the plan to have ("keep it to 2 slides"); 0 when it names none.',
});

/** The `document_language` of an outline, or of a change to one, made from a document. */
export const documentLanguageProperty: JsonSchema = {
  type: 'string',
  description: 'The BCP 47 tag of the language the document is written in ("en", "he"), whatever the language of the deck.',
};

/** The `passages` of a slide in the plan of a change to a deck made from a document. */
export const passagesChangeProperty = (passages: number): JsonSchema => ({
  type: 'array',
  items: { type: 'integer', minimum: 1, maximum: passages },
  description: 'The numbers of the passages of the document this slide carries after the change. Leave it out to keep those it has; [] when it gives them all up.',
});

/** The rules a change to a plan made from a document is given, after the rules on words. */
export const DOCUMENT_CHANGE_RULES: readonly string[] = [
  "The plan carries the document's own text. Keep each slide's passages unless the request moves, adds or takes out some. " +
    'A part of the document the request asks for goes on the slide it names, or on a new slide in the document\'s order: passages marked "on no slide" are not in the plan yet. ' +
    'To make a slide longer, give it more passages from the slide next to it or from those on no slide. No passage is on two slides. Put a passage the request takes out in "left_out".',
  'When the request names a number of slides ("keep it to 2 slides"), the plan has exactly that many: share the passages over them and set "asked_slides" to that number. ' +
    'With three slides or fewer every slide carries passages: take out an opening or closing slide the plan has unless the request keeps it. ' +
    'A slide that carries more than it holds at the template\'s size has its text set smaller. Otherwise set "asked_slides" to 0.',
];

/** The rule on passages a change to a deck made from a document is given, after the rule on words. */
export const DOCUMENT_EDIT_RULE =
  'passages: the passages of the document a slide carries after the change. Leave it out for a slide that keeps the ones it has, and give [] when it gives them all up. ' +
  'A part of the document the request asks for goes on the slide it names, or on a new slide: passages marked "on no slide" are on no slide yet. ' +
  'To make a slide longer, give it more passages from the slide next to it or from those on no slide. No passage is on two slides. ' +
  "A slide that carries more than it holds at the template's size has its text set smaller; to have it larger, the request must ask for more slides: then give the new slide some of its passages.";

/** Whether a number of slides is fixed: one number, or a range of one. */
export const isFixedCount = (count: unknown): boolean =>
  typeof count === 'number' || (typeof count === 'object' && count !== null && (count as { min?: unknown }).min === (count as { max?: unknown }).max && (count as { min?: unknown }).min !== undefined);

/** The number of slides an answer says the brief names: a whole number from 1 to 30, or a string of its digits. `undefined` for 0 and anything else. */
export function readAskedSlides(value: unknown): number | undefined {
  const count = typeof value === 'number' ? value : typeof value === 'string' && /^\s*\d+\s*$/.test(value) ? Number(value) : Number.NaN;
  return Number.isInteger(count) && count >= 1 && count <= MOST_DOCUMENT_SLIDES ? count : undefined;
}

/** What an answer that places too little of the document is told, or `undefined` when no more than half of the passages are on no slide and not left out. */
export function grossGap(entries: ReadonlyArray<{ passages?: number[] | undefined }>, leftOut: readonly number[], passages: number): string | undefined {
  const placed = new Set([...entries.flatMap((entry) => entry.passages ?? []), ...leftOut]);
  const missing = Array.from({ length: passages }, (_, at) => at + 1).filter((number) => !placed.has(number));
  return missing.length > passages / 2 ? `passages ${passageLabel(missing)} are on no slide and not in "left_out"` : undefined;
}

// --- placing passages ---

/** Passages that go on one slide together, for packing: one, or one that ends with a colon and the one after it. */
export interface PassageAtom {
  numbers: number[];
  /** Its length, in the characters it takes on a slide. */
  size: number;
  /** It opens a section of the document: a good place to start a slide. */
  start: boolean;
}

/** Some passages as atoms: a passage that ends with a colon goes with the one after it. */
export function atomsOf(doc: DocumentPlan, numbers: readonly number[], sizeOf: (numbers: readonly number[]) => number): PassageAtom[] {
  const atoms: PassageAtom[] = [];
  for (const number of [...numbers].sort((a, b) => a - b)) {
    const atom = atoms.at(-1);
    const last = atom?.numbers.at(-1);
    if (atom && last === number - 1 && doc.passages[last - 1]?.leads) atom.numbers.push(number);
    else {
      const passage = doc.passages[number - 1];
      const before = doc.passages[number - 2];
      atoms.push({ numbers: [number], size: 0, start: passage?.section !== undefined && passage.section !== before?.section });
    }
  }
  for (const atom of atoms) atom.size = sizeOf(atom.numbers);
  return atoms;
}

/**
 * Atoms in `groups` runs of about even length, in order. First the least that the longest run can
 * be; then, of the ways whose every run is within a tenth of that, the one that starts the most
 * runs where a section of the document starts.
 */
export function packPassages(atoms: readonly PassageAtom[], groups: number): number[][] {
  const count = atoms.length;
  const k = Math.max(1, Math.min(groups, count));
  if (count === 0) return [];
  const sums = [0];
  for (const atom of atoms) sums.push((sums.at(-1) as number) + atom.size);
  const span = (from: number, to: number) => (sums[to] as number) - (sums[from] as number);
  const most = minMax(atoms, k);
  const limit = 1.1 * most;
  // starts[j][i]: the most section starts among the cuts of the first i atoms in j runs within the limit, with the longest run as a second key.
  type Best = { starts: number; longest: number; from: number };
  const best: Array<Array<Best | undefined>> = Array.from({ length: k + 1 }, () => Array<Best | undefined>(count + 1).fill(undefined));
  (best[0] as Array<Best | undefined>)[0] = { starts: 0, longest: 0, from: -1 };
  for (let j = 1; j <= k; j++) {
    for (let i = j; i <= count; i++) {
      let found: Best | undefined;
      for (let from = j - 1; from < i; from++) {
        const before = best[j - 1]?.[from];
        const size = span(from, i);
        if (!before || size > limit + 1e-9) continue;
        const option: Best = { starts: before.starts + (from > 0 && atoms[from]?.start ? 1 : 0), longest: Math.max(before.longest, size), from };
        if (!found || option.starts > found.starts || (option.starts === found.starts && option.longest < found.longest)) found = option;
      }
      (best[j] as Array<Best | undefined>)[i] = found;
    }
  }
  const runs: number[][] = [];
  let end = count;
  for (let j = k; j >= 1; j--) {
    const step = best[j]?.[end] as Best;
    runs.unshift(atoms.slice(step.from, end).flatMap((atom) => atom.numbers));
    end = step.from;
  }
  return runs;
}

/** The least that the longest of `groups` runs of these atoms can be. */
export function minMax(atoms: readonly PassageAtom[], groups: number): number {
  const count = atoms.length;
  const k = Math.max(1, Math.min(groups, count));
  if (count === 0) return 0;
  const sums = [0];
  for (const atom of atoms) sums.push((sums.at(-1) as number) + atom.size);
  let row = sums.map((sum) => sum);
  for (let j = 2; j <= k; j++) {
    const next = Array<number>(count + 1).fill(Number.POSITIVE_INFINITY);
    for (let i = j; i <= count; i++) {
      for (let from = j - 1; from < i; from++) next[i] = Math.min(next[i] as number, Math.max(row[from] as number, (sums[i] as number) - (sums[from] as number)));
    }
    row = next;
  }
  return row[count] as number;
}

/** What settling the passages of an answer needs to know of the deck. */
export interface SettleOptions {
  /**
   * Whether a slide can carry passages where it stands first or last: a free slide, a layout with
   * room for running text, or one with a table that is given only tables of the document.
   */
  carries: (entry: { pattern: string; blocks?: unknown; passages?: readonly number[] }) => boolean;
  /** The number of slides is fixed: no slide is added. */
  fixed: boolean;
  /** How many slides the document takes of its own: added for it when the answer has none to carry it. */
  slides: number;
  /** How many slides may be added. */
  room: number;
  /** The pattern of a slide added to carry passages: the smallest layout of plain text, or a free slide of text. */
  textHome: () => { pattern: string; blocks?: ['text'] };
  /**
   * In a deck of three slides or fewer whose number is fixed, which carries the document on every
   * slide: whether an opening or a closing slide that shows no running text (`end`) keeps the
   * passages the answer gave it, and becomes a slide of text (`textHome`), in place of giving them
   * to the slide next to it (`to`, `undefined` where there is none). Left out, it never does.
   */
  keepEnd?: (end: { pattern: string; blocks?: unknown; passages: number[] }, to: { pattern: string; blocks?: unknown; passages: number[] } | undefined) => boolean;
  /** The size of some passages on a slide, for sharing them out. */
  sizeOf: (numbers: readonly number[]) => number;
  /** In a change to a plan: the passages the plan had on its slides. Only those are put back where an answer drops them. */
  previous?: ReadonlySet<number>;
}

/**
 * Puts right what an answer says of the passages, in code, without asking again. A passage on
 * several slides stays on the first, and a passage both on a slide and left out stays on the slide.
 * A passage the answer forgot joins the slide of the passage before it, else of the one after. An
 * opening or a closing slide that cannot show running text gives its passages to the slide next to
 * it, or keeps them as a slide of text where `keepEnd` says so. An answer that places no passage at
 * all has them shared over its slides of content, each
 * slide then about what its passages say, unless it leaves them all out: a brief that only asks
 * for a presentation. The order of the slides is the answer's.
 */
export function settlePassages<T extends { pattern: string; purpose: string; blocks?: unknown; passages?: number[] }>(entries: readonly T[], leftOut: readonly number[], doc: DocumentPlan, options: SettleOptions): { entries: T[]; leftOut: number[] } {
  const all = doc.passages;
  const total = all.length;
  const list = entries.map((entry) => ({ ...entry, passages: [...(entry.passages ?? [])] }));
  const placed = new Set<number>();
  for (const entry of list) {
    entry.passages = entry.passages.filter((number) => !placed.has(number));
    for (const number of entry.passages) placed.add(number);
  }
  const out = new Set(leftOut.filter((number) => !placed.has(number)));
  // A first line that only asks for the presentation ("make three slides of this:") is not text to carry.
  const first = all[0];
  const asks = !options.previous && first !== undefined && first.section === undefined && first.leads === true && first.chars < 300;

  if (placed.size > 0) {
    for (let number = 1; number <= total; number++) {
      if (placed.has(number) || out.has(number) || (options.previous && !options.previous.has(number))) continue;
      if (number === 1 && asks) {
        out.add(number);
        continue;
      }
      const holder = (from: number, step: 1 | -1) => {
        for (let at = from; at >= 1 && at <= total; at += step) {
          if (!placed.has(at)) continue;
          return list.find((entry) => entry.passages.includes(at));
        }
        return undefined;
      };
      const entry = holder(number - 1, -1) ?? holder(number + 1, 1);
      if (!entry) continue;
      entry.passages = [...entry.passages, number].sort((a, b) => a - b);
      placed.add(number);
    }
    // An opening or a closing slide that shows no running text gives its passages to the slide next
    // to it. In a deck of three slides or fewer whose number is fixed it may keep them, as a slide
    // of text (`keepEnd`): where they are text that the slide next to it would cut.
    if (list.length > 1) {
      for (const [at, step] of [
        [0, 1],
        [list.length - 1, -1],
      ] as const) {
        const end = list[at] as (typeof list)[number];
        if (end.passages.length === 0 || options.carries(end)) continue;
        let to: (typeof list)[number] | undefined;
        for (let other = at + step; other >= 0 && other < list.length && !to; other += step) if ((list[other] as (typeof list)[number]).passages.length > 0) to = list[other];
        for (let other = at + step; other >= 0 && other < list.length && !to; other += step) if (options.carries(list[other] as (typeof list)[number])) to = list[other];
        if (options.keepEnd?.(end, to)) {
          const { blocks: _blocks, arrange: _arrange, items: _items, ...rest } = end as typeof end & { arrange?: unknown; items?: unknown };
          list[at] = { ...rest, ...options.textHome() } as unknown as typeof end;
          continue;
        }
        if (!to) continue;
        to.passages = [...to.passages, ...end.passages].sort((a, b) => a - b);
        end.passages = [];
      }
    }
  } else {
    if (asks) out.add(1);
    const rest = Array.from({ length: total }, (_, at) => at + 1).filter((number) => !out.has(number) && (!options.previous || options.previous.has(number)));
    if (rest.length > 0) {
      const atoms = atomsOf(doc, rest, options.sizeOf);
      // The slides of content: all but an opening and a closing slide that show no running text.
      let content = list.map((_, at) => at).filter((at) => !(list.length > 1 && (at === 0 || at === list.length - 1) && !options.carries(list[at] as (typeof list)[number])));
      if (content.length === 0) {
        if (options.fixed) content = list.map((_, at) => at);
        else {
          // No more slides than there are passages to share: an added slide with none would have nothing to say.
          const added = Math.max(1, Math.min(options.slides, options.room, atoms.length));
          const home = options.textHome();
          const made = Array.from({ length: added }, () => ({ ...home, purpose: '', passages: [] as number[] }) as unknown as (typeof list)[number]);
          list.splice(1, 0, ...made);
          content = made.map((_, at) => at + 1);
        }
      }
      const groups = packPassages(atoms, content.length);
      groups.forEach((group, at) => {
        const entry = list[content[at] as number] as (typeof list)[number];
        entry.passages = group;
        entry.purpose = purposeOf(all, group);
      });
      for (const number of rest) placed.add(number);
    }
  }
  const settled = list.map((entry) => {
    const { passages, ...rest } = entry;
    return (passages.length > 0 ? { ...rest, passages } : rest) as unknown as T;
  });
  return { entries: settled, leftOut: [...out].filter((number) => !placed.has(number)).sort((a, b) => a - b) };
}

/** What a slide given passages by code is about: the heading they stand under, or else the first words of their text. */
export function purposeOf(all: readonly DocumentPassage[], numbers: readonly number[]): string {
  const section = all[(numbers[0] ?? 1) - 1]?.section;
  if (section) return section;
  return fillParagraphs(all, numbers).join(' ').split(/\s+/).filter(Boolean).slice(0, 12).join(' ');
}

// --- checks of what was written ---

/** The words of a slide as it was written: of every text, list and table on it. */
function slideTokens(slots: Record<string, SlotValue>): string[] {
  const texts: string[] = [];
  for (const value of Object.values(slots)) {
    if (typeof value === 'string' || (Array.isArray(value) && value.every((item) => typeof item === 'string'))) texts.push(...toParagraphs(value as string | string[]));
    else if (typeof value === 'object' && value !== null && Array.isArray((value as { header?: unknown }).header) && Array.isArray((value as { rows?: unknown }).rows)) {
      const table = value as { header: unknown[]; rows: unknown[] };
      texts.push(...table.header.map(String), ...table.rows.flatMap((row) => (Array.isArray(row) ? row.map(String) : [])));
    }
  }
  return tokensOf(texts.join('\n'));
}

/** The words of some passages that a written slide keeps (see {@link keptCount}), and how many they have. */
function keptWords(all: readonly DocumentPassage[], numbers: readonly number[], slots: Record<string, SlotValue>): { kept: number; words: number } {
  const source = numbers.flatMap((number) => (all[number - 1] ? passageTokens(all[number - 1] as DocumentPassage) : []));
  return { kept: keptCount(source, slideTokens(slots)), words: source.length };
}

/** A slide the writer writes from its passages, to be checked for the words of them it keeps. */
export interface KeptCheck {
  slide: number;
  label: string;
  own: number[];
  /** The slot an issue about it is said on: its first running text. */
  slot: string;
}

/**
 * Whether a slide written from its passages is checked for the words it keeps: its text is in the
 * deck's script, it was not asked to be shorter, it has no table, it has enough words to say, and
 * its room holds nearly all of it. `undefined` when it is not.
 */
export function keptCheck(doc: DocumentPlan, slide: number, carried: SlideDocument, room: number, slot: string | undefined): KeptCheck | undefined {
  if (!doc.same || carried.asked !== undefined || carried.table || slot === undefined) return undefined;
  if (passageWords(doc.passages, carried.own) < KEPT_LEAST_TOKENS || room < KEPT_ROOM * carried.need) return undefined;
  return { slide, label: carried.label, own: carried.own, slot };
}

/** The share of the words of some passages that a slide's slots keep. `undefined` for passages of no words. */
export function keptShare(doc: DocumentPlan, own: readonly number[], slots: Record<string, SlotValue>): number | undefined {
  const { kept, words } = keptWords(doc.passages, own, slots);
  return words > 0 ? kept / words : undefined;
}

/** The share of the words of its passages that a slide of a deck keeps. */
const keptPercent = (doc: DocumentPlan, check: KeptCheck, deck: Pick<DeckSpec, 'slides'>): number | undefined => {
  const slide = deck.slides[check.slide - 1];
  return slide ? keptShare(doc, check.own, slide.slots) : undefined;
};

/**
 * The slides of a draft that keep too few of the words of their passages, sent back to the writer
 * once. A slide that has an error, or runs over a limit of characters, lines or paragraphs, is
 * passed over: it was shortened for a reason, which the check of the deck says.
 */
export function keptIssues(doc: DocumentPlan, checks: readonly KeptCheck[], deck: Pick<DeckSpec, 'slides'>, issues: readonly Issue[]): Issue[] {
  return checks.flatMap((check): Issue[] => {
    if (issues.some((issue) => issue.slide === check.slide && (issue.severity === 'error' || /characters|paragraphs|lines/.test(issue.message)))) return [];
    const share = keptPercent(doc, check, deck);
    if (share === undefined || share >= KEPT_REPAIR) return [];
    return [
      {
        severity: 'warning',
        slide: check.slide,
        slot: check.slot,
        ...said('words-kept-low-retry', { percent: Math.round(100 * share), label: check.label }),
      },
    ];
  });
}

/** Said with a deck of each slide that still keeps too few of the words of its passages. */
export function keptWarnings(doc: DocumentPlan, checks: readonly KeptCheck[], deck: Pick<DeckSpec, 'slides'>): Issue[] {
  return checks.flatMap((check): Issue[] => {
    const share = keptPercent(doc, check, deck);
    if (share === undefined || share >= KEPT_REPAIR) return [];
    return [{ severity: 'warning', slide: check.slide, ...said('words-kept-low', { percent: Math.round(100 * share), label: check.label }) }];
  });
}

/**
 * How much of its document a deck carries: the words of the document, the words of it its slides
 * have, and the passages no slide carries. A deck in another language than its document has
 * `translated` in place of the words kept, which a translation does not keep. `undefined` for a
 * deck made from no document.
 */
export function documentKept(deck: Pick<DeckSpec, 'document' | 'slides' | 'language' | 'documentLanguage'>): { words: number; kept_words?: number; translated?: true; left_out?: string } | undefined {
  if (typeof deck.document !== 'string' || !deck.document.trim()) return undefined;
  const passages = documentPassages(deck.document);
  const placed = new Set<number>();
  let kept = 0;
  for (const slide of deck.slides) {
    const own = readPassageNumbers(slide.passages, passages.length).filter((number) => !placed.has(number));
    if (own.length === 0) continue;
    for (const number of own) placed.add(number);
    kept += keptWords(passages, own, slide.slots).kept;
  }
  const words = passages.reduce((sum, passage) => sum + passageTokens(passage).length, 0);
  const missing = passages.map((passage) => passage.number).filter((number) => !placed.has(number));
  const translated = deck.language !== undefined && !inDeckLanguage(passages, deck.language, readLanguageTag(deck.documentLanguage));
  return { words, ...(translated ? { translated: true as const } : { kept_words: kept }), ...(missing.length > 0 ? { left_out: passageLabel(missing) } : {}) };
}

/** Said of a slide that holds less of its text from the document than it carries. `freeOff` adds that free slides are off. */
export const textCut = (slide: number, label: string, need: number, room: number, freeOff = false): Issue => ({
  severity: 'warning',
  slide,
  ...said('text-cut', { label, chars: need, holds: room, freeOff }),
});

/**
 * Said of a slide that holds all of its text from the document, but only with its paragraphs
 * joined: its writer is told fewer paragraphs than the text has, as the lines of a long list are.
 */
export const linesJoined = (slide: number, label: string, paragraphs: number, told: number): Issue => ({
  severity: 'warning',
  slide,
  ...said('paragraphs-joined', { label, paragraphs, holdsParagraphs: told }),
});

/** The same of a slide whose text is set small to hold it (`FreeArrange.dense`), and still holds less. */
export const textCutDense = (slide: number, label: string, need: number, room: number, size: number): Issue => ({
  severity: 'warning',
  slide,
  ...said('text-cut-dense', { label, chars: need, holds: room, size }),
});

// --- what the writer is told ---

/** The block of rules the writer of a deck made from a document is given. `filled` says some slide's text is set by code. */
export const documentWriteRule = (filled: boolean): string =>
  'Slides with text from the document say that text. Write it in its own words and order, and keep its terms, names and figures. ' +
  "Shorten it only where a slot's limit makes you: leave out the sentences that matter least, then clauses. Never sum it up, and add nothing it does not say. " +
  'Where a slot takes several strings, give one for each paragraph of the text, and join paragraphs where there are more than the slot takes. ' +
  "A title is the heading in brackets at the top of the slide's text. Where there is none there, it says in a few words what the text is about: never a heading in brackets further down, which the slide shows in its text, and not the words the text opens with. " +
  'Keep a heading in brackets inside a text you write as a paragraph of its own, without the brackets. ' +
  "A short slot (a heading, a takeaway, a bottom line) takes its words from the slide's text too: leave an optional one out rather than say what the text does not. " +
  'A slide whose purpose ends in a number such as (2/3) goes on from the slide before it, under the same title. ' +
  'A slide without text from the document adds no facts of its own.' +
  (filled ? ' A slide whose running text the software sets from the document asks you only for what the software does not set, such as its title and its short slots: take them from that text.' : '');

/** The text of a slide from its document, as the writer is shown it under the slide's line. */
export const slideText = (number: number, passage: { label: string; text: string }): string => `Text of s${number} from the document (${passage.label}):\n${passage.text}`;
