// A warning or a note in a person's language: the sentence of its code, and with a deck the whole
// line as the chat tools give it.
import type { DeckSpec } from '../deck/spec.ts';
import type { Issue } from '../deck/validate.ts';
import { slotPart } from '../plan/free.ts';
import type { IssueCatalogue, IssueCode } from './codes.ts';
import { EN } from './en.ts';
import { HE, HE_KINDS, HE_PARTS, hebrewSlot, isolate } from './he.ts';

export interface FormatIssueOptions {
  /**
   * The deck the issue is about, as the result of a tool or of `writeDeck` has it. With it the line
   * begins with the slide and the slot, or the part of a free slide ("slide 3, part 7 (note): "),
   * as the chat tools write `warnings`. Without it the sentence stands alone.
   */
  deck?: Pick<DeckSpec, 'slides'>;
}

/** The languages that have sentences of their own, by primary language subtag. */
const CATALOGUES: Record<string, IssueCatalogue> = { en: EN, he: HE, iw: HE };

/** The catalogue of a language tag ("he", "he-IL"), or `undefined` for a language that has none. */
function catalogueOf(language: string): IssueCatalogue | undefined {
  const primary = String(language ?? '').trim().toLowerCase().split(/[-_]/)[0] ?? '';
  return Object.hasOwn(CATALOGUES, primary) ? CATALOGUES[primary] : undefined;
}

/** The sentence of an issue in a catalogue, or `undefined` where its code or values cannot make one. */
function sentenceOf(issue: Issue, catalogue: IssueCatalogue): string | undefined {
  const { code, values } = issue;
  if (code === undefined || !Object.hasOwn(catalogue, code) || typeof values !== 'object' || values === null) return undefined;
  try {
    const sentence = (catalogue[code as IssueCode] as (values: unknown) => string)(values);
    // Values that cannot make the sentence (one left out, or not a number) give the English message.
    return /\bundefined\b|\bNaN\b/.test(sentence) ? undefined : sentence;
  } catch {
    return undefined;
  }
}

/**
 * An issue as a person reads it. `language` is a BCP 47 tag: English ("en") and Hebrew ("he",
 * "he-IL") have sentences of their own; any other language, and an issue without a code, gives the
 * English message. With `options.deck` it is the whole line as the chat tools write it in
 * `warnings`: an issue with a slide begins with the slide and its slot or part (in Hebrew
 * "שקף 3: ", with the part's Hebrew name); one without is its sentence alone. The English line is
 * exactly the tool's.
 */
export function formatIssue(issue: Issue, language: string, options: FormatIssueOptions = {}): string {
  const catalogue = catalogueOf(language);
  const hebrew = catalogue === HE ? sentenceOf(issue, HE) : undefined;
  if (hebrew !== undefined) return options.deck ? hebrewPrefix(issue, options.deck) + hebrew : hebrew;
  const english = (catalogue === EN ? sentenceOf(issue, EN) : undefined) ?? issue.message;
  return options.deck ? englishPrefix(issue, options.deck) + english : english;
}

/** What comes before the sentence of an issue in the line the chat tools write. */
function englishPrefix(issue: Issue, deck: Pick<DeckSpec, 'slides'>): string {
  if (issue.slide === undefined) return '';
  const part = slotPart(deck.slides?.[issue.slide - 1]?.blocks, issue.slot);
  if (part) return `slide ${issue.slide}, part ${part.number} (${part.kind})${part.name !== undefined ? ` ${part.name}` : ''}: `;
  return `slide ${issue.slide}${issue.slot ? ` ${issue.slot}` : ''}: `;
}

/**
 * The same in Hebrew, each name after a comma: "שקף 3, חלק 7 (כרטיס), כותרת: ". A slot that is no
 * part is named in Hebrew where most templates name it alike ("שקף 1, גוף הטקסט: "), and keeps
 * its own name otherwise.
 */
function hebrewPrefix(issue: Issue, deck: Pick<DeckSpec, 'slides'>): string {
  if (issue.slide === undefined) return '';
  const part = slotPart(deck.slides?.[issue.slide - 1]?.blocks, issue.slot);
  if (part) {
    const name = part.name === undefined ? '' : `, ${Object.hasOwn(HE_PARTS, part.name) ? HE_PARTS[part.name] : isolate(part.name)}`;
    return `שקף ${issue.slide}, חלק ${part.number} (${HE_KINDS[part.kind]})${name}: `;
  }
  return issue.slot ? `שקף ${issue.slide}, ${hebrewSlot(issue.slot)}: ` : `שקף ${issue.slide}: `;
}
