import type { TemplateManifest } from '../template/manifest.ts';
import { freeSlots, isFreeBlockKind } from '../free/blocks.ts';
import { isChartValue, isFreeSlide, isTableValue, isTextValue, knownValues, toParagraphs } from './spec.ts';
import type { ChartValue, DeckSpec, SlideSpec } from './spec.ts';
import type { Issue } from './validate.ts';
import { said } from '../messages/en.ts';

/**
 * The numbers a text states. "3K" counts as both 3 and 3000, "1,200" as 1200, "45%" as 45.
 * Only Western digits are read.
 */
export function statedNumbers(text: string): Set<number> {
  const numbers = new Set<number>();
  for (const match of text.matchAll(/(\d[\d,]*(?:\.\d+)?)\s*([KkMm])?/g)) {
    const value = Number((match[1] as string).replace(/,/g, ''));
    if (!Number.isFinite(value)) continue;
    numbers.add(value);
    const suffix = match[2]?.toLowerCase();
    if (suffix === 'k') numbers.add(value * 1000);
    if (suffix === 'm') numbers.add(value * 1_000_000);
  }
  return numbers;
}

/**
 * Chart values that the deck's brief, or the document it was made from, does not state. A model
 * asked for a chart can fill it with plausible numbers; this flags them so a person can check.
 * Values worked out from the brief (a sum, a growth rate applied) are flagged too: the check
 * cannot tell them from invented ones.
 */
export function unstatedChartValues(manifest: TemplateManifest, deck: DeckSpec): Issue[] {
  // A deck made from a document may chart the figures the document states.
  const stated = statedNumbers(deck.document?.trim() ? `${deck.brief ?? ''}\n${deck.document}` : (deck.brief ?? ''));
  const issues: Issue[] = [];
  deck.slides.forEach((slide, index) => {
    const pattern = manifest.patterns.find((candidate) => candidate.id === slide.pattern);
    // A free slide names its charts itself: block 3 of kind chart is the slot b3Chart.
    const charts = isFreeSlide(slide) ? freeSlots((slide.blocks ?? []).filter(isFreeBlockKind)).filter((entry) => entry.part.holds === 'chart').map((entry) => ({ name: entry.name, kind: 'chart' as const })) : (pattern?.slots ?? []);
    for (const slot of charts) {
      const value = slide.slots[slot.name];
      if (slot.kind !== 'chart' || !isChartValue(value)) continue;
      const unknown = [...new Set(knownValues(value))].filter((number) => !stated.has(number));
      if (unknown.length > 0) {
        issues.push({
          severity: 'warning',
          slide: index + 1,
          slot: slot.name,
          ...said('chart-values-unstated', { numbers: unknown.join(', '), count: unknown.length }),
        });
      }
    }
  });
  return issues;
}

/** What a deck was given to say: its brief, the document it was made from, and the values of its fields. */
function givenText(deck: DeckSpec): string {
  return [deck.brief ?? '', deck.document ?? '', ...Object.values(deck.fields ?? {})].join('\n');
}

/** The digits of a text, for telling whether a phone number is the one given in another form. */
const digitsOf = (text: string): string => text.replace(/\D/g, '');

/**
 * Details a model can make up that read as facts: an email or web address, a phone number (that
 * starts with + or 0, as numbers to dial do, and has at least eight digits), "24/7", a percentage.
 * Each with how to tell that the given text states it.
 */
const DETAILS: Array<{ find: RegExp; stated: (found: string, given: string) => boolean; is?: (found: string) => boolean }> = [
  { find: /[\w.+-]+@[\w-]+(?:\.[\w-]+)+/g, stated: (found, given) => given.toLowerCase().includes(found.toLowerCase()) },
  { find: /\b(?:https?:\/\/|www\.)[^\s,;)"'<>]+/gi, stated: (found, given) => given.toLowerCase().includes(found.replace(/[.!?]+$/, '').toLowerCase()) },
  {
    find: /(?<![\w+])(?:\+|0)\d[\d\s().-]{6,}\d(?!\d)/g,
    stated: (found, given) => [...given.matchAll(/(?<![\w+])(?:\+|0)\d[\d\s().-]{6,}\d(?!\d)/g)].some((match) => digitsOf(match[0]) === digitsOf(found)),
    // A number to dial has nine digits or more; a date (05.11.2026, 01-12-2026) is none.
    is: (found) => digitsOf(found).length >= 9 && !/^\d{1,2}[./-]\d{1,2}[./-]\d{2,4}$/.test(found.trim()),
  },
  { find: /\b24\s*\/\s*7\b/g, stated: (_found, given) => /\b24\s*\/\s*7\b/.test(given) },
  { find: /(?<![\d.,])\d+(?:\.\d+)?\s?%/g, stated: (found, given) => statedNumbers(given).has(Number(found.replace(/[\s%]/g, ''))) },
];

/** Text that stands where a real value belongs: "XX%", "[name]", "TBD", "lorem ipsum", "???", "____", "John Doe", "ישראל ישראלי". */
const STAND_INS = /(?<![A-Za-z])X{2,}(?![A-Za-z])|\[[^\]\d]*\p{L}[^\]\d]{0,30}\]|\bTB[DA]\b|lorem ipsum|\?{3,}|_{3,}|\b(?:John|Jane) Doe\b|ישראל ישראלי|פלוני(?: אלמוני)?/giu;

/** Text that speaks about the slide where it should say its content: "this slide", "the chart below", "בשקף זה". */
const ABOUT_SLIDE =
  /\b(?:this|the following|the next|the previous) slide\b|\bon (?:this|the) slide\b|\bslide (?:below|above)\b|\b(?:see|in) the (?:chart|table|picture|image|graph) (?:below|above)\b|\bas shown (?:below|above)\b|(?:^|[\s"'(])[בה]?שקף (?:זה|הזה|הבא|הקודם)(?=$|[\s.,:;!?)"'])|(?:בתרשים|בטבלה|בתמונה|בגרף)\s+(?:שלהלן|שלמטה|הבא|הבאה)|כפי שמוצג (?:בשקף|בתרשים|בטבלה|להלן)/iu;

/** The texts of a slide by slot: its texts, and the cells of its tables as one text. */
function slideTexts(slots: Record<string, unknown>): Array<[slot: string, text: string]> {
  const texts: Array<[string, string]> = [];
  for (const [slot, value] of Object.entries(slots)) {
    if (isTextValue(value)) texts.push([slot, toParagraphs(value).join('\n')]);
    else if (isTableValue(value)) texts.push([slot, [value.header, ...value.rows].map((row) => row.map(String).join(' | ')).join('\n')]);
  }
  return texts;
}

/**
 * What the slides of a deck state that a person should check, as the chart numbers the brief does
 * not state are checked: contact details and figures the brief, the document and the deck's fields
 * do not give (an email address on a closing slide, a phone number, "24/7", a percentage), text that
 * stands where a value belongs ("XX%", "[name]"), and text about the slide itself ("this slide",
 * "בשקף זה"). Warnings to show a person, not problems for a model to repair: a detail the brief
 * gives in other words is flagged too. `slides` limits the check to those slides, by number.
 */
export function unstatedDetails(deck: DeckSpec, slides?: ReadonlySet<number>): Issue[] {
  const given = givenText(deck);
  const issues: Issue[] = [];
  deck.slides.forEach((slide, index) => {
    if (slides && !slides.has(index + 1)) return;
    for (const [slot, text] of slideTexts(slide.slots ?? {})) {
      // In the order the text says them; an address at the end of a sentence without its full stop.
      const seen: Array<{ at: number; value: string }> = [];
      for (const detail of DETAILS) {
        for (const match of text.matchAll(detail.find)) {
          const value = match[0].trim().replace(/[.,;:!?]+$/, '');
          if ((detail.is?.(value) ?? true) && !detail.stated(value, given)) seen.push({ at: match.index ?? 0, value });
        }
      }
      const unknown = new Set(seen.sort((a, b) => a.at - b.at).map((entry) => entry.value));
      if (unknown.size > 0) issues.push({ severity: 'warning', slide: index + 1, slot, ...said('text-details-unstated', { details: [...unknown].join(', '), count: unknown.size }) });
      const stand = [...new Set([...text.matchAll(STAND_INS)].map((match) => match[0]))].filter((found) => !given.includes(found));
      if (stand.length > 0) issues.push({ severity: 'warning', slide: index + 1, slot, ...said('text-stand-in', { found: stand.join(', ') }) });
      const about = ABOUT_SLIDE.exec(text)?.[0]?.trim().replace(/^["'(]/, '');
      if (about) issues.push({ severity: 'warning', slide: index + 1, slot, ...said('text-about-slide', { found: about }) });
    }
  });
  return issues;
}

/**
 * A slide written again in a change, with the gaps of its charts kept. A writer is given numbers
 * only, so where a chart had a value that is not known (`null`), it writes a number: that value
 * stays a gap unless `stated` (the numbers of the request, the brief and the document) has the new
 * number. The chart is matched by its slot, its categories and series by their place where it has
 * as many as before (a translated chart renames them), else its series by their names.
 */
export function keptGaps(before: SlideSpec, after: SlideSpec, stated: ReadonlySet<number>): SlideSpec {
  let changed = false;
  const slots = { ...after.slots };
  for (const [name, value] of Object.entries(after.slots ?? {})) {
    const old = before.slots?.[name];
    if (!isChartValue(value) || !isChartValue(old) || old.categories.length !== value.categories.length) continue;
    const byPlace = old.series.length === value.series.length;
    const series = value.series.map((entry, index): ChartValue['series'][number] => {
      const was = byPlace ? old.series[index] : old.series.find((candidate) => candidate.name === entry.name);
      if (!was || !Array.isArray(was.values) || !Array.isArray(entry.values)) return entry;
      const values = entry.values.map((number, at) => (was.values[at] === null && typeof number === 'number' && !stated.has(number) ? null : number));
      return values.some((number, at) => number !== entry.values[at]) ? { ...entry, values } : entry;
    });
    if (series.some((entry, index) => entry !== value.series[index])) {
      slots[name] = { ...value, series };
      changed = true;
    }
  }
  return changed ? { ...after, slots } : after;
}
