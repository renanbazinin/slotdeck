// What a warning or a note is, as a code, and the values its sentence is made of. A code is stable:
// once it has shipped it is never renamed. An app shows an issue in its own language with
// `formatIssue`, from these two alone.

/**
 * The values of each code: the numbers and labels its sentence states. A key that is optional is
 * left out where the sentence does not say it: `slide` where the issue's own `slide` is set (the
 * line then begins with the slide), `size` where the text is not set small.
 */
export interface IssueValues {
  /** A one-line box whose text is wider than the box. `shrunk`: it was set smaller first ("even at"). */
  'text-too-wide': { approximate: boolean; percent: number; shrunk: boolean; size: number; fitsChars: number };
  /** Measured text that runs over its box by `over` lines. */
  'text-overflow': { approximate: boolean; lines: number; shrunk: boolean; size: number; capacity: number; over: number; fitsChars: number };
  'colors-not-object': { slide?: number };
  /** `given` is the value as the deck gave it, in JSON. */
  'color-not-hex': { slide?: number; name: string; given: string };
  'colors-replace-not-map': { slide?: number };
  /** `from` and `to` are in JSON, as the deck gave them. */
  'colors-replace-not-pair': { slide?: number; from: string; to: string };
  'colors-unknown-setting': { slide?: number; name: string };
  'colors-kept-by-template': Record<string, never>;
  /** `given` is the deck's `arrange.columns`, in JSON. */
  'columns-not-whole': { slide: number; given: string };
  'columns-relaxed': { slide: number; asked: number; columns: number };
  'free-text-small': { slide: number; blocks: number; size: number };
  /** `half`: the slide has room for half of the larger text; otherwise for none of it. */
  'free-larger-text': { slide: number; blocks: number; half: boolean };
  'free-charts-bare': { slide: number; blocks: number; charts: number };
  'free-blocks-low': { slide: number; blocks: number; count: number };
  /** `why` is the reason in English, as `freeBaseRefusal` gives it. */
  'free-base-refused': { slide: number; base: string; why: string; stands: string };
  'free-base-gone': { slide: number; base: string; stands: string };
  'items-ignored': { pattern: string };
  'slot-unknown': { pattern: string };
  'slot-over-chars': { chars: number; most: number };
  'slot-over-paragraphs': { paragraphs: number; most: number };
  'table-over-columns': { columns: number; most: number };
  'table-over-rows': { rows: number; most: number };
  'table-cell-long': { chars: number; most: number };
  'chart-over-categories': { categories: number; most: number };
  'chart-over-series': { series: number; most: number };
  /** `format` is the value as the deck gave it, in JSON. */
  'chart-format-unknown': { format: string };
  /** `numbers` are the values, joined with ", ". */
  'chart-values-unstated': { numbers: string; count: number };
  /** Contact details and figures in a text that the brief does not give, as they stand on the slide. */
  'text-details-unstated': { details: string; count: number };
  /** Text that stands where a value belongs, such as "XX%" or "[name]". */
  'text-stand-in': { found: string };
  /** Words in which a text speaks about the slide, such as "this slide". */
  'text-about-slide': { found: string };
  /** `word` is the lone word as the message quotes it: in JSON, cut at 30 characters. */
  'text-orphan': { lines: number; word: string; chars: number };
  /** One of `sampleLines` (the lines the template's own text takes) and `capacity` (the lines the box holds). */
  'text-empty': { lines: number; sampleLines?: number; capacity?: number; chars: number };
  /** `beside` is the slot of the fullest text of the row. */
  'row-uneven': { lines: number; beside: string; besideLines: number; chars: number };
  /** `problem` is the review model's own sentence, in the language it wrote. */
  'seen-in-picture': { problem: string };
  'accent2-missing': { slide?: number };
  'color-near-page': { hex: string };
  /** Sent to the writer only: collapsed into `parts-unwritten` before a deck is returned. */
  'part-unwritten': Record<string, never>;
  /** Sent to the writer only: the deck says `photos-left-out`. */
  'photos-over': { most: number };
  'parts-left-out-empty': { slide: number; count: number };
  'parts-left-out-unusable': { slide: number; count: number };
  'photos-left-out': { slide: number; count: number; most: number };
  'parts-unchanged': { slide: number; count: number };
  'parts-unwritten': { slide: number; count: number };
  /** Sent to the writer only, once. */
  'written-under-retry': { chars: number; words: number; least: number; most: number };
  'written-under': { words: number; hasWords: number };
  /** `freeOff`: free slides are off, so the text is not set smaller. */
  'words-cut': { words: number; holdsWords: number; freeOff: boolean };
  'words-cut-dense': { words: number; holdsWords: number; size: number };
  /** `asked`: the words were asked for the slide now; otherwise they are the text it has. */
  'photo-left-out': { words: number; asked: boolean };
  /** Sent to the writer only, once. `label` names the passages, such as "p3-p5". */
  'words-kept-low-retry': { percent: number; label: string };
  'words-kept-low': { percent: number; label: string };
  /** `chars` is the length of the text, `holds` what the slide holds, both in characters. */
  'text-cut': { label: string; chars: number; holds: number; freeOff: boolean };
  'paragraphs-joined': { label: string; paragraphs: number; holdsParagraphs: number };
  'text-cut-dense': { label: string; chars: number; holds: number; size: number };
  'nothing-changed': { slide: number };
  'recolor-slide-removed': { slide: number };
  /** A note of `edit_presentation` that stops before a cut: one for each slide. `size` where the slide's text is set small. */
  'change-would-cut': { slide: number; words: number; holdsWords: number; size?: number; freeOff: boolean };
  /** A note of a plan: a slide that holds less than its text. `size` where its text is set well below the template's. */
  'slide-holds-less': { slide: number; words: number; holdsWords: number; size?: number; freeOff: boolean };
  /** A note of a plan: a slide that holds all its text, in type well below the template's. */
  'slide-set-small': { slide: number; size: number; templateSize: number };
  /** A note of a plan: a listed text that the topic retells, by its name. */
  'text-retold': { name: string };
}

/** What a warning or a note is. Stable: a code is never renamed once it has shipped. */
export type IssueCode = keyof IssueValues;

/** Every code, in the order of {@link IssueValues}. */
export const ISSUE_CODES: readonly IssueCode[] = [
  'text-too-wide',
  'text-overflow',
  'colors-not-object',
  'color-not-hex',
  'colors-replace-not-map',
  'colors-replace-not-pair',
  'colors-unknown-setting',
  'colors-kept-by-template',
  'columns-not-whole',
  'columns-relaxed',
  'free-text-small',
  'free-larger-text',
  'free-charts-bare',
  'free-blocks-low',
  'free-base-refused',
  'free-base-gone',
  'items-ignored',
  'slot-unknown',
  'slot-over-chars',
  'slot-over-paragraphs',
  'table-over-columns',
  'table-over-rows',
  'table-cell-long',
  'chart-over-categories',
  'chart-over-series',
  'chart-format-unknown',
  'chart-values-unstated',
  'text-details-unstated',
  'text-stand-in',
  'text-about-slide',
  'text-orphan',
  'text-empty',
  'row-uneven',
  'seen-in-picture',
  'accent2-missing',
  'color-near-page',
  'part-unwritten',
  'photos-over',
  'parts-left-out-empty',
  'parts-left-out-unusable',
  'photos-left-out',
  'parts-unchanged',
  'parts-unwritten',
  'written-under-retry',
  'written-under',
  'words-cut',
  'words-cut-dense',
  'photo-left-out',
  'words-kept-low-retry',
  'words-kept-low',
  'text-cut',
  'paragraphs-joined',
  'text-cut-dense',
  'nothing-changed',
  'recolor-slide-removed',
  'change-would-cut',
  'slide-holds-less',
  'slide-set-small',
  'text-retold',
];

/** The sentences of one language, by code: each a function of the code's values. */
export type IssueCatalogue = { [C in IssueCode]: (values: IssueValues[C]) => string };
