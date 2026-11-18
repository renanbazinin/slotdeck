// The English sentence of every code. Each one is the message slotdeck has always given: the
// builders of the warnings make their messages here, so that a message and its code never part.
import type { IssueCatalogue, IssueCode, IssueValues } from './codes.ts';

/** Said at the end of a cut where free slides are off: no slide's text is then set smaller. */
export const FREE_OFF = '; free slides are off, so its text is not set smaller';

const at = (slide: number | undefined): string => (slide !== undefined ? `slide ${slide}: ` : '');
const characters = (count: number): string => (count === 1 ? '1 character' : `${count} characters`);
const lines = (count: number): string => `${count} ${count === 1 ? 'line' : 'lines'}`;
const parts = (count: number): string => `${count} ${count === 1 ? 'part' : 'parts'}`;
/** The start a free slide's notes share: the slide and its number of blocks. */
const holds = (slide: number, blocks: number): string => `slide ${slide} holds ${blocks === 1 ? '1 block' : `${blocks} blocks`}, as asked`;
const fitsHint = (fitsChars: number): string => `Text of about ${characters(fitsChars)} fits.`;
const atSize = (shrunk: boolean, size: number): string => `${shrunk ? 'even at ' : 'at '}${size} pt`;
const textAt = (size: number | undefined): string => (size !== undefined ? ` (text at ${size} pt)` : '');

export const EN: IssueCatalogue = {
  'text-too-wide': (v) => `text is ${v.approximate ? 'about ' : ''}${v.percent}% wider than its box ${atSize(v.shrunk, v.size)}, and the box does not wrap. ${fitsHint(v.fitsChars)}`,
  'text-overflow': (v) =>
    `text takes ${v.approximate ? 'about ' : ''}${v.lines === 1 ? '1 line' : `${v.lines} lines`} ${atSize(v.shrunk, v.size)}; the box holds ${v.capacity} (${lines(v.over)} too many). ${fitsHint(v.fitsChars)}`,
  'colors-not-object': (v) => `${at(v.slide)}"colors" must be an object such as { "accent": "1565C0" }; it is ignored`,
  'color-not-hex': (v) => `${at(v.slide)}the colour "${v.name}" is ${v.given}, which is not six hex digits such as "1565C0"; it is ignored`,
  'colors-replace-not-map': (v) => `${at(v.slide)}"replace" must map colours to colours, such as { "0E7C7B": "1565C0" }; it is ignored`,
  'colors-replace-not-pair': (v) => `${at(v.slide)}the replacement ${v.from}: ${v.to} is not a pair of colours of six hex digits; it is ignored`,
  'colors-unknown-setting': (v) => `${at(v.slide)}"colors" has no setting called "${v.name}"; it is ignored`,
  'colors-kept-by-template': () => 'the template keeps its own colours in every deck; the colours the deck asks for are left out',
  'columns-not-whole': (v) => `slide ${v.slide}: the number of blocks side by side must be a whole number of 1 or more, not ${v.given}; it is ignored`,
  'columns-relaxed': (v) => `slide ${v.slide}: ${v.asked === 1 ? '1 side by side was' : `${v.asked} side by side were`} asked for; the slide holds ${v.columns}`,
  'free-text-small': (v) => `${holds(v.slide, v.blocks)}; its smallest text is set at ${v.size} pt`,
  'free-larger-text': (v) => `${holds(v.slide, v.blocks)}; the deck asks for larger text, and this slide has ${v.half ? 'room for half of that' : 'no room for it'}`,
  'free-charts-bare': (v) => `${holds(v.slide, v.blocks)}; ${v.charts === 1 ? 'one chart is' : `${v.charts} charts are`} drawn small, without axes, legend or labels`,
  'free-blocks-low': (v) => `${holds(v.slide, v.blocks)}; ${v.count === 1 ? 'one of them reaches' : `${v.count} of them reach`} below the room the slide has, and may be cut off at its edge`,
  'free-base-refused': (v) => `slide ${v.slide} was planned on the frame of the layout "${v.base}", which cannot carry a free slide (${v.why}); it stands on "${v.stands}"`,
  'free-base-gone': (v) => `slide ${v.slide} was planned on the frame of the layout "${v.base}", which the template no longer offers; it stands on "${v.stands}"`,
  'items-ignored': (v) => `pattern "${v.pattern}" has a fixed number of items; "items" is ignored`,
  'slot-unknown': (v) => `pattern "${v.pattern}" has no such slot; ignored`,
  'slot-over-chars': (v) => `text is ${v.chars} characters; the slot fits about ${v.most}`,
  'slot-over-paragraphs': (v) => `text has ${v.paragraphs} paragraphs; the slot fits ${v.most}`,
  'table-over-columns': (v) => `the table has ${v.columns} columns; the slot fits ${v.most}`,
  'table-over-rows': (v) => `the table has ${v.rows === 1 ? '1 row' : `${v.rows} rows`}; the slot fits ${v.most}`,
  'table-cell-long': (v) => `a cell is ${v.chars} characters; cells fit about ${v.most}`,
  'chart-over-categories': (v) => `the chart has ${v.categories} categories; the slot fits ${v.most}`,
  'chart-over-series': (v) => `the chart has ${v.series} series; the slot fits ${v.most}`,
  'chart-format-unknown': (v) => `chart format ${v.format} is not known and is ignored; use "number" or "percent"`,
  'chart-values-unstated': (v) => `chart values not stated in the brief: ${v.numbers}. Check them before using the deck.`,
  'text-orphan': (v) => `the last of its ${lines(v.lines)} holds one word, ${v.word}. Text of about ${v.chars} characters ends a line sooner; a few words more fill the line.`,
  'text-empty': (v) =>
    `text takes ${lines(v.lines)} ${v.sampleLines !== undefined ? `where the template's own text takes ${v.sampleLines}` : `in a box that holds ${v.capacity}`}, so its box will look empty. Text of about ${v.chars} characters fills it as designed.`,
  'row-uneven': (v) => `text takes ${lines(v.lines)} where "${v.beside}" beside it takes ${v.besideLines}, so the row will look uneven. Text of about ${v.chars} characters or more evens it out.`,
  'seen-in-picture': (v) => `seen in the picture of the slide: ${v.problem}`,
  'accent2-missing': (v) => `${at(v.slide)}the template has no second colour; "accent2" is left out`,
  'color-near-page': (v) => `the colour ${v.hex} is close to the page's own; fills, chart bars and icons in it will hardly show`,
  'part-unwritten': () => 'nothing was written for it, and it is to be written again',
  'photos-over': (v) => `at most ${v.most} photos can be made for one slide; use a supplied photo for the others`,
  'parts-left-out-empty': (v) => `slide ${v.slide}: ${parts(v.count)} had no content and ${v.count === 1 ? 'was' : 'were'} left out`,
  'parts-left-out-unusable': (v) => `slide ${v.slide}: ${parts(v.count)} could not be shown and ${v.count === 1 ? 'was' : 'were'} left out`,
  'photos-left-out': (v) => `slide ${v.slide}: ${v.count} ${v.count === 1 ? 'photo was' : 'photos were'} left out, since at most ${v.most} are made for one slide`,
  'parts-unchanged': (v) =>
    `slide ${v.slide}: what was written for ${parts(v.count)} could not be shown, and ${v.count === 1 ? 'it stays' : 'they stay'} as ${v.count === 1 ? 'it was' : 'they were'}`,
  'parts-unwritten': (v) => `slide ${v.slide}: nothing was written for ${parts(v.count)}, and ${v.count === 1 ? 'it stays as it was' : 'they stay as they were'}`,
  'written-under-retry': (v) => `is ${v.chars} characters, and ${v.words} words were asked for this slide: write ${v.least} to ${v.most} characters, as was asked`,
  'written-under': (v) => `${v.words} words were asked for this slide, and it has about ${v.hasWords}: ask for it again to have them all`,
  'words-cut': (v) => `${v.words} words were asked for this slide, and it holds about ${v.holdsWords}: ask for the text over more slides, or for a shorter one${v.freeOff ? FREE_OFF : ''}`,
  'words-cut-dense': (v) =>
    `${v.words} words were asked for this slide, and it holds about ${v.holdsWords} even with its text set small, at ${v.size} pt: the rest is cut. Ask for the text over more slides, or for a shorter one`,
  'photo-left-out': (v) => `the photo of this slide is left out to make room for ${v.asked ? `the ${v.words} words asked for it` : `its text of about ${v.words} words`}`,
  'words-kept-low-retry': (v) =>
    `keeps ${v.percent}% of the words of its text from the document (${v.label}), and the slide has room for nearly all of it: write that text in its own words and order, shortened only where a limit makes you`,
  'words-kept-low': (v) => `the slide keeps ${v.percent}% of the words of its text from the document (${v.label}), though it has room for nearly all of it`,
  'text-cut': (v) =>
    `the text from the document for this slide (${v.label}) is ${v.chars} characters, and the slide holds about ${v.holds}: it is cut to fit. Ask for more slides to keep it whole${v.freeOff ? FREE_OFF : ''}`,
  'paragraphs-joined': (v) =>
    `the text from the document for this slide (${v.label}) has ${v.paragraphs} paragraphs, and the slide holds it only as ${v.holdsParagraphs}: its paragraphs are joined. Ask for more slides to keep them apart`,
  'text-cut-dense': (v) =>
    `the text from the document for this slide (${v.label}) is ${v.chars} characters, and the slide holds about ${v.holds} even with its text set small, at ${v.size} pt: it is cut to fit. Ask for more slides to keep it whole`,
  'nothing-changed': (v) => `slide ${v.slide}: nothing on it was changed: the plan of the change named no part to write again`,
  'recolor-slide-removed': (v) => `the change removed slide ${v.slide}, which was to be recoloured`,
  'change-would-cut': (v) => `slide ${v.slide} would hold about ${v.holdsWords} of its ${v.words} words${textAt(v.size)}, so the change was not made${v.freeOff ? FREE_OFF : ''}`,
  'slide-holds-less': (v) => `slide ${v.slide} holds about ${v.holdsWords} of its ${v.words} words${textAt(v.size)}: the rest will be cut${v.freeOff ? FREE_OFF : ''}`,
  'slide-set-small': (v) => `slide ${v.slide} carries all its text in smaller type than the template's: ${v.size} pt, where the template's is ${v.templateSize} pt`,
  'text-retold': (v) => `the topic retells "${v.name}", and the slides will not carry that text word for word`,
};

/** A warning's message with its code and values, as `said` makes it: the parts of an `Issue` that say what it is. */
export interface Told {
  message: string;
  code: IssueCode;
  values: Record<string, string | number | boolean>;
}

/**
 * A code with its values and its English message, ready to spread into an `Issue`. A value left
 * undefined is not kept: the values hold what the sentence states, and nothing that is not set.
 */
export function said<C extends IssueCode>(code: C, values: IssueValues[C]): Told & { code: C } {
  const kept = Object.fromEntries(Object.entries(values).filter(([, value]) => value !== undefined)) as Record<string, string | number | boolean>;
  return { message: EN[code](values), code, values: kept };
}
