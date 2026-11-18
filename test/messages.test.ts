// What a warning or a note is, as a code with values, and its sentence in English and in Hebrew.
// An app shows the tools' warnings and notes in the person's language with formatIssue; the English
// line is exactly the line the tools write. No test calls a live model.
import assert from 'node:assert/strict';
import { readdirSync, readFileSync } from 'node:fs';
import { dirname, join } from 'node:path';
import { test } from 'node:test';
import { fileURLToPath } from 'node:url';
import { ISSUE_CODES, TOOL_NAMES, checkDesign, formatIssue, freeSlide, loadTemplate, memoryStore, presentationTools, renderDeck, validateDeck } from '../src/index.ts';
import type { DeckSpec, FontMetrics, Issue, IssueCode, IssueValues, JsonRequest, LlmProvider, TemplateManifest, ToolResult } from '../src/index.ts';
import { unstatedChartValues } from '../src/deck/numbers.ts';
import { overflowTold } from '../src/deck/validate.ts';
import { EN } from '../src/messages/en.ts';
import { HE, hebrewSlot } from '../src/messages/he.ts';
import type { TextFit } from '../src/text/measure.ts';
import { assertCoded } from './issues.ts';

const root = join(dirname(fileURLToPath(import.meta.url)), '..');
const manifest = JSON.parse(readFileSync(join(root, 'templates', 'demo.manifest.json'), 'utf8')) as TemplateManifest;
const demoPptx = new Uint8Array(readFileSync(join(root, 'templates', 'demo.pptx')));
const demo = loadTemplate(demoPptx, manifest);
const fields = { organizationDepartment: 'Riverbend' };
const deckOf = (...slides: DeckSpec['slides']): DeckSpec => ({ version: 1, fields, slides });

/** A realistic set of values for every code: the examples of the documentation. */
const SAMPLES: { [C in IssueCode]: IssueValues[C] } = {
  'text-too-wide': { approximate: true, percent: 30, shrunk: false, size: 12, fitsChars: 41 },
  'text-overflow': { approximate: false, lines: 7, shrunk: true, size: 15.5, capacity: 5, over: 2, fitsChars: 290 },
  'colors-not-object': { slide: 2 },
  'color-not-hex': { name: 'accent', given: '"blue"' },
  'colors-replace-not-map': {},
  'colors-replace-not-pair': { slide: 3, from: '"0E7C7B"', to: '"teal"' },
  'colors-unknown-setting': { name: 'background' },
  'colors-kept-by-template': {},
  'columns-not-whole': { slide: 4, given: '2.5' },
  'columns-relaxed': { slide: 4, asked: 12, columns: 5 },
  'free-text-small': { slide: 4, blocks: 60, size: 5.5 },
  'free-larger-text': { slide: 4, blocks: 9, half: true },
  'free-charts-bare': { slide: 4, blocks: 200, charts: 200 },
  'free-blocks-low': { slide: 4, blocks: 9, count: 1 },
  'free-base-refused': {
    slide: 4,
    base: 'closing',
    why: `the layout "closing" has its title low on the slide: blocks would get 25% of the slide's height, and a frame leaves them 45% or more`,
    stands: 'section',
  },
  'free-base-gone': { slide: 4, base: '3-cards', stands: 'text-only' },
  'items-ignored': { pattern: 'cover' },
  'slot-unknown': { pattern: 'text-only' },
  'slot-over-chars': { chars: 230, most: 211 },
  'slot-over-paragraphs': { paragraphs: 5, most: 4 },
  'table-over-columns': { columns: 6, most: 5 },
  'table-over-rows': { rows: 12, most: 8 },
  'table-cell-long': { chars: 40, most: 22 },
  'chart-over-categories': { categories: 12, most: 8 },
  'chart-over-series': { series: 4, most: 3 },
  'chart-format-unknown': { format: '"currency"' },
  'chart-values-unstated': { numbers: '-5, 12.5, 89', count: 3 },
  'text-orphan': { lines: 2, word: '"cc"', chars: 25 },
  'text-empty': { lines: 1, sampleLines: 5, chars: 45 },
  'row-uneven': { lines: 1, beside: 'card1Text', besideLines: 4, chars: 66 },
  'seen-in-picture': { problem: 'The text runs under the bottom edge; shorten it by a third.' },
  'accent2-missing': { slide: 1 },
  'color-near-page': { hex: 'F7F5EE' },
  'part-unwritten': {},
  'photos-over': { most: 6 },
  'parts-left-out-empty': { slide: 3, count: 2 },
  'parts-left-out-unusable': { slide: 3, count: 1 },
  'photos-left-out': { slide: 1, count: 1, most: 6 },
  'parts-unchanged': { slide: 2, count: 1 },
  'parts-unwritten': { slide: 3, count: 2 },
  'written-under-retry': { chars: 1500, words: 500, least: 2850, most: 3100 },
  'written-under': { words: 500, hasWords: 380 },
  'words-cut': { words: 673, holdsWords: 447, freeOff: true },
  'words-cut-dense': { words: 1000, holdsWords: 446, size: 7 },
  'photo-left-out': { words: 500, asked: true },
  'words-kept-low-retry': { percent: 60, label: 'p3-p5' },
  'words-kept-low': { percent: 72, label: 'p3' },
  'text-cut': { label: 'p3-p5', chars: 4100, holds: 3200, freeOff: true },
  'paragraphs-joined': { label: 'p1-p27', paragraphs: 53, holdsParagraphs: 25 },
  'text-cut-dense': { label: 'p1-p14', chars: 3727, holds: 2480, size: 7 },
  'nothing-changed': { slide: 3 },
  'recolor-slide-removed': { slide: 2 },
  'change-would-cut': { slide: 10, words: 1500, holdsWords: 446, size: 7, freeOff: false },
  'slide-holds-less': { slide: 3, words: 673, holdsWords: 447, size: 7, freeOff: false },
  'slide-set-small': { slide: 2, size: 10, templateSize: 18 },
  'text-retold': { name: 'pasted text 1' },
};

/** Other forms of some sentences: one of a thing, a clause left out or put in. */
const VARIANTS: Array<[IssueCode, Record<string, string | number | boolean>]> = [
  ['text-too-wide', { approximate: false, percent: 5, shrunk: true, size: 9.5, fitsChars: 1 }],
  ['text-overflow', { approximate: true, lines: 1, shrunk: false, size: 12, capacity: 0, over: 1, fitsChars: 1 }],
  ['colors-not-object', {}],
  ['color-not-hex', { slide: 2, name: 'accent2', given: '12' }],
  ['colors-replace-not-map', { slide: 5 }],
  ['colors-unknown-setting', { slide: 1, name: 'צבע' }],
  ['columns-relaxed', { slide: 1, asked: 1, columns: 2 }],
  ['free-text-small', { slide: 1, blocks: 1, size: 6 }],
  ['free-larger-text', { slide: 1, blocks: 3, half: false }],
  ['free-charts-bare', { slide: 1, blocks: 3, charts: 1 }],
  ['free-blocks-low', { slide: 1, blocks: 3, count: 2 }],
  ['table-over-rows', { rows: 1, most: 0 }],
  ['text-overflow', { approximate: false, lines: 3, shrunk: false, size: 19, capacity: 1, over: 2, fitsChars: 27 }],
  ['slot-over-chars', { chars: 2, most: 1 }],
  ['slot-over-paragraphs', { paragraphs: 2, most: 1 }],
  ['table-over-columns', { columns: 2, most: 1 }],
  ['chart-over-categories', { categories: 3, most: 1 }],
  ['chart-over-series', { series: 2, most: 1 }],
  ['paragraphs-joined', { label: 'p1-p4', paragraphs: 56, holdsParagraphs: 1 }],
  ['free-base-refused', { slide: 4, base: 'closing', why: `the layout "closing" has its title low on the slide: blocks would get 4% of the slide's height, and "cover" leaves them 30%`, stands: 'cover' }],
  ['free-base-refused', { slide: 2, base: 'agenda', why: 'the layout "agenda" has a row of items, which the software lays out itself', stands: 'text-only' }],
  ['free-base-refused', { slide: 2, base: 'quote', why: 'the layout "quote" has no title', stands: 'text-only' }],
  ['free-base-refused', { slide: 2, base: 'made', why: 'the layout "made" gives no frame of its own', stands: 'text-only' }],
  ['free-base-refused', { slide: 2, base: 'plain', why: 'the title of the layout "plain" takes its place from the slide layout, and neither that nor the master states one', stands: 'text-only' }],
  ['text-empty', { lines: 2, capacity: 9, chars: 300 }],
  ['text-empty', { lines: 1, sampleLines: 5, chars: 45 }],
  ['accent2-missing', {}],
  ['parts-left-out-empty', { slide: 1, count: 1 }],
  ['parts-left-out-unusable', { slide: 1, count: 3 }],
  ['photos-left-out', { slide: 2, count: 4, most: 6 }],
  ['parts-unchanged', { slide: 2, count: 3 }],
  ['parts-unwritten', { slide: 3, count: 1 }],
  ['words-cut', { words: 200, holdsWords: 92, freeOff: false }],
  ['photo-left-out', { words: 300, asked: false }],
  ['text-cut', { label: 'p3', chars: 615, holds: 297, freeOff: false }],
  ['change-would-cut', { slide: 2, words: 1000, holdsWords: 76, freeOff: true }],
  ['slide-holds-less', { slide: 1, words: 1000, holdsWords: 68, freeOff: true }],
  ['text-retold', { name: 'תוכנית 2027.docx' }],
];

const ISOLATE = (text: string) => `\u2068${text}\u2069`;
/** Latin words a Hebrew sentence may hold outside an isolate: the names of settings, and the colours of their examples. */
const LATIN_TERMS = ['colors', 'accent2', 'accent', 'replace', 'items', 'number', 'percent', '1565C0', '0E7C7B'];

function checkSentences(code: IssueCode, values: Record<string, string | number | boolean>): void {
  const en = (EN[code] as (values: unknown) => string)(values);
  const he = (HE[code] as (values: unknown) => string)(values);
  const what = `${code} ${JSON.stringify(values)}:\n${en}\n${he}`;
  for (const line of [en, he]) assert.doesNotMatch(line, /undefined|NaN|\[object|\$\{|\bnull\b/, what);
  // Every number the English states, the Hebrew states, as digits; a count of one may be the word
  // that follows its noun ("שורה אחת").
  const stated = new Set([...he.matchAll(/\d+(?:\.\d+)?/g)].map(([number]) => number));
  if (/(^|\s)(אחד|אחת)(?=[\s,.:;)]|$)/.test(he)) stated.add('1');
  // One of a thing is never counted with a plural noun.
  assert.doesNotMatch(he, /(^|[\s-])1 (שורות|חלקים|תווים|מילים|פסקאות|עמודות|קטגוריות|סדרות|תמונות|תרשימים)/, what);
  for (const [number] of en.matchAll(/\d+(?:\.\d+)?/g)) assert.ok(stated.has(number), `${number} is not in the Hebrew of ${what}`);
  // Every label or name stands in the Hebrew as given, in an isolate of its own, and in the English as given.
  for (const [key, value] of Object.entries(values)) {
    if (typeof value !== 'string') continue;
    assert.ok(en.includes(value), `${value} is not in the English of ${what}`);
    // A reason a layout cannot carry a free slide is said in Hebrew where it is known (its numbers are checked above).
    if (code === 'free-base-refused' && key === 'why') continue;
    // The slot beside a text in a row is named in Hebrew where Hebrew lines name such slots.
    if (code === 'row-uneven' && key === 'beside' && he.includes(`ה${hebrewSlot(value)}`)) continue;
    assert.ok(he.includes(ISOLATE(value)) || (code === 'colors-replace-not-pair' && he.includes(ISOLATE(`${values.from}: ${values.to}`))), `${value} is not isolated in the Hebrew of ${what}`);
  }
  // The Hebrew is Hebrew: it begins with a Hebrew word, so that a line set to the direction of its
  // first letter reads from the right, and holds no Latin word outside an isolate but the names of settings.
  assert.match(he, /^[א-ת]/, what);
  let bare = he.replace(/\u2068[^\u2069]*\u2069/g, '');
  for (const term of LATIN_TERMS) bare = bare.split(term).join('');
  assert.doesNotMatch(bare, /[A-Za-z]/, `a Latin word outside an isolate in ${what}`);
  // No imperative: what a person may do is said with "אפשר", what they should with "כדאי".
  assert.doesNotMatch(he, /(^|[\s:;])(בקש|בקשי|שים|שימי|כתוב|כתבי|הוסף|הוסיפי|בדוק|בדקי|קצרי)(\s|$)/, what);
}

test('every code has an English and a Hebrew sentence and a line in the API reference, and is named once', () => {
  assert.equal(new Set(ISSUE_CODES).size, ISSUE_CODES.length);
  for (const code of ISSUE_CODES) assert.match(code, /^[a-z][a-z0-9]*(?:-[a-z0-9]+)*$/, code);
  assert.deepEqual(Object.keys(EN), [...ISSUE_CODES]);
  assert.deepEqual(Object.keys(HE), [...ISSUE_CODES]);
  assert.deepEqual(Object.keys(SAMPLES), [...ISSUE_CODES]);
  const reference = readFileSync(join(root, 'docs', 'api-reference.md'), 'utf8');
  for (const code of ISSUE_CODES) assert.match(reference, new RegExp(`^\\| \`${code}\` \\|`, 'm'), `${code} has no line in docs/api-reference.md`);
});

test('every sentence states its values in English and in Hebrew, and the Hebrew is Hebrew', () => {
  for (const code of ISSUE_CODES) checkSentences(code, SAMPLES[code] as Record<string, string | number | boolean>);
  for (const [code, values] of VARIANTS) checkSentences(code, values);
  // The sentences of the production report.
  assert.equal(
    HE['text-cut-dense']({ label: 'p1-p14', chars: 3727, holds: 2480, size: 7 }),
    "הטקסט מהמסמך לשקף הזה (קטעים \u2068p1-p14\u2069) באורך 3727 תווים, והשקף מכיל כ-2480 תווים גם בגופן קטן של 7 נק': הטקסט נחתך כדי שייכנס. אפשר לבקש יותר שקפים כדי שהטקסט יישאר שלם",
  );
  assert.equal(HE['slide-holds-less']({ slide: 3, words: 673, holdsWords: 447, size: 7, freeOff: false }), "שקף 3 מכיל כ-447 מתוך 673 המילים שלו (בגופן 7 נק'): יתר הטקסט ייחתך");
  // The change was not made: what the slide would have held is said as such.
  assert.equal(HE['change-would-cut']({ slide: 10, words: 1500, holdsWords: 446, size: 7, freeOff: false }), "שקף 10 היה מכיל רק כ-446 מתוך 1500 המילים שלו (בגופן 7 נק'), ולכן השינוי לא בוצע");
  assert.equal(
    HE['change-would-cut']({ slide: 2, words: 1000, holdsWords: 76, freeOff: true }),
    'שקף 2 היה מכיל רק כ-76 מתוך 1000 המילים שלו, ולכן השינוי לא בוצע; שקפים חופשיים אינם פעילים, כך שהטקסט לא מוצג בגופן קטן יותר',
  );
  // What a box has room for is said with its noun, and one of a thing with the word.
  assert.equal(
    HE['text-overflow']({ approximate: false, lines: 3, shrunk: false, size: 19, capacity: 1, over: 2, fitsChars: 27 }),
    "הטקסט תופס 3 שורות ב-19 נק'; בתיבה יש מקום לשורה אחת (2 שורות יותר מדי). נכנס בה טקסט של כ-27 תווים.",
  );
  assert.equal(HE['slot-over-chars']({ chars: 230, most: 211 }), 'הטקסט באורך 230 תווים, ובתיבה יש מקום לכ-211 תווים');
  assert.equal(HE['table-over-rows']({ rows: 12, most: 8 }), 'בטבלה יש 12 שורות, ובתיבה יש מקום ל-8 שורות');
  assert.equal(
    HE['paragraphs-joined']({ label: 'p1-p27', paragraphs: 53, holdsParagraphs: 25 }),
    `בטקסט מהמסמך לשקף הזה (קטעים ${ISOLATE('p1-p27')}) יש 53 פסקאות, והשקף מציג אותו ב-25 פסקאות בלבד: חלק מהפסקאות חוברו זו לזו. אפשר לבקש יותר שקפים כדי שהפסקאות יישארו נפרדות`,
  );
  assert.equal(HE['parts-left-out-empty']({ slide: 1, count: 1 }), 'שקף 1: חלק אחד לא הכיל תוכן והושמט');
  assert.equal(
    HE['text-empty']({ lines: 1, sampleLines: 5, chars: 45 }),
    'הטקסט תופס שורה אחת, ואילו הטקסט של התבנית תופס 5 שורות, ולכן התיבה תיראה ריקה. טקסט של כ-45 תווים ימלא אותה כפי שתוכננה.',
  );
  // The reasons a layout cannot carry a free slide are said in Hebrew; one that is not known stays in English, isolated.
  assert.equal(
    HE['free-base-refused'](SAMPLES['free-base-refused']),
    `שקף 4 תוכנן על המסגרת של הפריסה "${ISOLATE('closing')}", שאי אפשר לבנות עליה שקף חופשי (הכותרת שלה נמוכה בשקף: לחלקים היו נשארים 25% מגובה השקף, ומסגרת משאירה להם 45% או יותר); לכן הוא בנוי על המסגרת של הפריסה "${ISOLATE('section')}"`,
  );
  const unknown = 'the frame is made of cheese';
  assert.ok(HE['free-base-refused']({ slide: 1, base: 'x', why: unknown, stands: 'y' }).includes(`(${ISOLATE(unknown)})`));
  // A negative number keeps its sign in front of it inside a right-to-left line: the list is isolated.
  assert.ok(HE['chart-values-unstated']({ numbers: '-5, 12.5', count: 2 }).includes('\u2068-5, 12.5\u2069'));
});

test('formatIssue: the English line is the line the tools write, and the Hebrew line names the slide and part in Hebrew', () => {
  const deck: DeckSpec = deckOf(
    { pattern: 'text-only', slots: { title: 'A title', body: 'A text.' } },
    freeSlide({ title: 'Ways in', blocks: [{ text: 'A text.' }, { card: { heading: 'Longer hours', text: 'Bikes unlock from five.' } }, { chart: { type: 'column', categories: ['A'], series: [{ name: 'S', values: [1] }] } }] }),
  );
  const over = (slide: number | undefined, slot?: string): Issue => ({ severity: 'warning', ...(slide !== undefined ? { slide } : {}), ...(slot ? { slot } : {}), message: EN['slot-over-chars']({ chars: 230, most: 211 }), code: 'slot-over-chars', values: { chars: 230, most: 211 } });
  const english = 'text is 230 characters; the slot fits about 211';
  const hebrew = 'הטקסט באורך 230 תווים, ובתיבה יש מקום לכ-211 תווים';
  const lines: Array<[Issue, string, string]> = [
    [over(1, 'body'), `slide 1 body: ${english}`, `שקף 1, גוף הטקסט: ${hebrew}`],
    [over(1, 'title'), `slide 1 title: ${english}`, `שקף 1, כותרת: ${hebrew}`],
    [over(1), `slide 1: ${english}`, `שקף 1: ${hebrew}`],
    [over(undefined), english, hebrew],
    // A part of a free slide is named by its number and kind; its slot within it where it has more than one.
    [over(2, 'b2Heading'), `slide 2, part 2 (card) heading: ${english}`, `שקף 2, חלק 2 (כרטיס), כותרת: ${hebrew}`],
    [over(2, 'b2Text'), `slide 2, part 2 (card) text: ${english}`, `שקף 2, חלק 2 (כרטיס), טקסט: ${hebrew}`],
    [over(2, 'b1Text'), `slide 2, part 1 (text): ${english}`, `שקף 2, חלק 1 (טקסט): ${hebrew}`],
    [over(2, 'b3Chart'), `slide 2, part 3 (chart): ${english}`, `שקף 2, חלק 3 (תרשים): ${hebrew}`],
    [over(2, 'title'), `slide 2 title: ${english}`, `שקף 2, כותרת: ${hebrew}`],
    // Slots most templates name alike are named in Hebrew; any other keeps its name, isolated.
    [over(1, 'subtitle'), `slide 1 subtitle: ${english}`, `שקף 1, כותרת משנה: ${hebrew}`],
    [over(1, 'card2Heading'), `slide 1 card2Heading: ${english}`, `שקף 1, כותרת של כרטיס 2: ${hebrew}`],
    [over(1, 'col1Text'), `slide 1 col1Text: ${english}`, `שקף 1, טקסט של עמודה 1: ${hebrew}`],
    [over(1, 'chartDescription'), `slide 1 chartDescription: ${english}`, `שקף 1, ${ISOLATE('chartDescription')}: ${hebrew}`],
  ];
  for (const [issue, en, he] of lines) {
    assert.equal(formatIssue(issue, 'en', { deck }), en);
    assert.equal(formatIssue(issue, 'he', { deck }), he);
    assert.equal(formatIssue(issue, 'he-IL', { deck }), he);
    assert.equal(formatIssue(issue, 'en'), english, 'without the deck, the sentence alone');
    assert.equal(formatIssue(issue, 'he'), hebrew);
  }
  // Any other language, an issue without a code, and values that cannot make the sentence, give the English.
  assert.equal(formatIssue(over(1, 'body'), 'fr', { deck }), `slide 1 body: ${english}`);
  const uncoded: Issue = { severity: 'warning', slide: 1, slot: 'body', message: 'seen by the app' };
  assert.equal(formatIssue(uncoded, 'he', { deck }), 'slide 1 body: seen by the app');
  assert.equal(formatIssue({ ...over(1), values: undefined }, 'he'), english);
  assert.equal(formatIssue({ ...over(1), code: 'no-such-code' as IssueCode }, 'he'), english);
});

/** A model that replays prepared answers, and records what it was asked. */
function scripted(...answers: unknown[]): { llm: LlmProvider; requests: JsonRequest[] } {
  const requests: JsonRequest[] = [];
  return {
    requests,
    llm: {
      async generateJson(request) {
        requests.push(request);
        if (answers.length === 0) throw new Error('the scripted model has no answer left');
        return answers.shift();
      },
    },
  };
}

/** What every result of the tools must hold: an issue for each line of `warnings`, in its order, making that line, and the language line in `next`. */
function checkResult(result: ToolResult, notes = 0): void {
  const warnings = (result.output.warnings as string[] | undefined) ?? [];
  const issues = result.issues ?? [];
  assert.equal(issues.length, warnings.length + notes, JSON.stringify(result.output));
  assert.deepEqual(
    issues.slice(0, warnings.length).map((issue) => formatIssue(issue, 'en', { deck: result.deck })),
    warnings,
  );
  for (const issue of issues) assertCoded(issue);
  const told = result.output.note !== undefined || result.output.warnings !== undefined || result.output.not_on_slides !== undefined;
  assert.equal(String(result.output.next ?? '').endsWith('in your own words and with the same slide numbers and figures; never paste its sentences as they are.'), told, String(result.output.next));
  if (!told) assert.equal(result.issues, undefined);
}

const asked = (words: number) => ({ pattern: 'text-only', purpose: 'Keep the bike share', words, shows: [], shows_main: 'auto', side_by_side: 0 });
const written = (chars: number) => ({ s1: { title: 'Keep it', b1Text: Array.from({ length: 12 }, () => 'Riders rely on the service every day. '.repeat(Math.ceil(chars / 12 / 37)).slice(0, Math.floor(chars / 12))) } });

test('the tools give the app every warning and note as an issue, and tell the chat model to speak the language of the user', async () => {
  // 1000 words on one slide: the plan says the cut, the presentation warns of it, in the same order.
  const outline = { title: 'Bikes', slides: [asked(1000)] };
  const model = scripted(outline, written(4350));
  let next = 0;
  const tools = presentationTools({ template: demo, llm: model.llm, store: memoryStore(), fields, layout: 'adaptive', colors: true, newId: (kind) => `${kind}_${++next}` });
  const planned = await tools.call(TOOL_NAMES.plan, { topic: 'The bike share', language: 'en', slide_count: 1, accent_color: 'FFFFFF' });
  checkResult(planned, 1);
  assert.deepEqual(planned.issues?.map((issue) => issue.code), ['slide-holds-less']);
  assert.equal(formatIssue(planned.issues?.[0] as Issue, 'he'), "שקף 1 מכיל כ-725 מתוך 1000 המילים שלו (בגופן 9 נק'): יתר הטקסט ייחתך");
  const created = await tools.call(TOOL_NAMES.create, { plan_id: 'plan_1' });
  checkResult(created);
  // The deck's cut, then what the render says of the whole deck: a colour as light as the page.
  assert.deepEqual(created.issues?.map((issue) => issue.code), ['words-cut-dense', 'color-near-page']);
  assert.deepEqual(
    created.issues?.map((issue) => formatIssue(issue, 'he', { deck: created.deck })),
    [
      "שקף 1: לשקף הזה התבקשו 1000 מילים, והוא מכיל כ-725 מהן גם בגופן קטן של 9 נק': יתר הטקסט נחתך. אפשר לבקש את הטקסט על פני יותר שקפים, או טקסט קצר יותר",
      'הצבע \u2068FFFFFF\u2069 קרוב לצבע הרקע של השקף; מילויים, עמודות של תרשימים וסמלים בצבע הזה כמעט לא ייראו',
    ],
  );

  // A plan without anything to tell says nothing, as before.
  const quiet = presentationTools({ template: demo, llm: scripted({ title: 'Bikes', slides: [{ pattern: 'text-only', purpose: 'Keep the bike share' }] }).llm, store: memoryStore(), fields, layout: 'adaptive' });
  checkResult(await quiet.call(TOOL_NAMES.plan, { topic: 'The bike share', language: 'en', slide_count: 1 }));

  // Free slides off: the note and the warning say so, and the issues carry it.
  const off = presentationTools({ template: demo, llm: scripted({ title: 'Bikes', slides: [asked(200)] }, { s1: { title: 'Keep it', body: ['Riders rely on it.'] } }).llm, store: memoryStore(), fields, layout: 'adaptive', free: false, newId: (kind) => `${kind}_${++next}` });
  const small = await off.call(TOOL_NAMES.plan, { topic: 'The bike share', language: 'en', slide_count: 1 });
  checkResult(small, 1);
  assert.deepEqual(small.issues?.map((issue) => [issue.code, issue.values?.freeOff]), [['slide-holds-less', true]]);
  const cut = await off.call(TOOL_NAMES.create, { plan_id: String(small.output.plan_id) });
  checkResult(cut);
  assert.deepEqual(cut.issues?.map((issue) => [issue.code, issue.values?.freeOff]), [['words-cut', true]]);
});

/** A metric in which every character is half an em wide: text measures the same on every machine. */
const fixedWidth: FontMetrics = { face: () => ({ has: () => true, advances: (text) => [...text].map(() => 0.5) }) };
const words = (count: number) => Array.from({ length: count }, () => 'word').join(' ');

test('every warning of the deck check, the render and the design check has its code', () => {
  const found: Issue[] = [];
  const keep = (issues: readonly Issue[]) => {
    found.push(...issues.filter((issue) => issue.severity === 'warning'));
    return issues;
  };
  const chart = { type: 'column', categories: Array.from({ length: 10 }, (_, at) => `Q${at + 1}`), series: Array.from({ length: 4 }, (_, at) => ({ name: `S${at}`, values: Array.from({ length: 10 }, () => at + 1) })), format: 'currency' };
  const table = { header: ['A', 'B', 'C', 'D', 'E'], rows: Array.from({ length: 6 }, () => ['A cell far longer than a cell holds', 'b', 'c', 'd', 'e']) };
  const busy: DeckSpec = {
    ...deckOf(
      { pattern: 'text-only', items: 2, slots: { title: 'A title far, far longer than the title of this slide can ever hold in its box, by a long way', body: ['One.', 'Two.', 'Three.', 'Four.'], nonsense: 'x' }, colors: 'red' as never },
      { pattern: 'table', slots: { title: 'A table', table }, colors: { replace: 'x' } as never },
      { pattern: 'chart-native', slots: { title: 'A chart', chart: chart as never, chartDescription: 'Rides.' }, colors: { accent2: 'C8102E' } },
      freeSlide({ title: 'Ways in', blocks: Array.from({ length: 4 }, (_, at) => ({ card: { heading: `Point ${at + 1}`, text: 'Bikes unlock from five.' } })), columns: 2.5 as never }),
      { ...freeSlide({ title: 'Ways in', blocks: [{ text: 'A text.' }] }), base: 'gone' },
      { ...freeSlide({ title: 'Ways in', blocks: [{ text: 'A text.' }] }), base: 'closing' },
      freeSlide({ title: 'Ways in', blocks: Array.from({ length: 12 }, (_, at) => ({ card: { heading: `Point ${at + 1}`, text: 'Bikes unlock from five.' } })), columns: 12 }),
      freeSlide({ title: 'Rides', blocks: Array.from({ length: 40 }, () => ({ chart: { type: 'line', categories: ['A', 'B'], series: [{ name: 'S', values: [1, 2] }] } })) }),
    ),
    colors: { accent: 'blue', accent2: '#12', replace: { '0E7C7B': 'teal' }, background: '000000' } as never,
  };
  keep(validateDeck(manifest, busy, { layout: 'adaptive' }));
  keep(validateDeck(demo, busy, { layout: 'adaptive' }));
  keep(validateDeck({ ...manifest, colors: false }, deckOf({ pattern: 'section', slots: { title: 'Part one' }, colors: { accent: '1565C0' } }), { layout: 'adaptive' }));
  // The render says what the file alone can tell: a colour close to the page, a second colour the template has not.
  keep(renderDeck(demo, { ...deckOf({ pattern: 'text-only', slots: { title: 'A title', body: 'A text.' } }), colors: { accent: 'FFFFFF' } }, { layout: 'adaptive' }).warnings);
  const single = loadTemplate(demoPptx, { ...manifest, patterns: manifest.patterns.filter((pattern) => ['text-only', 'stat-cards', 'icon-cards', 'two-columns'].includes(pattern.id)) });
  keep(renderDeck(single, { ...deckOf({ pattern: 'text-only', slots: { title: 'A title', body: 'A text.' }, colors: { accent2: 'C8102E' } }), colors: { accent2: 'C8102E' } }, { layout: 'adaptive' }).warnings);
  // Measured with fonts: text over its box, a box left empty, a row uneven, a heading that ends in a word.
  const measured = loadTemplate(demoPptx, manifest, { fonts: fixedWidth });
  const cards = (texts: string[]): DeckSpec['slides'][number] => ({
    pattern: 'icon-cards',
    items: texts.length,
    slots: { title: 'A title', ...Object.fromEntries(texts.flatMap((text, index) => [[`card${index + 1}Heading`, 'Heading'], [`card${index + 1}Text`, text], [`card${index + 1}Icon`, { asset: 'icon' }]])) },
  });
  const design = deckOf(
    { pattern: 'text-only', slots: { title: 'A title', subtitle: 'A subtitle', body: 'a'.repeat(4000) } },
    { pattern: 'text-and-photo', slots: { title: 'A title', subtitle: 'A subtitle', body: 'Two words', photo: { asset: 'photo' } } },
    cards([words(20), 'word', words(16)]),
    { pattern: 'text-and-photo', slots: { title: 'aaaaaaaaaaaa bbbbbbbbbbbb cc', subtitle: 'A subtitle', body: [words(30), words(14)], photo: { asset: 'photo' } } },
  );
  keep(checkDesign(measured, design));
  keep(validateDeck(measured, design, { layout: 'adaptive' }));
  // A box that does not wrap, and numbers of a chart the brief does not state.
  const wide = overflowTold({ fits: false, overflow: 0, widthRatio: 1.3, scale: 0.9, size: 12, fitsChars: 41, approximate: true } as TextFit);
  if (wide) found.push({ severity: 'warning', slide: 1, slot: 'title', ...wide });
  keep(unstatedChartValues(manifest, { ...deckOf({ pattern: 'chart-native', slots: { title: 'Rides', chart: { type: 'column', categories: ['A', 'B'], series: [{ name: 'S', values: [-5, 12.5] }] } } }), brief: 'Rides grew.' }));

  for (const issue of found) assertCoded(issue);
  const codes = new Set(found.map((issue) => issue.code));
  for (const code of [
    'colors-not-object',
    'color-not-hex',
    'colors-replace-not-map',
    'colors-replace-not-pair',
    'colors-unknown-setting',
    'colors-kept-by-template',
    'columns-not-whole',
    'columns-relaxed',
    'free-charts-bare',
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
    'text-overflow',
    'text-too-wide',
    'text-empty',
    'row-uneven',
    'text-orphan',
    'accent2-missing',
    'color-near-page',
  ] satisfies IssueCode[]) {
    assert.ok(codes.has(code), `no scenario gave ${code}: ${[...codes].join(', ')}`);
  }
  // A colour problem of a slide names the slide in its sentence and its values, and leaves the issue's slide unset.
  const slideColor = found.find((issue) => issue.code === 'colors-not-object');
  assert.deepEqual([slideColor?.slide, slideColor?.values, slideColor?.message], [undefined, { slide: 1 }, 'slide 1: "colors" must be an object such as { "accent": "1565C0" }; it is ignored']);
  assert.equal(formatIssue(slideColor as Issue, 'he'), 'שקף 1: ההגדרה "colors" צריכה להיות אובייקט כמו \u2068{ "accent": "1565C0" }\u2069; היא לא נלקחה בחשבון');
});

test('every warning the source builds is built with its code', () => {
  // A warning is made by spreading what `said` (or a function that calls it) returns into it. A
  // warning written with a message of its own would have no code: this finds one.
  const files = (dir: string): string[] =>
    readdirSync(dir, { withFileTypes: true }).flatMap((entry) => (entry.isDirectory() ? files(join(dir, entry.name)) : entry.name.endsWith('.ts') ? [join(dir, entry.name)] : []));
  // The template check speaks to template authors, in its own type: it is not looked at here.
  let built = 0;
  for (const file of files(join(root, 'src')).filter((path) => !/[\\/]template[\\/]/.test(path))) {
    const lines = readFileSync(file, 'utf8').split('\n');
    lines.forEach((line, at) => {
      if (!/severity: 'warning'[,}]/.test(line) && !/severity: 'warning',?$/.test(line)) return;
      built++;
      const object = lines.slice(at, at + 4).join(' ');
      assert.match(object, /\.\.\.(said\(|over\b|told\b|problem\b|\(why \?)/, `${file}:${at + 1} builds a warning without a code: ${line.trim()}`);
    });
  }
  assert.ok(built >= 40, `${built} places that build a warning`);
});

test('values that cannot make a sentence give the English message, and a slot named like a built-in keeps its name', () => {
  // A stored or hand-made issue may carry a code with values left out: its line is the English one,
  // never "undefined".
  const broken: Issue = { severity: 'warning', slide: 1, message: 'the slide holds about 2900', code: 'text-cut-dense', values: { label: 'p1-p3' } };
  assert.equal(formatIssue(broken, 'he'), 'the slide holds about 2900');
  // A part name that is a key of every object is no name in the Hebrew table: it stays as it is.
  const deck = deckOf({ pattern: 'free', blocks: ['text'], slots: { b1Constructor: 'x' } } as DeckSpec['slides'][number]);
  const line = formatIssue({ severity: 'warning', slide: 1, slot: 'b1Constructor', message: 'text is 10 characters; the slot fits about 5', code: 'slot-over-chars', values: { chars: 10, most: 5 } }, 'he', { deck });
  assert.doesNotMatch(line, /function|native code/, line);
});
