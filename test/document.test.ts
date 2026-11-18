// A deck made from a document a person gives: its text is cut into numbered passages, the outline
// puts each on a slide, and the slides carry the document's own words. A slide that holds its
// passages has them set into its running text by code. No test calls a live model.
import assert from 'node:assert/strict';
import { readFileSync } from 'node:fs';
import { dirname, join } from 'node:path';
import { test } from 'node:test';
import { fileURLToPath } from 'node:url';
import { MOST_DOCUMENT_CHARS, PlanError, TextCutError, documentPassages, editPlan, loadTemplate, planDeck, planEdit, planOutline, planTextFit, writeDeck } from '../src/index.ts';
import type { Brief, DeckPlan, DeckSpec, DocumentPassage, JsonRequest, LlmProvider, TemplateManifest } from '../src/index.ts';
import { unstatedChartValues } from '../src/deck/numbers.ts';
import { readableSize } from '../src/free/plan.ts';
import { buildCatalog } from '../src/plan/catalog.ts';
import type { CatalogContext, UsablePattern } from '../src/plan/catalog.ts';
import {
  DOCUMENT_CHANGE_RULES,
  DOCUMENT_EDIT_RULE,
  documentKept,
  fillParagraphs,
  isFixedCount,
  needOf,
  packPassages,
  passageLabel,
  passageText,
  passageWords,
  purposeOf,
  readAskedSlides,
  readDocument,
  readLanguageTag,
  readPassageNumbers,
  sameLanguage,
  scriptOf,
  tokensOf,
} from '../src/plan/document.ts';
import { charsOfWords, documentRoom, withPassages } from '../src/plan/length.ts';
import { systemPrompt } from '../src/plan/write.ts';
import { buildFixturePptx, fixtureManifest } from './fixture.ts';
import { withoutCodes } from './issues.ts';

const root = join(dirname(fileURLToPath(import.meta.url)), '..');
const demoManifest = JSON.parse(readFileSync(join(root, 'templates', 'demo.manifest.json'), 'utf8')) as TemplateManifest;
const demo = loadTemplate(new Uint8Array(readFileSync(join(root, 'templates', 'demo.pptx'))), demoManifest);
const rtl = loadTemplate(buildFixturePptx(), fixtureManifest);
const fields = { organizationDepartment: 'Retail' };
const options = { layout: 'adaptive' as const, fields };
const rtlOptions = { layout: 'adaptive' as const };

/** A model that replays prepared answers, or answers worked out from the request, and records what it was asked. */
function scripted(...answers: unknown[]): { llm: LlmProvider; requests: JsonRequest[] } {
  const requests: JsonRequest[] = [];
  return {
    requests,
    llm: {
      async generateJson(request) {
        requests.push(request);
        if (answers.length === 0) throw new Error('the scripted model has no answer left');
        const next = answers.shift();
        return typeof next === 'function' ? (next as (request: JsonRequest) => unknown)(request) : next;
      },
    },
  };
}

type Schema = { properties: Record<string, Schema>; required?: string[]; description?: string; maxItems?: number; type?: string };
const schemaOf = (request: JsonRequest | undefined): Schema => request?.schema as unknown as Schema;
/** A writer that gives every slide it is asked for a title and what else it must have, and nothing of its text. */
const titles = (request: JsonRequest): unknown =>
  Object.fromEntries(
    Object.entries(schemaOf(request).properties).map(([key, slide]) => [
      key,
      Object.fromEntries((slide.required ?? []).map((name) => [name, slide.properties[name]?.type === 'array' ? ['Point'] : name === 'title' ? `Slide ${key}` : 'Line'])),
    ]),
  );

/** An outline entry as a model answers it where free slides are on offer. */
const entry = (pattern: string, purpose: string, passages: unknown[] = [], more: Record<string, unknown> = {}) => ({ pattern, purpose, passages, shows: [], shows_main: 'auto', side_by_side: 0, ...more });
const freeText = { shows: [{ kind: 'text', count: 1 }], shows_main: 'none' };
const outlineOf = (slides: unknown[], more: Record<string, unknown> = {}) => ({ title: 'Bikes', left_out: [], asked_slides: 0, slides, ...more });
/** The rules block of an outline request: its last block. */
const rulesOf = (request: JsonRequest | undefined): string => (request?.prompt ?? '').split('\n\n').at(-1) ?? '';

// --- the documents ---

const SENTENCES = [
  'The city opened its bike share in 2019 with forty stations and four hundred bikes.',
  'Riders took more trips every year, and the busiest month was always June.',
  'Students and commuters make up most members, while visitors buy day passes.',
  'Each station has room for twelve bikes and a screen that shows the nearest dock.',
  'The council pays for the docks, and a private firm runs the fleet on contract.',
  'Repairs take two days on average, and a van moves bikes from full stations at night.',
  'Electric bikes were added in 2022 and now make up a third of all trips.',
  'A yearly pass costs sixty euros, and a single ride costs one euro.',
  'Surveys show that nine in ten riders would recommend the service to a friend.',
  'The main complaint is that docks near the station are often full in the morning.',
  'Crime is low: fewer than twenty bikes a year are stolen or lost.',
  'The firm proposes forty new stations in the outer districts by 2027.',
];
/** A paragraph of the sentences from `from` on, as many as make at least `chars` characters. */
const paragraph = (from: number, chars: number, sentences: readonly string[] = SENTENCES): string => {
  let text = '';
  for (let at = from; text.length < chars; at++) text += `${text ? ' ' : ''}${sentences[at % sentences.length]}`;
  return text;
};
/** A report of four sections and nine paragraphs, about 3,000 characters. */
const REPORT = [
  'Background',
  paragraph(0, 300),
  paragraph(4, 300),
  '',
  'Riders',
  paragraph(2, 300),
  paragraph(6, 300),
  '',
  'Costs',
  paragraph(7, 300),
  paragraph(9, 300),
  '',
  'Plans',
  paragraph(11, 300),
  paragraph(1, 300),
  paragraph(5, 300),
].join('\n');

const HEBREW = [
  'הצוות עובד יחד כל שבוע כדי לשפר את השירות לכל מי שגר בעיר ובסביבתה.',
  'הלקוחות מקבלים מענה מהיר יותר בזכות המערכת החדשה שהותקנה בכל הסניפים.',
  'בשנה הבאה נפתח שלושה סניפים חדשים בצפון ונגייס עשרים עובדים נוספים.',
  'התקציב לשנה הבאה גדל בעשרה אחוזים ויוקדש בעיקר להדרכה ולמחשוב.',
  'הסקר האחרון מראה ששמונה מכל עשרה לקוחות מרוצים מהיחס שקיבלו.',
  'הסיכון העיקרי הוא מחסור בעובדים מנוסים בתקופת החגים העמוסה.',
];
/** A Hebrew document of nine passages in three sections, about `chars` characters. */
const hebrewDocument = (chars = 3250): string => {
  const each = Math.floor(chars / 9);
  const part = (at: number) => paragraph(at, each - 30, HEBREW);
  return ['רקע', part(0), part(1), part(2), '', 'תוכנית', part(3), part(4), part(5), '', 'סיכונים', part(1), part(3), part(5)].join('\n');
};

const brief: Brief = { prompt: 'A deck from the bike share review', language: 'en', fields, imageDensity: 'none', document: REPORT };
// About 3,000 characters: what two slides of the narrow fixture hold at 7 to 8 pt, each column
// keeping a line free.
const hebrewBrief: Brief = { prompt: 'מצגת מהדוח', language: 'he', imageDensity: 'none', document: hebrewDocument(3000) };

// --- the passages ---

const FIXTURE_EN = [
  '# Bike share review',
  '',
  'Background',
  paragraph(0, 300),
  '',
  'What riders said:',
  '- Bikes are easy to find near the station',
  '- The app works well',
  '- Some docks are often full',
  '',
  'Costs',
  '| Item | 2023 | 2024 |',
  '|---|---|---|',
  '| Bikes | 120 | 140 |',
  '| Staff | 300 | 320 |',
  '',
  'Next steps',
  'Add twenty stations.',
  'Open at six.',
  paragraph(3, 1200),
].join('\n');

test('documentPassages: paragraphs, sections, lists and tables', () => {
  const passages = documentPassages(FIXTURE_EN);
  assert.deepEqual(passages, JSON.parse(readFileSync(join(root, 'test', 'fixtures', 'passages-en.json'), 'utf8')));
  // What the fixture holds, said outright.
  assert.deepEqual(
    passages.map((passage) => [passage.number, passage.kind, passage.section ?? null, passage.continues ?? null, passage.leads ?? null, passage.paragraphs]),
    [
      [1, 'text', 'Bike share review / Background', null, null, 1],
      [2, 'text', 'Bike share review / Background', null, true, 1],
      [3, 'list', 'Bike share review / Background', null, null, 3],
      [4, 'table', 'Costs', null, null, 3],
      [5, 'text', 'Next steps', null, null, 2],
      [6, 'text', 'Next steps', null, null, 1],
      [7, 'text', 'Next steps', true, null, 1],
      [8, 'text', 'Next steps', true, null, 1],
    ],
  );
  // A list is counted without its markers, a table with its cells joined and without its rule of dashes.
  assert.equal(passages[2]?.chars, 'Bikes are easy to find near the station'.length + 'The app works well'.length + 'Some docks are often full'.length);
  assert.equal(passages[3]?.chars, 'Item | 2023 | 2024'.length + 'Bikes | 120 | 140'.length + 'Staff | 300 | 320'.length);
  assert.equal(passageText(passages, [4]), '[Costs]\nItem | 2023 | 2024\nBikes | 120 | 140\nStaff | 300 | 320');
  assert.equal(passageLabel([4, 5, 6, 9]), 'p4-p6, p9');
  assert.equal(passageLabel([2]), 'p2');
});

test('documentPassages: a long paragraph is cut at sentence ends', () => {
  // Fourteen sentences of 99 characters: 1,400 with the spaces between them, in three pieces of five, five and four.
  const sentence = (at: number) => `${`Point ${at + 10} of the review says that riders use the bikes every day and want more docks`.padEnd(98, '!')}.`;
  const long = Array.from({ length: 14 }, (_, at) => sentence(at)).join(' ');
  assert.equal(long.length, 1399);
  const passages = documentPassages(long);
  assert.deepEqual(passages.map((passage) => passage.chars).sort(), [399, 499, 499]);
  assert.deepEqual(
    passages.map((passage) => passage.continues ?? null),
    [null, true, true],
  );
  for (const passage of passages) assert.match(passage.text, /\.$/);
  assert.equal(passageText(passages, [1, 2, 3]), long);
  // Set by code, the pieces are one paragraph again.
  assert.deepEqual(fillParagraphs(passages, [1, 2, 3]), [long]);
});

test('documentPassages: a sentence over 500 stays whole, and e.g. ends no sentence', () => {
  const sentence = `${'The council and the firm met many times to talk about the docks and the bikes and the vans '.repeat(6).trim()}.`;
  assert.ok(sentence.length > 500);
  assert.equal(documentPassages(sentence).length, 1);
  // Cut near the middle, where "e.g." stands, the cut goes to a real end of a sentence.
  const first = `${'Riders want more docks near the station and longer hours in summer '.repeat(3).trim()}.`;
  const tricky = `${first} Some, e.g. students, ${'ride late at night and want the lights on every dock '.repeat(3).trim()}. ${'Mr. Smith of the firm said the vans can do more '.repeat(4).trim()}.`;
  const passages = documentPassages(tricky);
  assert.ok(passages.length >= 2, String(passages.length));
  for (const passage of passages) {
    assert.doesNotMatch(passage.text, /(?:e\.g|Mr)\.$/);
    assert.doesNotMatch(passage.text, /^(?:students|Smith)/);
  }
  assert.equal(passageText(passages, passages.map((passage) => passage.number)), tricky);
});

test('documentPassages: soft wraps join, short lines merge, Hebrew letter items, CRLF and a byte order mark', () => {
  const wrapped = '\uFEFFThe city opened its bike share in 2019 with forty stations and\r\nfour hundred bikes, and riders took more trips every year, the busiest month\r\nalways being June.\r\n\r\nOne.\r\nTwo.\r\nThree.\r\n\r\nא. ראשון\r\nב. שני';
  const passages = documentPassages(wrapped);
  assert.deepEqual(
    passages.map((passage) => [passage.kind, passage.text]),
    [
      ['text', 'The city opened its bike share in 2019 with forty stations and four hundred bikes, and riders took more trips every year, the busiest month always being June.'],
      ['text', 'One.\nTwo.\nThree.'],
      ['list', 'א. ראשון\nב. שני'],
    ],
  );
  assert.equal(passages[1]?.paragraphs, 3);
  assert.equal(passages[2]?.chars, 'ראשון'.length + 'שני'.length);
});

test('documentPassages: headings are sections, stacked headings join, a last heading is text', () => {
  const text = ['# Part one', '## Costs', paragraph(0, 200), '', 'Plans', paragraph(3, 200), '', 'The end'].join('\n');
  const passages = documentPassages(text);
  assert.deepEqual(
    passages.map((passage) => [passage.section ?? null, passage.text.slice(0, 12)]),
    [
      ['Part one / Costs', paragraph(0, 200).slice(0, 12)],
      ['Plans', paragraph(3, 200).slice(0, 12)],
      ['Plans', 'The end'],
    ],
  );
  // A heading starts a slide's text in brackets for the writer, and code sets one that changes inside a slide as a paragraph.
  assert.equal(passageText(passages, [1, 2]).split('\n')[0], '[Part one / Costs]');
  assert.deepEqual(fillParagraphs(passages, [1, 2]), [paragraph(0, 200), 'Plans', paragraph(3, 200)]);
});

test('documentPassages keeps the text as given', () => {
  const marked = `\u200F${hebrewDocument(1500).replace('רקע', 'רקע\u200F')}`;
  const squeeze = (text: string) => text.replace(/\s+/g, ' ');
  const copy = squeeze(marked);
  for (const passage of [...documentPassages(FIXTURE_EN), ...documentPassages(marked)]) {
    for (const line of passage.text.split('\n')) assert.ok(copy.includes(squeeze(line)) || squeeze(FIXTURE_EN).includes(squeeze(line)), line);
  }
  // The marks that set direction stay in the text, and are not counted.
  const passages = documentPassages(marked);
  assert.ok(passages.some((passage) => passage.text.includes('\u200F') || passage.section?.includes('\u200F')));
  assert.equal(passages[0]?.chars, passages[0]?.text.replace(/\u200F/g, '').length);
});

test('documentPassages is stable', () => {
  const he = readFileSync(join(root, 'test', 'fixtures', 'passages-he.json'), 'utf8');
  const { document, passages } = JSON.parse(he) as { document: string; passages: DocumentPassage[] };
  assert.deepEqual(documentPassages(document), passages);
  assert.deepEqual(documentPassages(document), documentPassages(document));
});

test('documentPassages: a tab before a paragraph sets it in, and a tab between filled cells makes a row', () => {
  // Pasted from a word processor: a tab at the start of each paragraph, and CRLF line ends.
  const pasted = ['Annual report', ...[0, 3, 6].map((at) => `\t${paragraph(at, 150)}`)].join('\r\n');
  const passages = documentPassages(pasted);
  assert.deepEqual(
    passages.map((passage) => [passage.kind, passage.section ?? null, passage.paragraphs]),
    [
      ['text', 'Annual report', 1],
      ['text', 'Annual report', 1],
      ['text', 'Annual report', 1],
    ],
  );
  assert.deepEqual(fillParagraphs(passages, [1, 2, 3]), [0, 3, 6].map((at) => paragraph(at, 150)));
  assert.equal(readDocument(pasted, 'en').words, tokensOf([0, 3, 6].map((at) => paragraph(at, 150)).join(' ')).length);
  // A word processor's list: a marker, then a tab. A tab after the text of an item makes a row.
  const list = documentPassages(['Steps', '•\tOpen the app', '•\tFind a dock', '', 'No.\tItem\tCost', '1.\tBikes\t120', '2.\tStaff\t300'].join('\r\n'));
  assert.deepEqual(
    list.map((passage) => [passage.kind, passage.paragraphs, passage.chars]),
    [
      ['list', 2, 'Open the app'.length + 'Find a dock'.length],
      ['table', 3, 'No. | Item | Cost'.length + '1. | Bikes | 120'.length + '2. | Staff | 300'.length],
    ],
  );
  assert.deepEqual(fillParagraphs(list, [1]), ['• Open the app', '• Find a dock']);
  // Rows on lines one after another are a table; a blank line ends one.
  const tables = documentPassages(['Costs', 'Item:\t2023', 'Bikes:\t120', '', 'Staff:\t300', 'Vans:\t40'].join('\n'));
  assert.deepEqual(
    tables.map((passage) => [passage.kind, passage.text]),
    [
      ['table', 'Item:\t2023\nBikes:\t120'],
      ['table', 'Staff:\t300\nVans:\t40'],
    ],
  );
  assert.equal(passageText(tables, [1]), '[Costs]\nItem: | 2023\nBikes: | 120');
  // A table pasted from a spreadsheet keeps its empty corner; tabs that set in every row are no cells.
  const corner = documentPassages(['Results', '\t2023\t2024', 'Bikes\t120\t140', '\tStaff\t300'].join('\r\n'));
  assert.equal(passageText(corner, [1]), '[Results]\n | 2023 | 2024\nBikes | 120 | 140\n | Staff | 300');
  const setIn = documentPassages(['\tName:\tAlice', '\tRole:\tLead', '\t\tShift:\tLate'].join('\n'));
  assert.equal(passageText(setIn, [1]), 'Name: | Alice\nRole: | Lead\n | Shift: | Late');
  // A word processor writes the bullets of a list within a list as "o", then a tab.
  const nested = documentPassages(['Goals', '•\tGrow sales', 'o\tIn the north', 'o\tIn the south', '•\tCut costs'].join('\r\n'));
  assert.deepEqual(
    nested.map((passage) => [passage.kind, passage.paragraphs]),
    [['list', 4]],
  );
  assert.deepEqual(fillParagraphs(nested, [1]), ['• Grow sales', '◦ In the north', '◦ In the south', '• Cut costs']);
  // An "o" before a space is a word, not a bullet.
  assert.equal(documentPassages('o meu caro amigo\no teu caro amigo')[0]?.kind, 'text');
});

test('documentPassages: a run of short lines with no stops is text, all of it', () => {
  // A list typed without bullets: three short lines or more with no stops, one under another.
  const goals = ['Our goals for next year', 'Grow revenue by ten percent in every region', 'Improve customer satisfaction in all branches', 'Open two new branches in the north', 'Hire forty people for the new branches'];
  const passages = documentPassages(goals.join('\n'));
  assert.deepEqual(
    passages.map((passage) => [passage.section ?? null, passage.text]),
    [[null, goals.join('\n')]],
  );
  assert.deepEqual(fillParagraphs(passages, [1]), goals);
  // Whatever is over it: a sentence, a line that ends with a colon. No line of it is a heading.
  for (const over of ['The board agreed on the plan for next year.', 'Our goals for next year:']) {
    const read = documentPassages([over, ...goals.slice(1)].join('\r\n'));
    assert.ok(
      read.every((passage) => passage.section === undefined),
      JSON.stringify(read),
    );
    assert.deepEqual(fillParagraphs(read, read.map((passage) => passage.number)), [over, ...goals.slice(1)]);
  }
  // A blank line under a short line makes it a heading over the run, and a short line over one other is one.
  const titled = documentPassages([goals[0], '', ...goals.slice(1)].join('\n'));
  assert.deepEqual(
    titled.map((passage) => [passage.section, passage.text]),
    [[goals[0], goals.slice(1).join('\n')]],
  );
  assert.deepEqual(
    documentPassages(['Revenue', 'Up twelve percent in the north', '', 'Costs', 'Down three percent'].join('\n')).map((passage) => [passage.section, passage.text]),
    [
      ['Revenue', 'Up twelve percent in the north'],
      ['Costs', 'Down three percent'],
    ],
  );
  // A short line under a sentence is still a heading.
  const sections = documentPassages(['Background', paragraph(0, 200), 'Riders', paragraph(3, 200)].join('\n'));
  assert.deepEqual(
    sections.map((passage) => passage.section),
    ['Background', 'Riders'],
  );
});

// --- planning ---

test('no document, or a blank one, gives the requests of a brief without one', async () => {
  const plain: Brief = { prompt: 'Why the city should keep its bike share', language: 'en', fields, imageDensity: 'none', slideCount: 3 };
  const outline = { title: 'Bikes', slides: [entry('cover', 'Bikes'), entry('text-only', 'Keep it'), entry('closing', 'Thanks')] };
  const run = async (more: Partial<Brief>) => {
    const model = scripted(outline, titles);
    const result = await planDeck(demo, { ...plain, ...more }, { llm: model.llm, ...options });
    const change = scripted({ summary: 'Done.', language: 'en', title: 'Bikes', slides: outline.slides });
    await editPlan(demo, { version: 1, title: 'Bikes', brief: { ...plain, ...more }, slides: result.outline }, 'Shorter titles', { llm: change.llm, ...options });
    const deck: DeckSpec = more.document === undefined ? result.deck : { ...result.deck, document: '  ' };
    const edit = scripted({ summary: 'Done.', language: 'en', title: 'Bikes', slides: [{ from: 2, pattern: 'text-only', change: 'A shorter title', images: false, parts: [], shows_main: 'auto', side_by_side: 0 }], text_size: 'normal', fields }, titles);
    await planEdit(demo, deck, 'A shorter title on slide 2', { llm: edit.llm, ...options });
    const requests = [...model.requests, ...change.requests, ...edit.requests].map(({ system, prompt, schema, maxOutputTokens, effort }) => ({ system, prompt, schema, maxOutputTokens, effort }));
    return { requests, deck: result.deck };
  };
  const before = await run({});
  assert.equal('document' in before.deck, false);
  for (const document of [undefined, '', ' \n ']) {
    const after = await run({ document });
    assert.deepEqual(after.requests, before.requests, JSON.stringify(document));
    assert.equal('document' in after.deck, false);
  }
});

test('a document that is not text or is too long is a PlanError', async () => {
  await assert.rejects(planOutline(demo, { ...brief, document: 42 as unknown as string }, { llm: scripted().llm, ...options }), new PlanError('"document" must be text'));
  const long = paragraph(0, MOST_DOCUMENT_CHARS + 100);
  await assert.rejects(
    planOutline(demo, { ...brief, document: long }, { llm: scripted().llm, ...options }),
    new PlanError(`The document is ${long.length} characters, and a deck is made from at most 12000 (about 1,800 words): pass the part the deck is about`),
  );
  assert.equal(MOST_DOCUMENT_CHARS, 12000);
});

test('the outline shows the passages after the brief, and the rules last', async () => {
  const model = scripted(outlineOf([entry('cover', 'Bikes'), entry('free', 'All of it', [1, 2, 3, 4, 5, 6, 7, 8, 9], freeText), entry('closing', 'Thanks')]));
  await planOutline(demo, brief, { llm: model.llm, ...options });
  const request = model.requests[0];
  const prompt = request?.prompt ?? '';
  const block = prompt.indexOf('The document to make the deck from, in 9 numbered passages (');
  assert.ok(prompt.indexOf('Brief:') < block && block < prompt.indexOf('Plan a deck of 1 to 10 slides'), prompt.slice(0, 400));
  assert.match(prompt, /\n\[Background\]\np1 \(\d+\): The city opened its bike share/);
  const rules = rulesOf(request);
  assert.ok(rules.includes('- The slides carry the document\'s own text. Set "passages" on each slide'));
  assert.ok(rules.includes('Such a passage is part of the brief: do what it asks (a number of slides, a summary, all of the text).'));
  assert.ok(rules.includes('- The document takes about 4 slides of its own: one holds about 850 characters of it at the template\'s text size.'));
  assert.ok(rules.includes('make exactly that many and set "asked_slides" to that number.'));
  assert.ok(rules.includes('- A heading in brackets is a good place to start a slide.'));
  assert.ok(rules.includes('- An opening, a closing or a section slide carries no passage: give it [].'));
  assert.ok(rules.includes('- When the brief asks for less text than the document has'));
  assert.ok(rules.includes('- A deck of four or more slides opens with an opening pattern and ends with a closing pattern when the list has them'));
  assert.ok(!rules.includes('A deck of three or more slides opens'));
  assert.ok(rules.includes("- Otherwise vary the patterns among those whose slots hold each slide's passages."));
  assert.ok(!rules.includes('- Otherwise vary the patterns across the deck.'));
  // The document's rules come after the rule on words.
  assert.ok(rules.indexOf('set "words" on that slide') < rules.indexOf('The slides carry the document'));
  assert.equal(request?.system, systemPrompt('en', 'brief', true));
  assert.match(request?.system ?? '', /A slide that has text from the document says that text, in its words and order: shorten it only where a limit makes you/);
  assert.match(request?.system ?? '', /- Every figure on a slide must come from the brief or the document\./);
  assert.match(request?.system ?? '', /- Follow the brief\. Do not add topics that neither the brief nor the document calls for\./);
  assert.match(systemPrompt('en', 'deck', true), /shorten it only where a limit or the request makes you/);
  assert.match(systemPrompt('en', 'deck', true), /- Every figure on a slide must come from the request, the document or the deck as it is\./);
});

test('the outline asks for passages, left_out, asked_slides and the language of the document', async () => {
  const model = scripted(outlineOf([entry('free', 'All of it', [1, 2, 3, 4, 5, 6, 7, 8, 9], freeText)]));
  await planOutline(demo, brief, { llm: model.llm, ...options });
  const schema = schemaOf(model.requests[0]);
  assert.deepEqual(schema.required, ['title', 'slides', 'left_out', 'asked_slides', 'document_language']);
  assert.deepEqual(schema.properties.document_language, { type: 'string', description: 'The BCP 47 tag of the language the document is written in ("en", "he"), whatever the language of the deck.' });
  const slide = (schema.properties.slides as unknown as { items: Schema }).items;
  assert.ok(slide.required?.includes('passages'));
  assert.deepEqual(slide.properties.passages, { type: 'array', items: { type: 'integer', minimum: 1, maximum: 9 }, description: 'The numbers of the passages of the document this slide carries, in order. [] for a slide that carries none.' });
  assert.equal(slide.properties.purpose?.description, 'The one point this slide makes: a single sentence of at most 25 words');
  // 400 + 160 a slide, 12 for a free slide's parts, 8 a slide for its passages, 2 a passage and 40.
  assert.equal(model.requests[0]?.maxOutputTokens, 400 + 160 * 10 + 12 * 10 + 8 * 10 + 2 * 9 + 40);

  // With a number of slides there is no asked_slides.
  const counted = scripted(outlineOf([entry('free', 'All of it', [1, 2, 3, 4, 5, 6, 7, 8, 9], freeText)]));
  await planOutline(demo, { ...brief, slideCount: 1 }, { llm: counted.llm, ...options });
  assert.deepEqual(schemaOf(counted.requests[0]).required, ['title', 'slides', 'left_out', 'document_language']);
});

test('the count comes from the document', async () => {
  // An answer of three slides or fewer with one that carries nothing is asked again, and stands the second time.
  const one = (more: Partial<Brief>, slides: unknown[]) => {
    const model = scripted(outlineOf(slides), outlineOf(slides));
    return planOutline(demo, { ...brief, ...more }, { llm: model.llm, ...options }).then(() => model.requests[0] as JsonRequest);
  };
  const all = [1, 2, 3, 4, 5, 6, 7, 8, 9];
  const open = await one({}, [entry('free', 'All', all, freeText)]);
  const slides = (request: JsonRequest) => schemaOf(request).properties.slides as unknown as { minItems: number; maxItems: number };
  assert.deepEqual([slides(open).minItems, slides(open).maxItems], [1, 10]);

  const five = await one({ slideCount: 5 }, [1, 2, 3, 4, 5].map((at) => entry('free', `Part ${at}`, at === 5 ? [9] : [2 * at - 1, 2 * at], freeText)));
  assert.deepEqual([slides(five).minItems, slides(five).maxItems], [5, 5]);
  assert.ok(rulesOf(five).includes("- The deck has 5 slides: share the passages over those that carry them, in the document's order. A slide that carries more than it holds at the template's size has its text set smaller"));
  assert.ok(!rulesOf(five).includes('asked_slides'));
  assert.equal('asked_slides' in schemaOf(five).properties, false);

  for (const count of [3, 2]) {
    const request = await one({ slideCount: count }, Array.from({ length: count }, (_, at) => entry('free', `Part ${at}`, at === 0 ? all : [], freeText)));
    assert.ok(rulesOf(request).includes('A deck of three slides or fewer carries the document on every slide'));
    assert.ok(!rulesOf(request).includes('A deck of three or more slides opens with an opening pattern'));
  }
  for (const request of [open, five]) assert.ok(!request.prompt.includes('"left_out", or every slide is cut short'));
});

test('a small gap is put right by code with no retry', async () => {
  const model = scripted(outlineOf([entry('cover', 'Bikes'), entry('free', 'Background', [1, 2, 3, 4], freeText), entry('free', 'Costs', [6, 7, 8, 9], freeText), entry('closing', 'Thanks')]));
  const plan = await planOutline(demo, { ...brief, slideCount: 4 }, { llm: model.llm, ...options });
  assert.equal(model.requests.length, 1);
  assert.deepEqual(
    plan.slides.map((slide) => slide.passages ?? []),
    [[], [1, 2, 3, 4, 5], [6, 7, 8, 9], []],
  );
});

test('a gross gap is asked again, then packed', async () => {
  const bare = outlineOf([entry('cover', 'Bikes'), entry('free', 'One', [], freeText), entry('free', 'Two', [], freeText), entry('closing', 'Thanks')], { left_out: [9] });
  const model = scripted(bare, bare);
  const plan = await planOutline(demo, { ...brief, slideCount: 4 }, { llm: model.llm, ...options });
  assert.equal(model.requests.length, 2);
  assert.ok(model.requests[1]?.prompt.endsWith('Your previous answer could not be used: passages p1-p8 are on no slide and not in "left_out". Answer again, with every passage on one slide or in "left_out".'), String(model.requests[1]?.prompt.slice(-300)));
  // The passages that are not left out are shared over the slides of content, each about its own.
  assert.deepEqual(plan.slides.map((slide) => slide.pattern).filter((pattern) => pattern !== 'free'), ['cover', 'closing']);
  const placed = plan.slides.flatMap((slide) => slide.passages ?? []);
  assert.deepEqual(placed, [1, 2, 3, 4, 5, 6, 7, 8]);
  assert.deepEqual(
    plan.slides.slice(1, 3).map((slide) => slide.purpose),
    plan.slides.slice(1, 3).map((slide) => readDocument(REPORT, 'en').passages[(slide.passages?.[0] ?? 1) - 1]?.section),
  );
});

test('a brief that only asks for the presentation places nothing', async () => {
  const asks = `Make three slides of this:\n\n${REPORT}`;
  const passages = documentPassages(asks);
  assert.equal(passages[0]?.leads, true);
  // Every passage left out: no retry, nothing packed, no passage on a slide.
  const model = scripted(outlineOf([entry('cover', 'Bikes'), entry('text-only', 'Bikes')], { left_out: passages.map((passage) => passage.number) }));
  const plan = await planOutline(demo, { ...brief, document: asks }, { llm: model.llm, ...options });
  assert.equal(model.requests.length, 1);
  assert.ok(plan.slides.every((slide) => slide.passages === undefined));
  // The line that asks is left out where the answer forgets it; the rest is placed.
  const forgotten = outlineOf([entry('cover', 'Bikes'), entry('free', 'All', passages.slice(1).map((passage) => passage.number), freeText)]);
  const forgot = scripted(forgotten, forgotten);
  const placed = await planOutline(demo, { ...brief, document: asks, slideCount: 2 }, { llm: forgot.llm, ...options });
  assert.deepEqual(placed.slides[1]?.passages, passages.slice(1).map((passage) => passage.number));
});

test('repeats stay on their first slide, numbers out of range are dropped, left_out is kept', async () => {
  const model = scripted(outlineOf([entry('free', 'One', [1, 2, '3', 3, 40, -1, 'x'], freeText), entry('free', 'Two', [3, 4, 5, 6, 7, 8], freeText)], { left_out: [9, 2] }));
  const plan = await planOutline(demo, { ...brief, slideCount: 2 }, { llm: model.llm, ...options });
  assert.deepEqual(
    plan.slides.map((slide) => slide.passages),
    [
      [1, 2, 3],
      [4, 5, 6, 7, 8],
    ],
  );
  assert.deepEqual(readPassageNumbers([3, '2', 'p1', 2.5, 0, 10, null], 9), [1, 2, 3]);
});

test('passages on an opening or a closing slide move to the slide next to it', async () => {
  const model = scripted(outlineOf([entry('cover', 'Bikes', [1, 2]), entry('free', 'Middle', [3, 4, 5, 6], freeText), entry('free', 'Late', [7], freeText), entry('closing', 'Thanks', [8, 9])]));
  const plan = await planOutline(demo, { ...brief, slideCount: 4 }, { llm: model.llm, ...options });
  assert.deepEqual(
    plan.slides.map((slide) => slide.passages ?? []),
    [[], [1, 2, 3, 4, 5, 6], [7, 8, 9], []],
  );
});

test("a slide's length comes from its passages, not from words", async () => {
  const doc = readDocument(REPORT, 'en');
  const words = (numbers: number[]) => numbers.reduce((sum, number) => sum + tokensOf(fillParagraphs(doc.passages, [number]).join(' ')).length, 0);
  const all = [1, 2, 3, 4, 5, 6, 7, 8, 9];
  // Words of four fifths of the passages' or more ask for no shorter text.
  const whole = scripted(outlineOf([entry('text-only', 'All of it', all, { words: Math.ceil(0.8 * words(all)) })]));
  const plan = await planOutline(demo, { ...brief, slideCount: 1 }, { llm: whole.llm, ...options });
  assert.equal(plan.slides[0]?.words, undefined);
  assert.deepEqual(plan.slides[0]?.passages, all);
  assert.deepEqual(planTextFit(demo, plan, options)[0]?.words, words(all));
  // Fewer words ask for a shorter text, and the slide is homed for them: 45 words fit the text layout.
  const rest = { left_out: [3, 4, 5, 6, 7, 8, 9] };
  const short = scripted(outlineOf([entry('text-only', 'Background', [1, 2], { words: 45 })], rest));
  const shortened = await planOutline(demo, { ...brief, slideCount: 1 }, { llm: short.llm, ...options });
  assert.deepEqual(shortened.slides, [{ pattern: 'text-only', purpose: 'Background', words: 45, passages: [1, 2] }]);
  assert.deepEqual(planTextFit(demo, shortened, options), [{ slide: 1, words: 45, holds: 45, passages: 'p1-p2' }]);
  // Without the words the passages are longer than the layout holds, and the slide becomes a free text.
  const unasked = scripted(outlineOf([entry('text-only', 'Background', [1, 2])], rest));
  const homed = await planOutline(demo, { ...brief, slideCount: 1 }, { llm: unasked.llm, ...options });
  assert.deepEqual(homed.slides, [{ pattern: 'free', purpose: 'Background', passages: [1, 2], blocks: ['text'], arrange: { dense: true } }]);
  assert.equal(charsOfWords(45, 'en'), 270);
});

test('three or more paragraphs fill a layout to three quarters, and more than its slot takes join with no room lost', () => {
  const context: CatalogContext = { layout: 'adaptive', fields, icons: [], assets: [], imageDensity: 'none', figures: false, textScale: 1 };
  const textOnly = buildCatalog(demoManifest, context).patterns.find((pattern) => pattern.def.id === 'text-only') as UsablePattern;
  assert.equal(documentRoom(textOnly, context, { paragraphs: 1 }), Math.floor(0.9 * 331));
  assert.equal(documentRoom(textOnly, context, { paragraphs: 3 }), Math.floor(0.75 * 331));
  // The slot takes 3 paragraphs: 6 are joined into 3, and lose no characters by it. What the slide
  // is said to hold is the room the writer is told.
  assert.equal(documentRoom(textOnly, context, { paragraphs: 6 }), Math.floor(0.75 * 331));
  const told = withPassages(textOnly, 600, context, { label: 'p1-p6', shape: { paragraphs: 6 } }).slots.find((planned) => planned.slot.name === 'body')?.slot.description;
  assert.match(String(told), /^The slide's text from the document \(p1-p6\) shortened to 220 to 240 characters, the most this slide holds/);
  // A slot that takes four paragraphs or more is told one fewer for such a text: every paragraph
  // ends on a line of its own, and its last ran past the box. Two paragraphs, and a slot of three, keep their limit.
  const list = buildCatalog(demoManifest, { ...context, imageDensity: 'medium' }).patterns.find((pattern) => pattern.def.id === 'list-with-photo') as UsablePattern;
  const most = (pattern: UsablePattern, paragraphs: number, name: string) => (withPassages(pattern, 300, context, { label: 'p1', shape: { paragraphs } }).slots.find((planned) => planned.slot.name === name)?.slot as { maxParagraphs?: number }).maxParagraphs;
  assert.deepEqual([most(list, 3, 'bulletList'), most(list, 2, 'bulletList'), most(textOnly, 6, 'body')], [3, 4, 3]);
});

test('a long slide spreads at a readable size, each part with its own passages', async () => {
  // No number of slides: one slide given all nine passages goes over as many as it takes.
  const model = scripted(outlineOf([entry('cover', 'Bikes'), entry('text-and-photo', 'The review', [1, 2, 3, 4, 5, 6, 7, 8, 9]), entry('closing', 'Thanks')]));
  const plan = await planOutline(demo, { ...brief, imageDensity: 'low' }, { llm: model.llm, ...options });
  const parts = plan.slides.slice(1, -1);
  assert.ok(parts.length >= 3, JSON.stringify(plan.slides));
  assert.deepEqual(
    parts.flatMap((slide) => slide.passages ?? []),
    [1, 2, 3, 4, 5, 6, 7, 8, 9],
  );
  for (const slide of parts) {
    const list = slide.passages ?? [];
    assert.deepEqual(list, Array.from({ length: list.length }, (_, at) => (list[0] as number) + at));
  }
  parts.forEach((slide, at) => assert.equal(slide.purpose, `The review (${at + 1}/${parts.length})`));
  // A later part is a slide of text: the photo is not shown twice.
  for (const slide of parts.slice(1)) assert.ok(slide.pattern === 'text-only' || (slide.pattern === 'free' && JSON.stringify(slide.blocks) === '["text"]'), JSON.stringify(slide));
  // Of parts of about even length, as many as can start where a section of the report starts do:
  // the report has four sections, so three of the parts after the first.
  const doc = readDocument(REPORT, 'en');
  const starts = parts.slice(1).filter((slide) => {
    const first = slide.passages?.[0] as number;
    return doc.passages[first - 1]?.section !== doc.passages[first - 2]?.section;
  });
  assert.equal(starts.length, Math.min(parts.length - 1, 3), JSON.stringify(parts.map((part) => part.passages)));
  for (const fit of planTextFit(demo, plan, options)) assert.equal(fit.holds, fit.words, JSON.stringify(fit));
});

test('a slide asked to be shorter is never spread', async () => {
  const model = scripted(outlineOf([entry('cover', 'Bikes'), entry('text-only', 'The review', [1, 2, 3, 4, 5, 6, 7, 8, 9], { words: 100 }), entry('closing', 'Thanks')]));
  const plan = await planOutline(demo, brief, { llm: model.llm, ...options });
  assert.equal(plan.slides.length, 3);
  assert.equal(plan.slides[1]?.words, 100);
});

test('spare slides are shared in proportion', async () => {
  // Two slides that each want three, and room for two slides more: each gets one.
  const model = scripted(outlineOf([entry('cover', 'Bikes'), entry('free', 'First half', [1, 2, 3, 4, 5], freeText), entry('free', 'Second half', [6, 7, 8, 9], freeText)]));
  const plan = await planOutline(demo, { ...brief, slideCount: { min: 1, max: 5 } }, { llm: model.llm, ...options });
  assert.deepEqual(
    plan.slides.map((slide) => slide.purpose),
    ['Bikes', 'First half (1/2)', 'First half (2/2)', 'Second half (1/2)', 'Second half (2/2)'],
  );
  // Without a range, each gets what it wants.
  const open = scripted(outlineOf([entry('cover', 'Bikes'), entry('free', 'First half', [1, 2, 3, 4, 5], freeText), entry('free', 'Second half', [6, 7, 8, 9], freeText)]));
  const wide = await planOutline(demo, brief, { llm: open.llm, ...options });
  assert.ok(wide.slides.length >= 5, JSON.stringify(wide.slides.map((slide) => slide.purpose)));
});

test('a fixed count keeps its slides and sets their text smaller', async () => {
  const model = scripted(outlineOf([entry('free', 'רקע ותוכנית', [1, 2, 3, 4, 5], freeText), entry('free', 'סיכונים', [6, 7, 8, 9], freeText)]));
  const plan = await planOutline(rtl, { ...hebrewBrief, slideCount: 2 }, { llm: model.llm, ...rtlOptions });
  assert.deepEqual(
    plan.slides.map((slide) => [slide.pattern, slide.blocks, slide.arrange, slide.passages]),
    [
      ['free', ['text'], { lead: 'none', dense: true }, [1, 2, 3, 4, 5]],
      ['free', ['text'], { lead: 'none', dense: true }, [6, 7, 8, 9]],
    ],
  );
  const doc = readDocument(hebrewBrief.document as string, 'he');
  const written = await writeDeck(rtl, plan, { llm: scripted(titles).llm, ...rtlOptions });
  assert.deepEqual(
    written.deck.slides.map((slide) => slide.slots.b1Text),
    [fillParagraphs(doc.passages, [1, 2, 3, 4, 5]), fillParagraphs(doc.passages, [6, 7, 8, 9])],
  );
  assert.ok(!written.warnings.some((issue) => /cut to fit/.test(issue.message)), JSON.stringify(written.warnings));

  // One slide for 7,000 characters: set at the least size that reads, and cut, said before and after writing.
  const long = hebrewDocument(7000);
  const cutDoc = readDocument(long, 'he');
  const every = cutDoc.passages.map((passage) => passage.number);
  const one = scripted(outlineOf([entry('free', 'הכול', every, freeText)]));
  const cut = await planOutline(rtl, { ...hebrewBrief, document: long, slideCount: 1 }, { llm: one.llm, ...rtlOptions });
  const [fit] = planTextFit(rtl, cut, rtlOptions);
  assert.ok(fit && fit.holds < fit.words, JSON.stringify(fit));
  assert.equal(fit?.size, readableSize(9144000));
  // Its characters, and those of the headings that start inside it, which code sets as paragraphs.
  const need = needOf(cutDoc, every, 'he');
  assert.ok(need > cutDoc.chars);
  const writer = scripted((request: JsonRequest) => ({ s1: { title: 'הכול', b1Text: ['טקסט'], ...(titles(request) as { s1: object }).s1 } }));
  const result = await writeDeck(rtl, cut, { llm: writer.llm, ...rtlOptions });
  const warning = result.warnings.find((issue) => issue.message.startsWith('the text from the document'));
  assert.match(warning?.message ?? '', new RegExp(`^the text from the document for this slide \\(p1-p${every.length}\\) is ${need} characters, and the slide holds about \\d+ even with its text set small, at 7 pt: it is cut to fit\\. Ask for more slides to keep it whole$`));
  assert.equal(warning?.slide, 1);
});

test('asked_slides fixes the count', async () => {
  const two = outlineOf([entry('free', 'First', [1, 2, 3, 4], freeText), entry('free', 'Second', [5, 6, 7, 8, 9], freeText)], { asked_slides: 2 });
  const model = scripted(two);
  const plan = await planOutline(demo, brief, { llm: model.llm, ...options });
  assert.equal(plan.slides.length, 2);
  assert.equal(plan.brief.slideCount, undefined);
  assert.equal(readAskedSlides('2'), 2);
  for (const none of [0, -1, 31, 'two', 2.5, null]) assert.equal(readAskedSlides(none), undefined, String(none));
  assert.equal(isFixedCount(3), true);
  assert.equal(isFixedCount({ min: 2, max: 2 }), true);
  assert.equal(isFixedCount({ min: 2, max: 4 }), false);
  assert.equal(isFixedCount(undefined), false);

  // An answer that does not meet the count it names is sent back once; a second one stands.
  const four = outlineOf([entry('free', 'A', [1, 2], freeText), entry('free', 'B', [3, 4], freeText), entry('free', 'C', [5, 6, 7], freeText), entry('free', 'D', [8, 9], freeText)], { asked_slides: 2 });
  const again = scripted(four, four);
  const kept = await planOutline(demo, brief, { llm: again.llm, ...options });
  assert.equal(again.requests.length, 2);
  assert.ok(again.requests[1]?.prompt.endsWith('Your previous answer could not be used: the brief asks for 2 slides and the answer has 4: answer with exactly 2. Answer again, with every passage on one slide or in "left_out".'));
  assert.equal(kept.slides.length, 4);
  assert.ok(kept.slides.every((slide) => !/\(\d+\/\d+\)$/.test(slide.purpose)));
});

test('a fixed count packs the passages again when one slide is over', async () => {
  const over = scripted(outlineOf([entry('free', 'Most', [1, 2, 3, 4, 5, 6, 7, 8], freeText), entry('free', 'Last', [9], freeText)]));
  const plan = await planOutline(rtl, { ...hebrewBrief, slideCount: 2 }, { llm: over.llm, ...rtlOptions });
  const lists = plan.slides.map((slide) => slide.passages ?? []);
  assert.deepEqual(lists.flat(), [1, 2, 3, 4, 5, 6, 7, 8, 9]);
  assert.ok(lists.every((list) => list.length > 0 && list.every((number, at) => at === 0 || number === (list[at - 1] as number) + 1)), JSON.stringify(lists));
  for (const fit of planTextFit(rtl, plan, rtlOptions)) assert.equal(fit.holds, fit.words, JSON.stringify(fit));
  // Out of order, nothing moves.
  const odd = scripted(outlineOf([entry('free', 'Most', [2, 3, 4, 5, 6, 7, 8, 9], freeText), entry('free', 'First', [1], freeText)]));
  const kept = await planOutline(rtl, { ...hebrewBrief, slideCount: 2 }, { llm: odd.llm, ...rtlOptions });
  assert.deepEqual(
    kept.slides.map((slide) => slide.passages),
    [[2, 3, 4, 5, 6, 7, 8, 9], [1]],
  );
});

test('a fixed count with no slide of content packs over the slides it has', async () => {
  const bare = outlineOf([entry('cover', 'Bikes'), entry('closing', 'Thanks')]);
  const model = scripted(bare, bare);
  const plan = await planOutline(demo, { ...brief, slideCount: 2 }, { llm: model.llm, ...options });
  assert.equal(plan.slides.length, 2);
  assert.ok(plan.slides.every((slide) => (slide.passages?.length ?? 0) > 0), JSON.stringify(plan.slides));
  assert.deepEqual(
    plan.slides.flatMap((slide) => slide.passages ?? []),
    [1, 2, 3, 4, 5, 6, 7, 8, 9],
  );
});

test("a fixed count's slide of several texts that would be cut is one text set dense, filled by code with its own passages", async () => {
  // As a live outline answered "two slides of about 500 words each, the text as it is": a text for each passage.
  const texts = (count: number) => ({ shows: [{ kind: 'text', count }], shows_main: 'none', words: 500 });
  const model = scripted(outlineOf([entry('free', 'רקע ותוכנית', [1, 2, 3, 4, 5], texts(5)), entry('free', 'סיכונים', [6, 7, 8, 9], texts(4))]));
  const plan = await planOutline(rtl, { ...hebrewBrief, slideCount: 2 }, { llm: model.llm, ...rtlOptions });
  assert.deepEqual(
    plan.slides.map((slide) => [slide.blocks, slide.arrange, slide.passages, slide.words]),
    [
      [['text'], { dense: true }, [1, 2, 3, 4, 5], undefined],
      [['text'], { dense: true }, [6, 7, 8, 9], undefined],
    ],
  );
  for (const fit of planTextFit(rtl, plan, rtlOptions)) assert.equal(fit.holds, fit.words, JSON.stringify(fit));
  const writer = scripted(titles);
  const written = await writeDeck(rtl, plan, { llm: writer.llm, ...rtlOptions });
  const doc = readDocument(hebrewBrief.document as string, 'he');
  assert.deepEqual(
    written.deck.slides.map((slide) => slide.slots.b1Text),
    [fillParagraphs(doc.passages, [1, 2, 3, 4, 5]), fillParagraphs(doc.passages, [6, 7, 8, 9])],
  );
  assert.deepEqual(
    Object.values(schemaOf(writer.requests[0]).properties).map((slide) => Object.keys(slide.properties)),
    [['title'], ['title']],
  );
  assert.deepEqual(written.warnings, []);
});

test('a fixed count of three slides or fewer asks again for a slide that carries nothing, and no slide of text stands empty while another is cut', async () => {
  // "Two slides of this text", answered with an opening slide and one slide of all of it, which would be cut.
  const text = Array.from({ length: 14 }, (_, at) => paragraph(at, 450)).join('\n\n');
  const all = documentPassages(text).map((passage) => passage.number);
  const opened = outlineOf([entry('cover', 'Bikes'), entry('free', 'All', all, freeText)]);
  const two = outlineOf([entry('free', 'First', all.slice(0, 7), freeText), entry('free', 'Then', all.slice(7), freeText)]);
  const model = scripted(opened, two);
  const plan = await planOutline(demo, { ...brief, document: text, slideCount: 2 }, { llm: model.llm, ...options });
  assert.equal(model.requests.length, 2);
  assert.ok(
    model.requests[1]?.prompt.endsWith(
      'Your previous answer could not be used: slide 1 carries no passage, and a deck of three slides or fewer carries the document on every slide unless the brief asks for an opening or a closing slide. Answer again, with every passage on one slide or in "left_out".',
    ),
    String(model.requests[1]?.prompt.slice(-300)),
  );
  assert.deepEqual(
    plan.slides.map((slide) => slide.passages),
    [all.slice(0, 7), all.slice(7)],
  );
  for (const fit of planTextFit(demo, plan, options)) assert.equal(fit.holds, fit.words, JSON.stringify(fit));

  // Slides of text the answer leaves empty, twice: they take their share, in order, each about its own passages.
  const sparse = outlineOf([entry('free', 'One', [1], freeText), entry('free', 'Two', [], freeText), entry('free', 'Three', [], freeText)]);
  const shared = await planOutline(demo, { ...brief, document: text, slideCount: 3 }, { llm: scripted(sparse, sparse).llm, ...options });
  const lists = shared.slides.map((slide) => slide.passages ?? []);
  assert.deepEqual(lists.flat(), all);
  assert.ok(lists.every((list) => list.length > 0 && list.every((number, at) => at === 0 || number === (list[at - 1] as number) + 1)), JSON.stringify(lists));
  const doc = readDocument(text, 'en');
  assert.deepEqual(
    shared.slides.map((slide) => slide.purpose),
    ['One', ...shared.slides.slice(1).map((slide) => (doc.passages[(slide.passages?.[0] as number) - 1]?.text as string).split(' ').slice(0, 12).join(' '))],
  );
  for (const fit of planTextFit(demo, shared, options)) assert.equal(fit.holds, fit.words, JSON.stringify(fit));

  // Six slides, a slide that would be cut beside an empty chart and an empty slide of text: shared
  // again over the two slides that carry passages, they hold them all, and the others are kept.
  const chart = entry('chart-native', 'Trips per year chart');
  const questions = entry('free', 'Questions', [], freeText);
  const six = outlineOf([entry('cover', 'Bikes'), entry('free', 'Start', all.slice(0, 2), freeText), entry('free', 'Rest', all.slice(2), freeText), chart, questions, entry('closing', 'Thanks')]);
  const evened = await planOutline(demo, { ...brief, document: text, slideCount: 6 }, { llm: scripted(six).llm, ...options });
  assert.deepEqual(
    evened.slides.map((slide) => [slide.pattern, slide.purpose, (slide.passages?.length ?? 0) > 0]),
    [
      ['cover', 'Bikes', false],
      ['free', 'Start', true],
      ['free', 'Rest', true],
      ['chart-native', 'Trips per year chart', false],
      ['free', 'Questions', false],
      ['closing', 'Thanks', false],
    ],
  );
  assert.deepEqual(
    evened.slides.flatMap((slide) => slide.passages ?? []),
    all,
  );
  for (const fit of planTextFit(demo, evened, options)) assert.equal(fit.holds, fit.words, JSON.stringify(fit));
  // Five, where the one slide that carries passages is cut even so: the empty slide of text takes its share, the chart does not.
  const five = outlineOf([entry('cover', 'Bikes'), entry('free', 'All', all, freeText), chart, questions, entry('closing', 'Thanks')]);
  const widened = await planOutline(demo, { ...brief, document: text, slideCount: 5 }, { llm: scripted(five).llm, ...options });
  const second = widened.slides[3]?.passages ?? [];
  assert.deepEqual(
    widened.slides.map((slide) => [slide.pattern, slide.purpose, (slide.passages?.length ?? 0) > 0]),
    [
      ['cover', 'Bikes', false],
      ['free', 'All', true],
      ['chart-native', 'Trips per year chart', false],
      ['free', purposeOf(doc.passages, second), true],
      ['closing', 'Thanks', false],
    ],
  );
  assert.deepEqual(
    widened.slides.flatMap((slide) => slide.passages ?? []),
    all,
  );
  for (const fit of planTextFit(demo, widened, options)) assert.equal(fit.holds, fit.words, JSON.stringify(fit));
});

test('with no count and nothing placed, no more slides are added than the document has passages', async () => {
  // One paragraph of about 3,000 characters and no stop: one passage, which is never cut.
  const text = `${'the council pays for the docks and a private firm runs the fleet on contract while repairs take two days on average and a van moves bikes from full stations at night '.repeat(18).trim()}`;
  assert.equal(documentPassages(text).length, 1);
  const bare = outlineOf([entry('cover', 'Bikes'), entry('closing', 'Thanks')]);
  const plan = await planOutline(demo, { ...brief, document: text }, { llm: scripted(bare, bare).llm, ...options });
  assert.deepEqual(
    plan.slides.map((slide) => [slide.pattern, slide.passages ?? []]),
    [
      ['cover', []],
      ['free', [1]],
      ['closing', []],
    ],
  );
});

// --- the person's text first (0.6.3) ---

/** Free slides of a text and some other part, as an outline answers them. */
const textAnd = (kind: string, count = 1) => ({ shows: [{ kind: 'text', count: 1 }, { kind, count }], shows_main: 'none' });
/** Whether the passages of the slides of a plan are all of them, each slide a run in order. */
const inRuns = (slides: DeckPlan['slides'], count: number): boolean => {
  const lists = slides.map((slide) => slide.passages ?? []);
  return JSON.stringify(lists.flat()) === JSON.stringify(Array.from({ length: count }, (_, at) => at + 1)) && lists.every((list) => list.length > 0 && list.every((number, at) => at === 0 || number === (list[at - 1] as number) + 1));
};

test('a slide whose text would be cut beside cards, a chart or a table is its text alone, set dense, and holds it', async () => {
  // As the production case: two slides of a long text, each a text with a card, which held about 1 character of it.
  for (const kind of ['card', 'chart', 'table', 'list', 'quote', 'note', 'figure', 'step']) {
    const model = scripted(outlineOf([entry('free', 'Background and riders', [1, 2, 3, 4], textAnd(kind)), entry('free', 'Costs and plans', [5, 6, 7, 8, 9], textAnd(kind))]));
    const plan = await planOutline(demo, { ...brief, slideCount: 2 }, { llm: model.llm, ...options });
    assert.deepEqual(
      plan.slides.map((slide) => [slide.pattern, slide.blocks, slide.arrange]),
      [
        ['free', ['text'], { dense: true }],
        ['free', ['text'], { dense: true }],
      ],
      kind,
    );
    assert.ok(inRuns(plan.slides, 9), JSON.stringify(plan.slides));
    for (const fit of planTextFit(demo, plan, options)) assert.equal(fit.holds, fit.words, `${kind}: ${JSON.stringify(fit)}`);
  }
  // A layout the outline chose that would become a free slide of a text and its chart or its table: the same.
  for (const layout of ['chart-native', 'table']) {
    const model = scripted(outlineOf([entry(layout, 'The review', [1, 2, 3, 4, 5, 6, 7, 8, 9])]));
    const plan = await planOutline(demo, { ...brief, slideCount: 1 }, { llm: model.llm, ...options });
    assert.deepEqual(plan.slides, [{ pattern: 'free', purpose: 'The review', passages: [1, 2, 3, 4, 5, 6, 7, 8, 9], blocks: ['text'], arrange: { dense: true } }], layout);
    const written = await writeDeck(demo, plan, { llm: scripted(titles).llm, ...options });
    assert.deepEqual(written.deck.slides[0]?.slots.b1Text, fillParagraphs(readDocument(REPORT, 'en').passages, [1, 2, 3, 4, 5, 6, 7, 8, 9]), layout);
    assert.ok(!written.warnings.some((issue) => /cut to fit/.test(issue.message)), JSON.stringify(written.warnings));
  }
});

/** A document of a lead line and `count` items with no full stop, each `chars` characters or a word more. */
const listDocument = (count: number, chars: number): string => {
  const words = 'riders took more trips every year and the busiest month was always June while students and commuters make up most members and visitors buy day passes at each station with room for twelve bikes and a screen that shows the nearest dock'.split(' ');
  let at = 0;
  const item = (): string => {
    let line = '';
    while (line.length < chars) line += `${line ? ' ' : ''}${words[at++ % words.length]}`;
    return line.charAt(0).toUpperCase() + line.slice(1);
  };
  return ['Tasks for the quarter', '', ...Array.from({ length: count }, () => [item(), '']).flat()].join('\n');
};

test('a list whose items have no full stop counts the headings code sets, and code sets it only where it fits', async () => {
  // Lines with no full stop are read as headings of the document, and code sets each as a paragraph
  // of the slide's text. Left out of what a slide holds, 53 such lines were planned as held,
  // set word for word, and ran past their box in PowerPoint.
  const listOf = (count: number): string => listDocument(count, 58);
  const planned = async (text: string) => {
    const doc = readDocument(text, 'en');
    const every = doc.passages.map((passage) => passage.number);
    const model = scripted(outlineOf([entry('free', 'All of it', every, freeText)]));
    return { doc, every, plan: await planOutline(demo, { ...brief, document: text, slideCount: 1 }, { llm: model.llm, ...options }) };
  };
  // 52 items: most of them headings. What the slide needs counts them, as many paragraphs as code sets.
  const long = await planned(listOf(52));
  const filled = fillParagraphs(long.doc.passages, long.every);
  assert.ok(long.doc.passages.filter((passage) => passage.section !== undefined).length > 20);
  assert.equal(needOf(long.doc, long.every, 'en'), filled.reduce((sum, paragraph) => sum + paragraph.length, 0));
  assert.equal(filled.length, 53);
  // The slide does not hold them as they are, each item counted two lines: code does not set them,
  // and the writer writes the text, told the most paragraphs that hold all of it, more than twelve.
  const writer = scripted((request: JsonRequest) => {
    const count = schemaOf(request).properties.s1?.properties.b1Text?.maxItems ?? 1;
    const per = Math.ceil(filled.length / count);
    return { s1: { title: 'All of it', b1Text: Array.from({ length: Math.ceil(filled.length / per) }, (_, at) => filled.slice(at * per, at * per + per).join(' ')) } };
  });
  const [whole] = planTextFit(demo, long.plan, options);
  assert.ok(whole && whole.dense === true && whole.holds === whole.words, JSON.stringify(whole));
  const joined = await writeDeck(demo, long.plan, { llm: writer.llm, ...options });
  const told = schemaOf(writer.requests[0]).properties.s1?.properties.b1Text;
  assert.ok(told && (told.maxItems ?? 0) > 12 && (told.maxItems ?? 0) < 53, `the writer writes the text: ${told?.maxItems}`);
  // Nothing is cut, and the deck says its lines are joined.
  assert.deepEqual(
    joined.warnings.map((issue) => issue.message),
    [`the text from the document for this slide (p1-p${long.every.length}) has 53 paragraphs, and the slide holds it only as ${told?.maxItems}: its paragraphs are joined. Ask for more slides to keep them apart`],
  );
  // 90 items: more characters than the slide holds. The plan says it is cut, the writer is told
  // twelve paragraphs and what the slide holds, and the deck says the text is cut.
  const longer = await planned(listOf(90));
  const [fit] = planTextFit(demo, longer.plan, options);
  assert.ok(fit && fit.dense === true && fit.holds < fit.words, JSON.stringify(fit));
  const again = scripted((request: JsonRequest) => {
    const text = schemaOf(request).properties.s1?.properties.b1Text;
    const top = Number(/ to (\d+) characters/.exec(text?.description ?? '')?.[1]);
    const count = text?.maxItems ?? 1;
    return { s1: { title: 'All of it', b1Text: Array.from({ length: count }, (_, at) => paragraph(at, Math.floor(top / count)).slice(0, Math.floor(top / count))) } };
  });
  const written = await writeDeck(demo, longer.plan, { llm: again.llm, ...options });
  assert.equal(schemaOf(again.requests[0]).properties.s1?.properties.b1Text?.maxItems, 12);
  assert.ok(written.warnings.some((issue) => /cut to fit/.test(issue.message)), JSON.stringify(written.warnings));
  assert.ok(!written.warnings.some((issue) => /the slot fits/.test(issue.message)), JSON.stringify(written.warnings));
  // 16 items the slide holds as they are: code sets them, headings and all, and the deck says nothing.
  const short = await planned(listOf(16));
  const [held] = planTextFit(demo, short.plan, options);
  assert.ok(held && held.holds === held.words, JSON.stringify(held));
  const set = await writeDeck(demo, short.plan, { llm: scripted(titles).llm, ...options });
  assert.deepEqual(set.deck.slides[0]?.slots.b1Text, fillParagraphs(short.doc.passages, short.every));
  assert.deepEqual(set.warnings, []);
});

test('a list its slide holds only with some lines joined keeps all its words', async () => {
  // Seventy short lines are more paragraphs than the slide holds as they are, but not more
  // characters. The writer is told the most paragraphs that hold them all, more than twelve, and
  // nothing is cut. Told twelve, as a text that is cut is, it would join six lines into each. The
  // deck says the lines are joined.
  const text = listDocument(70, 30);
  const doc = readDocument(text, 'en');
  const every = doc.passages.map((passage) => passage.number);
  const plan = await planOutline(demo, { ...brief, document: text, slideCount: 1 }, { llm: scripted(outlineOf([entry('free', 'All of it', every, freeText)])).llm, ...options });
  const [fit] = planTextFit(demo, plan, options);
  assert.ok(fit && fit.dense === true && fit.holds === fit.words, JSON.stringify(fit));
  const writer = scripted((request: JsonRequest) => {
    const told = schemaOf(request).properties.s1?.properties.b1Text;
    const count = told?.maxItems ?? 1;
    const lines = fillParagraphs(doc.passages, every);
    const per = Math.ceil(lines.length / count);
    return { s1: { title: 'All of it', b1Text: Array.from({ length: Math.ceil(lines.length / per) }, (_, at) => lines.slice(at * per, at * per + per).join(' ')) } };
  });
  const written = await writeDeck(demo, plan, { llm: writer.llm, ...options });
  const told = schemaOf(writer.requests[0]).properties.s1?.properties.b1Text;
  assert.ok((told?.maxItems ?? 0) > 12 && (told?.maxItems ?? 0) < 71, String(told?.maxItems));
  assert.equal(writer.requests.length, 1);
  assert.deepEqual(
    written.warnings.map((issue) => issue.message),
    [`the text from the document for this slide (p1-p${every.length}) has 71 paragraphs, and the slide holds it only as ${told?.maxItems}: its paragraphs are joined. Ask for more slides to keep them apart`],
  );
});

test('a chart a change adds beside a text from the document stays at the next change', async () => {
  // A plan of two slides of text; a change adds a chart beside the text of slide 1, and the next
  // moves a passage. The chart was asked for: it stays, and the plan says what the text keeps.
  const two = scripted(outlineOf([entry('free', 'Background and riders', [1, 2, 3, 4], freeText), entry('free', 'Costs and plans', [5, 6, 7, 8, 9], freeText)]));
  const plan = await planOutline(demo, { ...brief, slideCount: 2 }, { llm: two.llm, ...options });
  const chart = { shows: [{ kind: 'text', count: 1 }, { kind: 'chart', count: 1 }], shows_main: 'auto', side_by_side: 0 };
  const change = (first: number[], second: number[]) => ({
    summary: 'Done.',
    language: 'en',
    title: 'Bikes',
    left_out: [],
    asked_slides: 0,
    document_language: 'en',
    slides: [
      { pattern: 'free', purpose: 'Background and riders', ...chart, passages: first },
      { pattern: 'free', purpose: 'Costs and plans', ...freeText, side_by_side: 0, passages: second },
    ],
  });
  const { plan: next } = await editPlan(demo, plan, 'Add a chart beside the text on slide 1', { llm: scripted(change([1, 2, 3, 4], [5, 6, 7, 8, 9])).llm, ...options });
  assert.deepEqual(next.slides[0]?.blocks, ['text', 'chart']);
  assert.equal(next.slides[0]?.partsAsked, true);
  const { plan: moved } = await editPlan(demo, next, 'Move p5 to slide 1', { llm: scripted(change([1, 2, 3, 4, 5], [6, 7, 8, 9])).llm, ...options });
  assert.deepEqual([moved.slides[0]?.blocks, moved.slides[0]?.passages], [['text', 'chart'], [1, 2, 3, 4, 5]]);
  assert.equal(moved.slides[0]?.partsAsked, true);
  const [fit] = planTextFit(demo, moved, options);
  assert.ok(fit && fit.holds < fit.words, JSON.stringify(fit));
});

test('a slide that holds its text beside its other parts keeps them', async () => {
  // A chart beside one passage holds it: the chart stays.
  const model = scripted(outlineOf([entry('free', 'Background', [1], textAnd('chart')), entry('free', 'Riders and costs', [2, 3, 4, 5], freeText), entry('free', 'Plans', [6, 7, 8, 9], freeText)]));
  const plan = await planOutline(demo, { ...brief, slideCount: 3 }, { llm: model.llm, ...options });
  assert.deepEqual(plan.slides[0]?.blocks, ['text', 'chart']);
  assert.deepEqual(
    plan.slides.map((slide) => slide.passages),
    [[1], [2, 3, 4, 5], [6, 7, 8, 9]],
  );
  for (const fit of planTextFit(demo, plan, options)) assert.equal(fit.holds, fit.words, JSON.stringify(fit));
});

test('editPlan: a slide given more passages than it holds beside its chart is its text alone', async () => {
  const plan: DeckPlan = {
    version: 1,
    title: 'Bikes',
    brief: { ...brief, slideCount: 2 },
    slides: [
      { pattern: 'free', purpose: 'Background', passages: [1], blocks: ['text', 'chart'] },
      { pattern: 'free', purpose: 'The rest', passages: [2, 3, 4, 5, 6, 7, 8, 9], blocks: ['text'], arrange: { dense: true } },
    ],
    documentLanguage: 'en',
  };
  const model = scripted(changeOf([entry('free', 'Background', [1, 2, 3, 4, 5], textAnd('chart')), entry('free', 'The rest', [6, 7, 8, 9], freeText)]));
  const { plan: next } = await editPlan(demo, plan, 'Move the riders part to the first slide', { llm: model.llm, ...options });
  assert.deepEqual(
    next.slides.map((slide) => [slide.blocks, slide.arrange?.dense, slide.passages]),
    [
      [['text'], true, [1, 2, 3, 4, 5]],
      [['text'], true, [6, 7, 8, 9]],
    ],
  );
  for (const fit of planTextFit(demo, next, options)) assert.equal(fit.holds, fit.words, JSON.stringify(fit));
});

test('a deck of three slides or fewer keeps the passages an answer gives its opening and closing slides, on slides of text, where the slide next to them would cut them', async () => {
  // A report too long for one slide: a cover and a closing slide that show no running text each keep their passages, as a slide of text.
  const long = [REPORT, '', 'More', ...[0, 1, 2, 3, 4, 5, 6, 7].map((at) => paragraph(at, 400))].join('\n');
  const count = documentPassages(long).length;
  const upTo = (from: number, to: number) => Array.from({ length: to - from + 1 }, (_, at) => from + at);
  const model = scripted(outlineOf([entry('cover', 'Bikes', upTo(1, 9)), entry('closing', 'Thanks', upTo(10, count))]));
  const plan = await planOutline(demo, { ...brief, document: long, slideCount: 2 }, { llm: model.llm, ...options });
  assert.equal(model.requests.length, 1);
  assert.deepEqual(
    plan.slides.map((slide) => [slide.pattern, slide.blocks, slide.passages]),
    [
      ['free', ['text'], upTo(1, 9)],
      ['free', ['text'], upTo(10, count)],
    ],
  );
  for (const fit of planTextFit(demo, plan, options)) assert.equal(fit.holds, fit.words, JSON.stringify(fit));
  // Where the slide next to it holds them beside its own, the cover gives them on, as before, and stays a cover.
  const fits = scripted(outlineOf([entry('cover', 'Bikes', [1, 2, 3, 4]), entry('free', 'The rest', [5, 6, 7, 8, 9], freeText)]));
  const kept = await planOutline(demo, { ...brief, slideCount: 2 }, { llm: fits.llm, ...options });
  assert.deepEqual(
    kept.slides.map((slide) => [slide.pattern, slide.passages ?? []]),
    [
      ['cover', []],
      ['free', upTo(1, 9)],
    ],
  );
  // So does a cover given only the document's title line, however long the rest: a title is what it shows.
  const titled = `The bike share in 2026\n\n${long}`;
  const titleModel = scripted(outlineOf([entry('cover', 'Bikes', [1]), entry('free', 'Part one', upTo(2, 10), freeText), entry('free', 'Part two', upTo(11, count + 1), freeText)]));
  const cover = await planOutline(demo, { ...brief, document: titled, slideCount: 3 }, { llm: titleModel.llm, ...options });
  assert.deepEqual(
    cover.slides.map((slide) => [slide.pattern, slide.passages?.[0]]),
    [
      ['cover', undefined],
      ['free', 1],
      ['free', 11],
    ],
  );
  // A deck whose number of slides the brief names (asked_slides) does the same, and with free slides off the slide is a layout of plain text.
  const named = scripted(outlineOf([entry('cover', 'Bikes', [1, 2]), entry('closing', 'Thanks', [3])], { asked_slides: 2, left_out: [4, 5, 6, 7, 8, 9] }));
  const off = await planOutline(demo, brief, { llm: named.llm, ...options, free: false });
  assert.deepEqual(
    off.slides.map((slide) => [slide.pattern, slide.passages]),
    [
      ['text-only', [1, 2]],
      ['text-only', [3]],
    ],
  );
  // Four slides: the passages of an opening or a closing slide still go to the slide next to it.
  const four = scripted(outlineOf([entry('cover', 'Bikes', [1, 2]), entry('free', 'Middle', [3, 4, 5, 6], freeText), entry('free', 'Late', [7], freeText), entry('closing', 'Thanks', [8, 9])]));
  const moved = await planOutline(demo, { ...brief, slideCount: 4 }, { llm: four.llm, ...options });
  assert.deepEqual(
    moved.slides.map((slide) => [slide.pattern, slide.passages ?? []]),
    [
      ['cover', []],
      ['free', [1, 2, 3, 4, 5, 6]],
      ['free', [7, 8, 9]],
      ['closing', []],
    ],
  );
  // A change to a plan that names two slides does the same.
  const longBrief = { ...brief, document: long, slideCount: 3 };
  const thirds = [upTo(1, 6), upTo(7, 12), upTo(13, count)];
  const stored: DeckPlan = { version: 1, title: 'Bikes', brief: longBrief, slides: (await planOutline(demo, longBrief, { llm: scripted(outlineOf(thirds.map((part, at) => entry('free', `Part ${at + 1}`, part, freeText)))).llm, ...options })).slides, documentLanguage: 'en' };
  const change = scripted(changeOf([entry('cover', 'Bikes', upTo(1, 9)), entry('free', 'The rest', upTo(10, count), freeText)], { asked_slides: 2 }));
  const { plan: two } = await editPlan(demo, stored, 'Keep it to 2 slides', { llm: change.llm, ...options });
  assert.deepEqual(
    two.slides.map((slide) => [slide.pattern, slide.blocks]),
    [
      ['free', ['text']],
      ['free', ['text']],
    ],
  );
  assert.ok(inRuns(two.slides, count), JSON.stringify(two.slides));
});

test('with free slides off a slide of a document moves to the layout of plain text that holds most, says what it holds as the writer is told it, and why it is not set smaller', async () => {
  const off = { ...options, free: false };
  // One slide of quotes for two passages: the quote holds none of them, the text layout most.
  const model = scripted(outlineOf([entry('quote', 'Background', [1, 2])], { left_out: [3, 4, 5, 6, 7, 8, 9] }));
  const plan = await planOutline(demo, { ...brief, slideCount: 1 }, { llm: model.llm, ...off });
  assert.deepEqual(plan.slides, [{ pattern: 'text-only', purpose: 'Background', passages: [1, 2] }]);
  const doc = readDocument(REPORT, 'en');
  const need = (doc.passages[0] as DocumentPassage).chars + (doc.passages[1] as DocumentPassage).chars;
  const words = passageWords(doc.passages, [1, 2]);
  // Two paragraphs: nine tenths of the 331 characters of the text layout.
  const room = Math.floor(0.9 * 331);
  assert.deepEqual(planTextFit(demo, plan, off), [{ slide: 1, words, holds: Math.floor((words * room) / need), passages: 'p1-p2' }]);
  const writer = scripted((request: JsonRequest) => ({ s1: { title: 'Background', body: ['Riders rely on it.'], ...(titles(request) as { s1: object }).s1 } }));
  const written = await writeDeck(demo, plan, { llm: writer.llm, ...off });
  const body = (schemaOf(writer.requests[0]).properties.s1 as Schema).properties.body?.description ?? '';
  assert.ok(body.startsWith("The slide's text from the document (p1-p2) shortened to 260 to 290 characters, the most this slide holds"), body);
  assert.deepEqual(
    withoutCodes(written.warnings.filter((issue) => /cut to fit/.test(issue.message))),
    [
      {
        severity: 'warning',
        slide: 1,
        message: `the text from the document for this slide (p1-p2) is ${need} characters, and the slide holds about ${room}: it is cut to fit. Ask for more slides to keep it whole; free slides are off, so its text is not set smaller`,
      },
    ],
  );
  // A list of patterns without free is the same; with free slides on, nothing is added to the message.
  const listed = await writeDeck(demo, { ...plan, brief: { ...plan.brief, patterns: ['text-only', 'quote'] } }, { llm: scripted(titles).llm, ...options });
  assert.ok(listed.warnings.some((issue) => issue.message.endsWith('; free slides are off, so its text is not set smaller')), JSON.stringify(listed.warnings));
  const on = await writeDeck(demo, plan, { llm: scripted(titles).llm, ...options });
  assert.ok(on.warnings.some((issue) => issue.message.endsWith(': it is cut to fit. Ask for more slides to keep it whole')), JSON.stringify(on.warnings));
});

test('with free slides off a table of the document on a layout of a table that holds none of its text moves to plain text, its rows counted', async () => {
  const off = { ...options, free: false };
  const table = [`Costs`, paragraph(0, 190), '| Item | 2023 | 2024 |', '|---|---|---|', '| Bikes | 120 | 140 |', '| Staff | 300 | 320 |', '| Repairs | 80 | 95 |'].join('\n');
  const model = scripted(outlineOf([entry('table', 'Costs', [1, 2])]));
  const plan = await planOutline(demo, { ...brief, document: table, slideCount: 1 }, { llm: model.llm, ...off });
  assert.deepEqual(plan.slides.map((slide) => [slide.pattern, slide.passages]), [['text-only', [1, 2]]]);
  // The layout of plain text is judged by the text with the rows of the table in it, as they will be set there.
  const doc = readDocument(table, 'en');
  const [fit] = planTextFit(demo, plan, off);
  assert.equal(fit?.words, passageWords(doc.passages, [1, 2]));
  assert.ok((fit?.holds ?? 0) > 0, JSON.stringify(fit));
  // A layout of a table that holds the text beside its table keeps its table: here a table alone.
  const alone = ['Costs', '| Item | 2023 | 2024 |', '| Bikes | 120 | 140 |', '| Staff | 300 | 320 |'].join('\n');
  const kept = await planOutline(demo, { ...brief, document: alone, slideCount: 1 }, { llm: scripted(outlineOf([entry('table', 'Costs', [1])])).llm, ...off });
  assert.equal(kept.slides[0]?.pattern, 'table');
});

test('with free slides off a change to a deck keeps the layout of a slide whose passages change, and its parts', async () => {
  const off = { ...options, free: false };
  const plan: DeckPlan = { version: 1, title: 'Bikes', brief: { ...brief, slideCount: 2 }, documentLanguage: 'en', slides: [{ pattern: 'stat-cards', purpose: 'Background', passages: [1] }, { pattern: 'text-only', purpose: 'Rest', passages: [2, 3, 4, 5, 6] }] };
  const deck = (await writeDeck(demo, plan, { llm: scripted(titles).llm, ...off })).deck;
  assert.equal(deck.slides[0]?.pattern, 'stat-cards');
  const model = scripted(editOf([{ from: 1, pattern: 'stat-cards', change: 'Add the riders', passages: [1, 2, 3] }, { from: 2, pattern: 'text-only', passages: [4, 5, 6] }]), titles, titles, titles);
  const edited = await planEdit(demo, deck, 'Move the riders part to slide 1', { llm: model.llm, ...off });
  assert.deepEqual(
    edited.deck.slides.map((slide) => [slide.pattern, slide.passages]),
    [
      ['stat-cards', [1, 2, 3]],
      ['text-only', [4, 5, 6]],
    ],
  );
  assert.ok('stat1Number' in (edited.deck.slides[0]?.slots ?? {}));
  assert.ok(edited.warnings.some((issue) => issue.slide === 1 && issue.message.endsWith('; free slides are off, so its text is not set smaller')), JSON.stringify(edited.warnings));
});

test('a text of more paragraphs than a layout takes is set by code on a free slide, as before', async () => {
  // Six short sentences, one passage of six paragraphs: a text layout takes three, a free slide all six as they are.
  const lines = `Notes\n${['Trips by bike grew a fifth this year.', 'Most new riders are under forty.', 'Lanes on Main Street are full at eight.', 'Parking for bikes runs out by nine.', 'Two schools asked for safe routes.', 'The council will vote in March.'].join('\n')}`;
  const model = scripted(outlineOf([entry('text-only', 'Notes', [1])]));
  const plan = await planOutline(demo, { ...brief, document: lines, slideCount: 1 }, { llm: model.llm, ...options });
  assert.deepEqual(plan.slides.map((slide) => [slide.pattern, slide.blocks]), [['free', ['text']]]);
  const writer = scripted(titles);
  const written = await writeDeck(demo, plan, { llm: writer.llm, ...options });
  assert.equal(writer.requests.length, 1);
  assert.equal((written.deck.slides[0]?.slots.b1Text as string[]).length, 6);
});

test('editPlan keeps a part the request asks for beside a text it then cuts, and the plan says so', async () => {
  const stored: DeckPlan = {
    version: 1,
    title: 'Bikes',
    brief: { ...brief, slideCount: 2 },
    slides: [
      { pattern: 'free', purpose: 'Background', passages: [1, 2, 3, 4, 5], blocks: ['text'], arrange: { dense: true } },
      { pattern: 'free', purpose: 'The rest', passages: [6, 7, 8, 9], blocks: ['text'], arrange: { dense: true } },
    ],
    documentLanguage: 'en',
  };
  const model = scripted(changeOf([entry('free', 'Background', [1, 2, 3, 4, 5], textAnd('chart')), entry('free', 'The rest', [6, 7, 8, 9], freeText)]));
  const { plan: next } = await editPlan(demo, stored, 'Add a chart of the trips to slide 1', { llm: model.llm, ...options });
  assert.deepEqual(next.slides[0]?.blocks, ['text', 'chart']);
  const [first] = planTextFit(demo, next, options);
  assert.ok(first !== undefined && first.holds < first.words, JSON.stringify(first));
});

test('with free slides off a change to a plan moves a free slide of a text and a card to plain text where that holds more of it', async () => {
  // A plan made with free slides on, of two slides of a text and a card each, which hold little of their text.
  const long = [REPORT, '', 'More', ...[0, 1, 2, 3].map((at) => paragraph(at, 400))].join('\n');
  const count = documentPassages(long).length;
  const upTo = (from: number, to: number) => Array.from({ length: to - from + 1 }, (_, at) => from + at);
  const half = Math.ceil(count / 2);
  const stored: DeckPlan = {
    version: 1,
    title: 'Bikes',
    brief: { ...brief, document: long, slideCount: 2 },
    slides: [
      { pattern: 'free', purpose: 'First', passages: upTo(1, half), blocks: ['text', 'card'] },
      { pattern: 'free', purpose: 'Then', passages: upTo(half + 1, count), blocks: ['text', 'card'] },
    ],
    documentLanguage: 'en',
  };
  const off = { ...options, free: false };
  const cut = planTextFit(demo, stored, off);
  const model = scripted(changeOf([entry('free', 'First', upTo(1, half), textAnd('card')), entry('free', 'Then', upTo(half + 1, count), textAnd('card'))], { asked_slides: 2 }));
  const { plan: next } = await editPlan(demo, stored, 'Keep it to 2 slides', { llm: model.llm, ...off });
  // The first moves to the layout of plain text, which holds more of it; the second holds more as it is.
  assert.deepEqual(
    next.slides.map((slide) => [slide.pattern, slide.blocks]),
    [
      ['text-only', undefined],
      ['free', ['text', 'card']],
    ],
  );
  const fits = planTextFit(demo, next, off);
  assert.ok((fits[0]?.holds ?? 0) > (cut[0]?.holds ?? 0), JSON.stringify([fits[0], cut[0]]));
  assert.deepEqual(fits[1], cut[1]);
  // A free slide of its text alone stays one.
  const alone: DeckPlan = { ...stored, slides: stored.slides.map((slide) => ({ ...slide, blocks: ['text'], arrange: { dense: true } })) };
  const same = scripted(changeOf([entry('free', 'First', upTo(1, half), freeText), entry('free', 'Then', upTo(half + 1, count), freeText)], { asked_slides: 2 }));
  const { plan: kept } = await editPlan(demo, alone, 'Keep it to 2 slides', { llm: same.llm, ...off });
  assert.deepEqual(
    kept.slides.map((slide) => [slide.pattern, slide.blocks]),
    [
      ['free', ['text']],
      ['free', ['text']],
    ],
  );
});

test("a layout's slot of four paragraphs or more takes one fewer of a text of three or more: code sets it on a free slide, or the writer joins them", async () => {
  // One passage of four lines of about 66 characters. The demo's list beside a photo takes four
  // paragraphs, and so three of such a text.
  const lines = [
    'Notes',
    'Trips by bike grew by a fifth this year, the most in the region.',
    'Most of the new riders are under forty and live near the centre.',
    'The lanes on Main Street are full at eight every working morning.',
    'Parking for bikes at the station runs out by nine on most days.',
  ].join('\n');
  assert.equal(documentPassages(lines).length, 1);
  const assets = [{ id: 'bike', kind: 'photo' as const, description: 'A bike at a dock', data: new Uint8Array() }];
  const listed = { ...brief, imageDensity: 'medium' as const, document: lines, slideCount: 1 };
  // Its writer keeps the words, and joins the last two lines.
  const joined = [...lines.split('\n').slice(1, 3), lines.split('\n').slice(3).join(' ')];
  /** A writer that gives every slide it is asked for what it must have, the photo too. */
  const withPhoto = (request: JsonRequest): unknown =>
    Object.fromEntries(
      Object.entries(schemaOf(request).properties).map(([key, slide]) => [
        key,
        Object.fromEntries((slide.required ?? []).map((name) => [name, slide.properties[name]?.type === 'object' ? { asset: 'bike' } : slide.properties[name]?.type === 'array' ? joined : 'Notes'])),
      ]),
    );
  // With free slides on, a free slide takes the text, and code sets its four paragraphs as they are.
  const on = await planOutline(demo, listed, { llm: scripted(outlineOf([entry('list-with-photo', 'Notes', [1])])).llm, ...options, assets });
  assert.equal(on.slides[0]?.pattern, 'free');
  const set = await writeDeck(demo, on, { llm: scripted(withPhoto, withPhoto).llm, ...options, assets });
  assert.deepEqual(set.deck.slides[0]?.slots.b1Text, lines.split('\n').slice(1));
  // With free slides off the layout stays, and its writer is told three and joins them: code sets none.
  const off = { ...options, free: false, assets };
  const kept = await planOutline(demo, listed, { llm: scripted(outlineOf([entry('list-with-photo', 'Notes', [1])])).llm, ...off });
  assert.equal(kept.slides[0]?.pattern, 'list-with-photo');
  const writer = scripted(withPhoto, withPhoto);
  const written = await writeDeck(demo, kept, { llm: writer.llm, ...off });
  assert.equal(schemaOf(writer.requests[0]).properties.s1?.properties.bulletList?.maxItems, 3);
  assert.deepEqual(written.deck.slides[0]?.slots.bulletList, joined);
  // A text of three such lines is within the rule: code sets it there as it is.
  const three = lines.split('\n').slice(0, 4).join('\n');
  const short = await planOutline(demo, { ...listed, document: three }, { llm: scripted(outlineOf([entry('list-with-photo', 'Notes', [1])])).llm, ...options, assets });
  assert.equal(short.slides[0]?.pattern, 'list-with-photo');
  const filled = await writeDeck(demo, short, { llm: scripted(withPhoto, withPhoto).llm, ...options, assets });
  assert.deepEqual(filled.deck.slides[0]?.slots.bulletList, three.split('\n').slice(1));
});

test("an opening or a closing slide of a table given only the document's tables keeps them, as one in the middle of a deck does", async () => {
  const table = ['| Item | 2023 | 2024 |', '|---|---|---|', '| Bikes | 120 | 140 |', '| Staff | 300 | 320 |'].join('\n');
  const doc = `${REPORT}\n\n${table}`;
  const passages = documentPassages(doc);
  const at = (passages.find((passage) => passage.kind === 'table') as DocumentPassage).number;
  const text = passages.filter((passage) => passage.number !== at).map((passage) => passage.number);
  const shared = { ...brief, document: doc };
  for (const [label, slides, free] of [
    ['two slides, the table last', [entry('free', 'The review', text, freeText), entry('table', 'Costs', [at])], true],
    ['two slides, the table first', [entry('table', 'Costs', [at]), entry('free', 'The review', text, freeText)], true],
    ['three slides, the table last', [entry('free', 'Part one', text.slice(0, 4), freeText), entry('free', 'Part two', text.slice(4), freeText), entry('table', 'Costs', [at])], true],
    ['two slides, the table last, free slides off', [entry('text-only', 'The review', text), entry('table', 'Costs', [at])], false],
  ] as const) {
    const model = scripted(outlineOf([...slides]));
    const plan = await planOutline(demo, { ...shared, slideCount: slides.length }, { llm: model.llm, ...options, ...(free ? {} : { free: false }) });
    assert.equal(model.requests.length, 1, label);
    const table = plan.slides.find((slide) => slide.pattern === 'table');
    assert.deepEqual(table?.passages, [at], label);
    assert.deepEqual(plan.slides.filter((slide) => slide !== table).flatMap((slide) => slide.passages ?? []), text, label);
  }
  // The table slide is written from the document's table: its writer is given the rows.
  const plan = await planOutline(demo, { ...shared, slideCount: 2 }, { llm: scripted(outlineOf([entry('free', 'The review', text, freeText), entry('table', 'Costs', [at])])).llm, ...options });
  const writer = scripted(() => {
    throw new Error('stop');
  });
  await writeDeck(demo, plan, { llm: writer.llm, ...options }).catch(() => undefined);
  const prompt = writer.requests[0]?.prompt ?? '';
  const block = prompt.slice(prompt.indexOf('Text of s2 from the document'));
  assert.ok(prompt.includes('Text of s2 from the document') && block.includes('Bikes') && block.includes('120'), prompt.slice(-600));
});

test('a table on a slide with no table is running text: it needs room there, and a spread slide gives its table to the part with the table', async () => {
  const rows = Array.from({ length: 10 }, (_, at) => `| Station ${at + 1} | ${100 + at * 7} trips | ${20 + at} bikes |`);
  const text = ['Stations', 'The figures below are for the first half of the year, counted at every dock in the city.', '| Station | Trips | Bikes |', '|---|---|---|', ...rows, '', 'Plans', SENTENCES[11]].join('\n');
  const doc = readDocument(text, 'en');
  assert.deepEqual(
    doc.passages.map((passage) => passage.kind),
    ['text', 'table', 'text'],
  );
  // On a free slide of text, its rows are paragraphs of the text, set by code where the slide holds them.
  const model = scripted(outlineOf([entry('free', 'Stations', [1, 2], freeText), entry('free', 'Plans', [3], freeText)]));
  const plan = await planOutline(demo, { ...brief, document: text, slideCount: 2 }, { llm: model.llm, ...options });
  const [fit] = planTextFit(demo, plan, options);
  assert.equal(fit?.holds, fit?.words);
  const writer = scripted(titles);
  const written = await writeDeck(demo, plan, { llm: writer.llm, ...options });
  assert.deepEqual(written.deck.slides[0]?.slots.b1Text, fillParagraphs(doc.passages, [1, 2]));
  assert.equal('b1Text' in (schemaOf(writer.requests[0]).properties.s1?.properties ?? {}), false);
  assert.deepEqual(documentKept(written.deck), { words: doc.words, kept_words: doc.words });

  // A free slide of a text and a table, spread over several: the part that carries the table has the table.
  const long = [paragraph(0, 450), paragraph(2, 450), paragraph(4, 450), paragraph(6, 450), '', '| Item | 2023 | 2024 |', '|---|---|---|', '| Bikes | 120 | 140 |', '| Staff | 300 | 320 |', '', paragraph(8, 450)].join('\n');
  const passages = documentPassages(long);
  const table = (passages.find((passage) => passage.kind === 'table') as DocumentPassage).number;
  const both = { shows: [{ kind: 'text', count: 1 }, { kind: 'table', count: 1 }], shows_main: 'none' };
  const spread = await planOutline(demo, { ...brief, document: long }, { llm: scripted(outlineOf([entry('free', 'Everything', passages.map((passage) => passage.number), both)])).llm, ...options });
  assert.ok(spread.slides.length > 2, JSON.stringify(spread.slides));
  const hasTable = (slide: DeckPlan['slides'][number]): boolean => slide.pattern === 'table' || (slide.blocks ?? []).includes('table');
  for (const slide of spread.slides) assert.equal(hasTable(slide), (slide.passages ?? []).includes(table), JSON.stringify(slide));
  // So does a layout with a table: it goes to the part with the table, and the others are of text.
  const layout = await planOutline(demo, { ...brief, document: long }, { llm: scripted(outlineOf([entry('table', 'Everything', passages.map((passage) => passage.number))])).llm, ...options });
  assert.ok(layout.slides.length > 2, JSON.stringify(layout.slides));
  for (const slide of layout.slides) assert.equal(hasTable(slide), (slide.passages ?? []).includes(table), JSON.stringify(slide));

  // A fixed count shares the passages again: the table of a slide goes with the table of its passages.
  const table4 = ['| Item | 2023 | 2024 |', '|---|---|---|', '| Bikes | 120 | 140 |', '| Staff | 300 | 320 |'];
  const over = [...[0, 2, 4].map((at) => paragraph(at, 150)), '', ...table4, '', paragraph(1, 80)].join('\n');
  const shared = documentPassages(over);
  const at = (shared.find((passage) => passage.kind === 'table') as DocumentPassage).number;
  const first = shared.filter((passage) => passage.number <= at).map((passage) => passage.number);
  const outline = outlineOf([entry('free', 'First', first, both), entry('free', 'Then', shared.filter((passage) => passage.number > at).map((passage) => passage.number), freeText)]);
  const balanced = await planOutline(demo, { ...brief, document: over, slideCount: 2 }, { llm: scripted(outline).llm, ...options });
  assert.notDeepEqual(balanced.slides[0]?.passages, first);
  for (const slide of balanced.slides) assert.equal(hasTable(slide), (slide.passages ?? []).includes(at), JSON.stringify(slide));
  for (const fit of planTextFit(demo, balanced, options)) assert.equal(fit.holds, fit.words, JSON.stringify(fit));
  // A table the outline gave a slide whose passages have none stays: it is no table of the document.
  const own = outlineOf([entry('free', 'First', first.slice(0, -1), both), entry('free', 'Then', shared.filter((passage) => passage.number >= at).map((passage) => passage.number), both)]);
  const kept = await planOutline(demo, { ...brief, document: over, slideCount: 2 }, { llm: scripted(own).llm, ...options });
  assert.notDeepEqual(kept.slides[0]?.passages, first.slice(0, -1));
  assert.ok(kept.slides.every(hasTable), JSON.stringify(kept.slides));

  // Where the slide with the table is cut even once they are shared again, the person's text comes
  // first: the slide is its text alone, set dense, the table's rows paragraphs of it, and holds it all.
  const longer = [...[0, 2, 4, 6, 8].map((at) => paragraph(at, 450)), '', ...table4, '', paragraph(1, 200), paragraph(3, 200)].join('\n');
  const all = documentPassages(longer);
  const row = (all.find((passage) => passage.kind === 'table') as DocumentPassage).number;
  const cut = outlineOf([entry('free', 'First', all.filter((passage) => passage.number <= row).map((passage) => passage.number), both), entry('free', 'Then', all.filter((passage) => passage.number > row).map((passage) => passage.number), freeText)]);
  const alone = await planOutline(demo, { ...brief, document: longer, slideCount: 2 }, { llm: scripted(cut).llm, ...options });
  assert.ok(!alone.slides.some(hasTable), JSON.stringify(alone.slides));
  assert.deepEqual(
    alone.slides.map((slide) => [slide.blocks, slide.arrange?.dense]),
    [
      [['text'], true],
      [['text'], true],
    ],
  );
  for (const fit of planTextFit(demo, alone, options)) assert.equal(fit.holds, fit.words, JSON.stringify(fit));
  const tabled = alone.slides.findIndex((slide) => (slide.passages ?? []).includes(row));
  const set = await writeDeck(demo, alone, { llm: scripted(titles).llm, ...options });
  assert.deepEqual(set.deck.slides[tabled]?.slots.b1Text, fillParagraphs(readDocument(longer, 'en').passages, alone.slides[tabled]?.passages ?? []));
  assert.ok((set.deck.slides[tabled]?.slots.b1Text as string[]).includes('Bikes | 120 | 140'));
});

test('packPassages shares runs evenly and starts them at sections', () => {
  const atoms = [100, 100, 100, 100, 100, 100].map((size, at) => ({ numbers: [at + 1], size, start: at === 2 || at === 4 }));
  assert.deepEqual(packPassages(atoms, 2), [
    [1, 2, 3],
    [4, 5, 6],
  ]);
  // Within a tenth of the even share, a run starts where a section does.
  const sections = [100, 100, 100, 100, 100, 100, 100].map((size, at) => ({ numbers: [at + 1], size, start: at === 3 }));
  assert.deepEqual(packPassages(sections, 2), [
    [1, 2, 3],
    [4, 5, 6, 7],
  ]);
});

// --- writing ---

test('the writer gets each slide\'s text once', async () => {
  const translated: Brief = { ...brief, language: 'fr' };
  const model = scripted(outlineOf([entry('cover', 'Bikes'), entry('free', 'Background', [1, 2, 3], freeText), entry('free', 'Riders', [4, 5], freeText), entry('closing', 'Thanks')], { left_out: [6, 7, 8, 9] }));
  const plan = await planOutline(demo, { ...translated, slideCount: 4 }, { llm: model.llm, ...options });
  const writer = scripted(titles);
  await writeDeck(demo, plan, { llm: writer.llm, ...options });
  const prompt = writer.requests[0]?.prompt ?? '';
  const doc = readDocument(REPORT, 'en');
  const [briefBlock] = prompt.split('\n\nWrite the content');
  assert.ok(!briefBlock?.includes(SENTENCES[0] as string));
  for (const number of [1, 2, 3, 4, 5]) {
    const text = doc.passages[number - 1]?.text as string;
    assert.equal(prompt.split(text).length - 1, 1, `p${number}`);
  }
  for (const number of [6, 7, 8, 9]) assert.ok(!prompt.includes(doc.passages[number - 1]?.text as string), `p${number}`);
  assert.ok(prompt.includes('s2: free (b1: a text). Background\nText of s2 from the document (p1-p3):\n[Background]\n'));
  assert.ok(prompt.includes('Slides with text from the document say that text. Write it in its own words and order'));
  assert.equal(writer.requests[0]?.system, systemPrompt('fr', 'brief', true));
});

test('a slide that holds its passages is filled by code', async () => {
  // Demo: a free text and the text layout. The layout takes one passage of under its room.
  const short = `Background\n${paragraph(0, 150)}\n\nRiders\n${paragraph(4, 200)}\n${paragraph(6, 400)}\n${paragraph(8, 300)}`;
  const model = scripted(outlineOf([entry('text-only', 'Background', [1]), entry('free', 'Riders', [2, 3, 4], freeText)]));
  const plan = await planOutline(demo, { ...brief, document: short, slideCount: 2 }, { llm: model.llm, ...options });
  assert.deepEqual(
    plan.slides.map((slide) => slide.pattern),
    ['text-only', 'free'],
  );
  const writer = scripted(titles);
  const written = await writeDeck(demo, plan, { llm: writer.llm, ...options });
  const doc = readDocument(short, 'en');
  assert.deepEqual(written.deck.slides[0]?.slots.body, fillParagraphs(doc.passages, [1]));
  assert.deepEqual(written.deck.slides[1]?.slots.b1Text, fillParagraphs(doc.passages, [2, 3, 4]));
  const schema = schemaOf(writer.requests[0]);
  assert.equal('body' in (schema.properties.s1?.properties ?? {}), false);
  assert.equal('b1Text' in (schema.properties.s2?.properties ?? {}), false);
  const prompt = writer.requests[0]?.prompt ?? '';
  assert.ok(prompt.includes('s1: text-only. Background\nText of s1 from the document (p1):\n[Background]\n'));
  assert.ok(prompt.includes('A slide whose running text the software sets from the document asks you only for its title and short slots: take them from that text.'));
  assert.deepEqual(documentKept(written.deck), { words: doc.words, kept_words: doc.words });
  assert.deepEqual(written.warnings, []);

  // The fixture, in Hebrew: a free text whose passages are a cut paragraph and a section that changes.
  const hebrew = ['רקע', paragraph(0, 700, HEBREW), '', 'תוכנית', paragraph(3, 200, HEBREW)].join('\n');
  const passages = documentPassages(hebrew);
  assert.deepEqual(
    passages.map((passage) => passage.continues ?? null),
    [null, true, null],
  );
  const he = scripted(outlineOf([entry('free', 'הכול', [1, 2, 3], freeText)]));
  const hePlan = await planOutline(rtl, { ...hebrewBrief, document: hebrew, slideCount: 1 }, { llm: he.llm, ...rtlOptions });
  const heDeck = await writeDeck(rtl, hePlan, { llm: scripted(titles).llm, ...rtlOptions });
  assert.deepEqual(heDeck.deck.slides[0]?.slots.b1Text, [paragraph(0, 700, HEBREW), 'תוכנית', paragraph(3, 200, HEBREW)]);
  assert.deepEqual(documentKept(heDeck.deck)?.kept_words, documentKept(heDeck.deck)?.words);
});

test('not filled by code: a translation, a shorter text, a table, two running texts, too many paragraphs', async () => {
  const filledBy = async (more: Partial<Brief>, slide: ReturnType<typeof entry>, answer: (request: JsonRequest) => unknown = titles, free = true) => {
    const passages = documentPassages(more.document ?? REPORT).map((passage) => passage.number);
    const model = scripted(outlineOf([slide], { left_out: passages.filter((number) => !slide.passages.includes(number)) }));
    const off = free ? {} : { free: false };
    const plan = await planOutline(demo, { ...brief, slideCount: 1, ...more }, { llm: model.llm, ...options, ...off });
    const writer = scripted(answer, answer, answer);
    await writeDeck(demo, plan, { llm: writer.llm, ...options, ...off });
    return schemaOf(writer.requests[0]).properties.s1?.properties ?? {};
  };
  // In another script the writer translates it.
  const translated = await filledBy({ language: 'he' }, entry('free', 'Background', [1, 2], freeText));
  assert.match(String(translated.b1Text?.description), /^The slide's text from the document \(p1-p2\), in its own words: \d+ to \d+ characters in all\. One string for each paragraph of the text; join paragraphs where there are more than the slot takes/);
  // Asked for fewer words, the writer shortens it.
  const shorter = await filledBy({}, entry('free', 'Background', [1, 2], { ...freeText, words: 40 }));
  assert.match(String(shorter.b1Text?.description), /^40 words of the slide's text from the document \(p1-p2\), 210 to \d+ characters: keep its words and leave out what matters least/);
  // A table is written into a table.
  const table = `Costs\n| Item | 2023 | 2024 |\n| Bikes | 120 | 140 |\n| Staff | 300 | 320 |`;
  const tabled = await filledBy({ document: table }, entry('table', 'Costs', [1]), () => ({ s1: { title: 'Costs', table: { header: ['Item', '2023', '2024'], rows: [['Bikes', '120', '140']] } } }));
  assert.ok('table' in tabled);
  // Two running texts are each written, with their share.
  const pair = `Background\n${paragraph(0, 150)}\n${paragraph(3, 150)}`;
  const two = await filledBy({ document: pair }, entry('two-columns', 'Background', [1, 2]), () => ({ s1: { title: 'Background', col1Heading: 'One', col1Text: ['x'], col2Heading: 'Two', col2Text: ['y'] } }));
  assert.match(String(two.col1Text?.description), /^This slot's share of the slide's text from the document \(p1-p2\), in order and in its own words: \d+ to \d+ characters/);
  // More paragraphs than any layout of plain text takes, with free slides off, are joined by the
  // writer. With free slides on, a free slide takes them, and code sets them as they are.
  const lines = `Notes\n${['One.', 'Two.', 'Three.', 'Four.', 'Five.', 'Six.', 'Seven.'].join('\n')}`;
  const many = await filledBy({ document: lines }, entry('text-only', 'Notes', [1]), titles, false);
  assert.match(String(many.body?.description), /^The slide's text from the document \(p1\), in its own words/);
  assert.equal(many.body?.maxItems, 3);
  const freed = await filledBy({ document: lines }, entry('text-only', 'Notes', [1]));
  assert.equal('body' in freed || 'b1Text' in freed, false, JSON.stringify(Object.keys(freed)));
});

test("a list set by code into a layout's list has no bullet of its own there; a free text keeps it", async () => {
  const text = ['What riders said:', '- Bikes are easy to find near the station', '1. Some docks are often full'].join('\n');
  const assets = [{ id: 'bike', kind: 'photo' as const, description: 'A bike at a dock', data: new Uint8Array() }];
  // The writer gives the photo and the title: the list is set by code.
  const photo = (request: JsonRequest): unknown =>
    Object.fromEntries(Object.entries(schemaOf(request).properties).map(([key, slide]) => [key, Object.fromEntries((slide.required ?? []).map((name) => [name, slide.properties[name]?.type === 'object' ? { asset: 'bike' } : 'Riders']))]));
  const model = scripted(outlineOf([{ pattern: 'list-with-photo', purpose: 'Riders', passages: [1, 2], items: 3 }]));
  const plan = await planOutline(demo, { ...brief, imageDensity: 'medium', document: text, slideCount: 1 }, { llm: model.llm, ...options, assets });
  const writer = scripted(photo);
  const written = await writeDeck(demo, plan, { llm: writer.llm, ...options, assets });
  assert.equal('bulletList' in (schemaOf(writer.requests[0]).properties.s1?.properties ?? {}), false);
  // The layout draws its own bullets; a number is the text's, and stays.
  assert.deepEqual(written.deck.slides[0]?.slots.bulletList, ['What riders said:', 'Bikes are easy to find near the station', '1. Some docks are often full']);
  // A free text draws none: there an item keeps its bullet.
  assert.deepEqual(fillParagraphs(documentPassages(text), [1, 2]), text.split('\n'));
  // A lead line and three items are four paragraphs, one more than the layout's list of four takes
  // of such a text. With free slides on, a free slide takes them, and code sets them word for word,
  // each item with its bullet. With free slides off, the layout stays and its writer is told three.
  const four = ['What riders said:', '- Bikes are easy to find near the station', '- The app works well', '1. Some docks are often full'].join('\n');
  const listed = { ...brief, imageDensity: 'medium' as const, document: four, slideCount: 1 };
  const freed = await planOutline(demo, listed, { llm: scripted(outlineOf([{ pattern: 'list-with-photo', purpose: 'Riders', passages: [1, 2], items: 4 }])).llm, ...options, assets });
  assert.equal(freed.slides[0]?.pattern, 'free');
  const freeWriter = scripted(photo);
  const freeSet = await writeDeck(demo, freed, { llm: freeWriter.llm, ...options, assets });
  assert.equal('b1Text' in (schemaOf(freeWriter.requests[0]).properties.s1?.properties ?? {}), false);
  assert.deepEqual(freeSet.deck.slides[0]?.slots.b1Text, four.split('\n'));
  const off = { ...options, free: false, assets };
  const kept = await planOutline(demo, listed, { llm: scripted(outlineOf([{ pattern: 'list-with-photo', purpose: 'Riders', passages: [1, 2], items: 4 }])).llm, ...off });
  assert.equal(kept.slides[0]?.pattern, 'list-with-photo');
  const keptWriter = scripted(photo, photo);
  await writeDeck(demo, kept, { llm: keptWriter.llm, ...off }).catch(() => undefined);
  assert.equal(schemaOf(keptWriter.requests[0]).properties.s1?.properties.bulletList?.maxItems, 3);
  // A word processor's tab after a marker is one space, the bullet dropped or not.
  const pasted = documentPassages(['1.\tOpen the app', '2.\tFind a dock', 'o\tScan the code'].join('\r\n'));
  assert.deepEqual(fillParagraphs(pasted, [1], { bullets: false }), ['1. Open the app', '2. Find a dock', 'Scan the code']);
  assert.deepEqual(fillParagraphs(pasted, [1]), ['1. Open the app', '2. Find a dock', '◦ Scan the code']);
});

test('slot descriptions say the text, not a number of words', () => {
  const context: CatalogContext = { layout: 'adaptive', fields, icons: [], assets: [], imageDensity: 'none', figures: false, textScale: 1 };
  const catalog = buildCatalog(demoManifest, context).patterns;
  const pattern = (id: string) => catalog.find((one) => one.def.id === id) as UsablePattern;
  const describe = (id: string, need: number, paragraphs: number, words?: number) =>
    withPassages(pattern(id), need, context, { label: 'p4-p5', shape: { paragraphs }, words }).slots.flatMap((planned) => (planned.slot.description?.includes('p4-p5') ? [`${planned.slot.name}: ${planned.slot.description}`] : []));
  // One slot that holds it: from nine tenths of the text to a little over it, never over the room.
  assert.deepEqual(describe('text-only', 200, 1), [
    "body: The slide's text from the document (p4-p5), in its own words: 180 to 210 characters in all. One string for each paragraph of the text; join paragraphs where there are more than the slot takes",
  ]);
  // Short: shortened to what the slide holds.
  assert.deepEqual(describe('text-only', 600, 3), [
    "body: The slide's text from the document (p4-p5) shortened to 220 to 240 characters, the most this slide holds: keep its words and leave out what matters least. One string for each paragraph of the text; join paragraphs where there are more than the slot takes",
  ]);
  // Several slots: each its share, in order.
  assert.deepEqual(describe('two-columns', 400, 2), [
    "col1Text: This slot's share of the slide's text from the document (p4-p5), in order and in its own words: 180 to 210 characters. One string for each paragraph of the text; join paragraphs where there are more than the slot takes",
    "col2Text: This slot's share of the slide's text from the document (p4-p5), in order and in its own words: 180 to 210 characters. One string for each paragraph of the text; join paragraphs where there are more than the slot takes",
  ]);
  // Fewer words asked for: that many of its words.
  assert.deepEqual(describe('text-only', 240, 1, 40), [
    "body: 40 words of the slide's text from the document (p4-p5), 210 to 250 characters: keep its words and leave out what matters least. One string for each paragraph of the text; join paragraphs where there are more than the slot takes",
  ]);
});

test('a slide that sums up its text is sent back once', async () => {
  // Two passages of twenty words each, on a layout of two running texts that holds them.
  const words = (from: number) => Array.from({ length: 20 }, (_, at) => `w${from + at}x`);
  const first = words(0);
  const second = words(100);
  const document = `Notes\n${first.join(' ')}.\n${second.join(' ')}.`;
  const model = scripted(outlineOf([entry('two-columns', 'Notes', [1, 2])]));
  const plan = await planOutline(demo, { ...brief, document, slideCount: 1 }, { llm: model.llm, ...options });
  assert.equal(plan.slides[0]?.pattern, 'two-columns');
  const answer = (one: number, two: number) => ({ s1: { title: 'Notes', col1Heading: 'One', col1Text: [`${first.slice(0, one).join(' ')}.`], col2Heading: 'Two', col2Text: [`${second.slice(0, two).join(' ')}.`] } });
  // 28 of 40 words, twice: one repair, then the warning.
  const summing = scripted(answer(14, 14), answer(14, 14));
  const summed = await writeDeck(demo, plan, { llm: summing.llm, ...options });
  assert.equal(summing.requests.length, 2);
  assert.ok(
    summing.requests[1]?.prompt.includes(
      '- s1.col1Text: keeps 70% of the words of its text from the document (p1-p2), and the slide has room for nearly all of it: write that text in its own words and order, shortened only where a limit makes you',
    ),
    String(summing.requests[1]?.prompt.slice(-700)),
  );
  assert.deepEqual(
    withoutCodes(summed.warnings.filter((issue) => /keeps \d+%/.test(issue.message))),
    [{ severity: 'warning', slide: 1, message: 'the slide keeps 70% of the words of its text from the document (p1-p2), though it has room for nearly all of it' }],
  );
  // 36 of 40 words: no repair.
  const close = scripted(answer(18, 18));
  const kept = await writeDeck(demo, plan, { llm: close.llm, ...options });
  assert.equal(close.requests.length, 1);
  assert.deepEqual(kept.warnings, []);
  // A slide asked to be shorter is not held to the words of its text.
  const shortened = scripted(outlineOf([entry('two-columns', 'Notes', [1, 2], { words: 20 })]));
  const shortPlan = await planOutline(demo, { ...brief, document, slideCount: 1 }, { llm: shortened.llm, ...options });
  const brief20 = scripted(answer(10, 10));
  await writeDeck(demo, shortPlan, { llm: brief20.llm, ...options });
  assert.equal(brief20.requests.length, 1);
});

test('Hebrew prefixes count as kept words', () => {
  assert.deepEqual(tokensOf('The CITY\u200F, 2019!'), ['the', 'city', '2019']);
  // A slide that writes the words of its text with a prefix more or one less keeps them, each once.
  const source = 'הלקוחות מקבלים מענה מהיר בסניפים החדשים';
  const kept = (written: string) => documentKept({ document: source, slides: [{ pattern: 'text-only', passages: [1], slots: { body: [written] } }] });
  assert.deepEqual(kept('והלקוחות מקבלים למענה מהיר סניפים החדשים'), { words: 6, kept_words: 6 });
  assert.deepEqual(kept('לקוחות מקבלים מענה'), { words: 6, kept_words: 3 });
  // Two prefixes more are another word, and a word written once is kept once.
  assert.deepEqual(kept('ושהלקוחות מקבלים מקבלים'), { words: 6, kept_words: 1 });
});

test('the deck keeps the document and each slide its passages', async () => {
  const model = scripted(outlineOf([entry('cover', 'Bikes'), entry('free', 'Background', [1, 2, 3, 4], freeText), entry('free', 'Plans', [5, 6, 7, 8], freeText), entry('closing', 'Thanks')], { left_out: [9] }));
  const plan = await planOutline(demo, { ...brief, slideCount: 4 }, { llm: model.llm, ...options });
  const written = await writeDeck(demo, plan, { llm: scripted(titles).llm, ...options });
  assert.equal(Object.keys(written.deck).at(-1), 'document');
  assert.equal(written.deck.document, REPORT);
  assert.deepEqual(
    written.deck.slides.map((slide) => slide.passages ?? null),
    [null, [1, 2, 3, 4], [5, 6, 7, 8], null],
  );
  assert.equal(documentKept(written.deck)?.left_out, 'p9');
});

test('chart numbers stated in the document are not flagged', () => {
  const deck: DeckSpec = {
    version: 1,
    brief: 'The bike share',
    slides: [{ pattern: 'chart-native', slots: { title: 'Trips', chart: { type: 'column', categories: ['2023', '2024'], series: [{ name: 'Trips', values: [1.1, 1.2] }] } } }],
  };
  assert.equal(unstatedChartValues(demoManifest, deck).length, 1);
  assert.deepEqual(unstatedChartValues(demoManifest, { ...deck, document: 'Riders took 1.1 million trips in 2023 and 1.2 million in 2024.' }), []);
});

test('a plan made from a document says what each slide carries before writing', async () => {
  // An opening slide asked again for passages and kept: the brief asks for it.
  const answer = outlineOf([entry('cover', 'Bikes'), entry('free', 'All', [1, 2, 3, 4, 5, 6, 7, 8, 9], freeText)]);
  const model = scripted(answer, answer);
  const plan: DeckPlan = await planOutline(demo, { ...brief, slideCount: 2 }, { llm: model.llm, ...options });
  const doc = readDocument(REPORT, 'en');
  const fits = planTextFit(demo, plan, options);
  assert.equal(fits.length, 1);
  assert.equal(fits[0]?.slide, 2);
  assert.equal(fits[0]?.passages, 'p1-p9');
  assert.equal(fits[0]?.words, doc.words);
  assert.equal(fits[0]?.dense, true);
});

// --- the language of the document ---

test('the language of a document is read as a tag, and one language is its primary subtag', () => {
  assert.equal(readLanguageTag('en'), 'en');
  assert.equal(readLanguageTag(' he-IL '), 'he-IL');
  for (const none of ['English', 'und', 'mul', '', 'e', 42, null, undefined]) assert.equal(readLanguageTag(none), undefined, String(none));
  assert.equal(sameLanguage('en-GB', 'en-US'), true);
  assert.equal(sameLanguage('iw', 'he'), true);
  assert.equal(sameLanguage('en', 'fr'), false);
});

test('the language of the document decides whether code sets it into a slide', async () => {
  const short = `Background\n${paragraph(0, 150)}\n\nRiders\n${paragraph(4, 200)}`;
  const run = async (language: string, answered?: unknown) => {
    const model = scripted(outlineOf([entry('free', 'Background', [1, 2], freeText)], answered === undefined ? {} : { document_language: answered }));
    const plan = await planOutline(demo, { ...brief, language, document: short, slideCount: 1 }, { llm: model.llm, ...options });
    const writer = scripted(titles, titles, titles);
    const written = await writeDeck(demo, plan, { llm: writer.llm, ...options });
    return { plan, deck: written.deck, filled: !('b1Text' in (schemaOf(writer.requests[0]).properties.s1?.properties ?? {})) };
  };
  // An English document in a French deck is the writer's to translate, though both are in one script.
  const french = await run('fr', 'en');
  assert.equal(french.plan.documentLanguage, 'en');
  assert.equal(french.filled, false);
  assert.deepEqual(Object.keys(french.deck).slice(-2), ['documentLanguage', 'document']);
  assert.equal(french.deck.documentLanguage, 'en');
  // In English of any region, code sets it as it is.
  const english = await run('en', 'en-GB');
  assert.equal(english.filled, true);
  assert.equal(english.deck.documentLanguage, 'en-GB');
  // An answer that names no language, or no tag, leaves the script to decide, as for a plan stored before it was asked.
  const unsaid = await run('fr');
  assert.equal(unsaid.plan.documentLanguage, undefined);
  assert.equal(unsaid.filled, true);
  assert.equal('documentLanguage' in unsaid.deck, false);
  assert.equal((await run('fr', 'English')).plan.documentLanguage, undefined);
  assert.equal((await run('he')).filled, false);
});

test("a language the text's own script belies is not kept, and the script decides", async () => {
  const memo = HEBREW.slice(0, 3).join('\n\n');
  assert.equal(readLanguageTag('en', memo), undefined);
  assert.equal(readLanguageTag('he-IL', memo), 'he-IL');
  // A language whose script is not known is taken at its word.
  assert.equal(readLanguageTag('ja', memo), 'ja');
  const model = scripted(outlineOf([entry('free', 'Memo', [1], freeText)], { document_language: 'en' }));
  const plan = await planOutline(demo, { ...brief, document: memo, slideCount: 1 }, { llm: model.llm, ...options });
  assert.equal(plan.documentLanguage, undefined);
  // A Hebrew text in an English deck is the writer's to translate, also for a plan stored with the slip.
  for (const given of [plan, { ...plan, documentLanguage: 'en' }]) {
    const writer = scripted(titles, titles);
    const written = await writeDeck(demo, given, { llm: writer.llm, ...options });
    assert.ok('b1Text' in (schemaOf(writer.requests[0]).properties.s1?.properties ?? {}), JSON.stringify(given.documentLanguage));
    // Its words are not kept, and the deck says it is translated.
    assert.deepEqual(documentKept(written.deck), { words: tokensOf(memo).length, translated: true });
  }
});

test('a Hebrew text full of English terms is still Hebrew where the planner says so', async () => {
  // More Latin letters than Hebrew ones: the script most letters are in is not the text's language.
  const memo = [
    'הצוות העביר את ה-Kubernetes cluster ל-Azure, שדרג את PostgreSQL ל-version 16 והטמיע Microsoft Entra ID לכל העובדים.',
    'ברבעון הבא נעביר את ה-Jenkins pipelines ל-GitHub Actions ונחליף את Splunk ב-Azure Monitor.',
  ].join('\n\n');
  assert.equal(scriptOf(memo), 'latin');
  assert.equal(readLanguageTag('he', memo), 'he');
  assert.equal(readDocument(memo, 'he', 'he').same, true);
  // So it is set word for word into a Hebrew deck, and its words are counted as kept.
  const model = scripted(outlineOf([entry('free', 'תשתיות', [1, 2], freeText)], { document_language: 'he' }));
  const plan = await planOutline(rtl, { ...hebrewBrief, document: memo, slideCount: 1 }, { llm: model.llm, ...rtlOptions });
  assert.equal(plan.documentLanguage, 'he');
  const written = await writeDeck(rtl, plan, { llm: scripted(titles).llm, ...rtlOptions });
  const doc = readDocument(memo, 'he', 'he');
  assert.deepEqual(documentKept(written.deck), { words: doc.words, kept_words: doc.words });
});

// --- changes ---

/** A plan of five slides made from the report, as stored: its first text on a layout too small for it, p7 to p9 on no slide. */
const storedPlan = (): DeckPlan => ({
  version: 1,
  title: 'Bikes',
  brief: { ...brief, slideCount: 5 },
  slides: [
    { pattern: 'cover', purpose: 'Bikes' },
    { pattern: 'text-only', purpose: 'Background', passages: [1, 2] },
    { pattern: 'free', purpose: 'Riders', passages: [3, 4], blocks: ['text'], arrange: { lead: 'none', dense: true } },
    { pattern: 'free', purpose: 'Costs', passages: [5, 6], blocks: ['text'], arrange: { lead: 'none', dense: true } },
    { pattern: 'closing', purpose: 'Thanks' },
  ],
  documentLanguage: 'en',
});
/** The answer to a change of a plan made from a document. */
const changeOf = (slides: unknown[], more: Record<string, unknown> = {}) => ({ summary: 'Done.', language: 'en', title: 'Bikes', left_out: [], asked_slides: 0, document_language: 'en', slides, ...more });

test('editPlan shows the passages and fits only the slides whose passages moved', async () => {
  const plan = storedPlan();
  const model = scripted(changeOf([entry('cover', 'Bikes'), entry('text-only', 'Background', [1, 2]), entry('free', 'Riders', [3, 4, 5], freeText), entry('free', 'Costs', [6], freeText), entry('closing', 'Thanks')]));
  const { plan: next } = await editPlan(demo, plan, 'Move the passage on prices to the riders slide', { llm: model.llm, ...options });
  const request = model.requests[0] as JsonRequest;
  const prompt = request.prompt;
  // The document comes after the brief, with the passages no slide carries marked.
  const block = prompt.indexOf('The document the deck is made from, in 9 numbered passages (');
  assert.ok(prompt.indexOf('Brief:') < block && block < prompt.indexOf('The plan as it is now'), prompt.slice(0, 300));
  assert.match(prompt, /\np7 \(\d+, on no slide\): /);
  assert.doesNotMatch(prompt, /\np6 \(\d+, on no slide\)/);
  // Each line says the passages of its slide and their length, what is cut and the most a slide set dense holds.
  const doc = readDocument(REPORT, 'en');
  const chars = (numbers: number[]) => numbers.reduce((sum, number) => sum + (doc.passages[number - 1]?.chars ?? 0), 0);
  assert.ok(prompt.includes(`\n2. text-only (p1-p2, ${chars([1, 2])} characters; cut to fit): Background\n`), prompt.slice(prompt.indexOf('The plan as it is now'), prompt.indexOf('Requested change')));
  assert.match(prompt, new RegExp(`\\n3\\. free \\(a text; [^)]*p3-p4, ${chars([3, 4])} characters; at most about 720 words\\): Riders\\n`));
  for (const rule of DOCUMENT_CHANGE_RULES) assert.ok(prompt.includes(`\n- ${rule}`), rule);
  assert.ok(prompt.indexOf('- A slide given words holds them all') < prompt.indexOf("- The plan carries the document's own text."));
  // The answer names each slide's passages, those the request leaves out, the slides it asks for and the language of the document.
  const schema = schemaOf(request);
  assert.deepEqual(schema.required, ['summary', 'language', 'title', 'slides', 'left_out', 'asked_slides', 'document_language']);
  assert.ok((schema.properties.slides as unknown as { items: Schema }).items.required?.includes('passages'));
  assert.equal(schema.properties.asked_slides?.description, 'The number of slides the request asks the plan to have ("keep it to 2 slides"); 0 when it names none.');
  assert.equal(request.system, systemPrompt('en', 'brief', true));
  // 400 + 160 a slide of five and ten more, 12 for each and for each run of parts, 8 a slide for its passages, 2 a passage and 40.
  assert.equal(request.maxOutputTokens, 400 + 160 * 15 + 12 * (15 + 2) + 8 * 15 + 2 * 9 + 40);
  // The slides whose passages stayed are as they were, the one that is cut too; the two that changed are fitted again, on their slides.
  assert.deepEqual([next.slides[0], next.slides[1], next.slides[4]], [plan.slides[0], plan.slides[1], plan.slides[4]]);
  assert.deepEqual(
    next.slides.map((slide) => slide.passages ?? []),
    [[], [1, 2], [3, 4, 5], [6], []],
  );
  assert.deepEqual(next.slides[2]?.arrange, { lead: 'none', dense: true });
  assert.equal(next.documentLanguage, 'en');
  assert.equal(next.brief.slideCount, undefined);
  // An answer that names no language keeps the one the plan has.
  const unsaid = scripted({ ...changeOf(plan.slides.map((slide) => entry(slide.pattern, slide.purpose, slide.passages ?? [], slide.blocks ? freeText : {}))), document_language: undefined });
  assert.equal((await editPlan(demo, plan, 'Shorter titles', { llm: unsaid.llm, ...options })).plan.documentLanguage, 'en');
});

test('editPlan puts back a passage the answer dropped, and asks again for an answer that places none', async () => {
  const model = scripted(changeOf([entry('cover', 'Bikes'), entry('text-only', 'Background', [1, 2]), entry('free', 'Riders', [3], freeText), entry('free', 'Costs', [5, 6], freeText), entry('closing', 'Thanks')]));
  const { plan: next } = await editPlan(demo, storedPlan(), 'A shorter title on the riders slide', { llm: model.llm, ...options });
  assert.equal(model.requests.length, 1);
  // p4 goes back next to p3. p7 to p9 were on no slide, and stay so.
  assert.deepEqual(
    next.slides.map((slide) => slide.passages ?? []),
    [[], [1, 2], [3, 4], [5, 6], []],
  );

  // No passage on any slide: asked again, then the passages the plan had are shared over the slides of content.
  const bare = changeOf([entry('cover', 'Bikes'), entry('free', 'One', [], freeText), entry('free', 'Two', [], freeText), entry('closing', 'Thanks')]);
  const twice = scripted(bare, bare);
  const packed = await editPlan(demo, storedPlan(), 'Two slides of content', { llm: twice.llm, ...options });
  assert.equal(twice.requests.length, 2);
  assert.ok(
    twice.requests[1]?.prompt.endsWith('Your previous answer could not be used: passages p1-p9 are on no slide and not in "left_out". Answer again, with every passage on one slide or in "left_out".'),
    String(twice.requests[1]?.prompt.slice(-300)),
  );
  assert.equal(packed.plan.slides.length, 4);
  assert.deepEqual(
    packed.plan.slides.flatMap((slide) => slide.passages ?? []),
    [1, 2, 3, 4, 5, 6],
  );
});

test('editPlan never adds slides for a longer slide', async () => {
  // Two slides of four: each stays one slide. Only where the answer names the number of slides are the passages shared again.
  const four = scripted(outlineOf([entry('free', 'א', [1, 2], freeText), entry('free', 'ב', [3, 4], freeText), entry('free', 'ג', [5, 6], freeText), entry('free', 'ד', [7, 8, 9], freeText)]));
  const plan = await planOutline(rtl, { ...hebrewBrief, slideCount: 4 }, { llm: four.llm, ...rtlOptions });
  const two = (asked: number) => changeOf([entry('free', 'הרוב', [1, 2, 3, 4, 5, 6, 7, 8], freeText), entry('free', 'הסוף', [9], freeText)], { language: 'he', asked_slides: asked, document_language: 'he' });
  const kept = await editPlan(rtl, plan, 'Keep it to 2 slides', { llm: scripted(two(0)).llm, ...rtlOptions });
  assert.deepEqual(
    kept.plan.slides.map((slide) => slide.passages),
    [[1, 2, 3, 4, 5, 6, 7, 8], [9]],
  );
  assert.ok(kept.plan.slides.every((slide) => slide.arrange?.dense === true && !/\(\d+\/\d+\)$/.test(slide.purpose)));
  const [over] = planTextFit(rtl, kept.plan, rtlOptions);
  assert.ok(over && over.holds < over.words, JSON.stringify(over));
  const shared = await editPlan(rtl, plan, 'Keep it to 2 slides', { llm: scripted(two(2)).llm, ...rtlOptions });
  const lists = shared.plan.slides.map((slide) => slide.passages ?? []);
  assert.equal(lists.length, 2);
  assert.deepEqual(lists.flat(), [1, 2, 3, 4, 5, 6, 7, 8, 9]);
  assert.ok(lists.every((list) => list.length > 0 && list.every((number, at) => at === 0 || number === (list[at - 1] as number) + 1)), JSON.stringify(lists));
  for (const fit of planTextFit(rtl, shared.plan, rtlOptions)) assert.equal(fit.holds, fit.words, JSON.stringify(fit));
});

test('editPlan shares the passages over every slide when the request names a number of slides, kept slides too, in order', async () => {
  const text = Array.from({ length: 12 }, (_, at) => paragraph(at, 420, HEBREW)).join('\n\n');
  const parts = [
    [1, 2, 3],
    [4, 5, 6],
    [7, 8, 9],
    [10, 11, 12],
  ];
  const four = outlineOf(parts.map((own, at) => entry('free', `חלק ${at + 1}`, own, freeText)));
  const plan = await planOutline(rtl, { ...hebrewBrief, document: text, slideCount: 4 }, { llm: scripted(four).llm, ...rtlOptions });
  const asked = (slides: unknown[], count: number) => changeOf(slides, { language: 'he', asked_slides: count, document_language: 'he' });
  // Slide 1 as it was, and the rest on slide 2: shared over both.
  const kept = asked([entry('free', 'חלק 1', [1, 2, 3], freeText), entry('free', 'השאר', [4, 5, 6, 7, 8, 9, 10, 11, 12], freeText)], 2);
  const two = await editPlan(rtl, plan, 'Keep it to 2 slides', { llm: scripted(kept).llm, ...rtlOptions });
  assert.deepEqual(
    two.plan.slides.map((slide) => slide.passages),
    [
      [1, 2, 3, 4, 5, 6],
      [7, 8, 9, 10, 11, 12],
    ],
  );
  // Slide 2 as it was, between two that changed: no passage is moved across it.
  const middle = asked([entry('free', 'חלק 1 בקצרה', [1, 2, 3], freeText), entry('free', 'חלק 2', [4, 5, 6], freeText), entry('free', 'חלקים 3 ו-4', [7, 8, 9, 10, 11, 12], freeText)], 3);
  const three = await editPlan(rtl, plan, 'Keep it to 3 slides', { llm: scripted(middle).llm, ...rtlOptions });
  assert.deepEqual(
    three.plan.slides.map((slide) => slide.passages),
    [
      [1, 2, 3, 4],
      [5, 6, 7, 8],
      [9, 10, 11, 12],
    ],
  );
  for (const fit of planTextFit(rtl, three.plan, rtlOptions)) assert.equal(fit.holds, fit.words, JSON.stringify(fit));

  // Slide 1 kept as it was, asked to be shorter than its passages: it keeps them, and the slides after it share theirs again.
  const short = outlineOf(parts.map((own, at) => entry('free', `חלק ${at + 1}`, own, at === 0 ? { ...freeText, words: 40 } : freeText)));
  const shortPlan = await planOutline(rtl, { ...hebrewBrief, document: text, slideCount: 4 }, { llm: scripted(short).llm, ...rtlOptions });
  const shorter = asked([entry('free', 'חלק 1', [1, 2, 3], { ...freeText, words: 40 }), entry('free', 'חלק 2', [4, 5, 6], freeText), entry('free', 'חלקים 3 ו-4', [7, 8, 9, 10, 11, 12], freeText)], 3);
  const fenced = await editPlan(rtl, shortPlan, 'Keep it to 3 slides', { llm: scripted(shorter).llm, ...rtlOptions });
  assert.deepEqual(
    fenced.plan.slides.map((slide) => [slide.passages, slide.words]),
    [
      [[1, 2, 3], 40],
      [[4, 5, 6, 7], undefined],
      [[8, 9, 10, 11, 12], undefined],
    ],
  );

  // Two slides answered with an opening slide are asked again, once.
  const opened = changeOf([entry('cover', 'Bikes'), entry('free', 'All', [1, 2, 3, 4, 5, 6], freeText)], { asked_slides: 2 });
  const halves = changeOf([entry('free', 'Background', [1, 2, 3], freeText), entry('free', 'Costs', [4, 5, 6], freeText)], { asked_slides: 2 });
  const again = scripted(opened, halves);
  const next = await editPlan(demo, storedPlan(), 'Keep it to 2 slides', { llm: again.llm, ...options });
  assert.equal(again.requests.length, 2);
  assert.ok(
    again.requests[1]?.prompt.endsWith(
      'Your previous answer could not be used: slide 1 carries no passage, and a deck of three slides or fewer carries the document on every slide unless the request asks for an opening or a closing slide. Answer again, with every passage on one slide or in "left_out".',
    ),
    String(again.requests[1]?.prompt.slice(-300)),
  );
  assert.deepEqual(
    next.plan.slides.map((slide) => slide.passages),
    [
      [1, 2, 3],
      [4, 5, 6],
    ],
  );
});

/** A deck of four slides made from the report: p1-p3 and p4-p6 on free texts set by code, p7 to p9 left out. */
async function reportDeck(): Promise<DeckSpec> {
  const model = scripted(outlineOf([entry('cover', 'Bikes'), entry('free', 'Background', [1, 2, 3], freeText), entry('free', 'Riders and costs', [4, 5, 6], freeText), entry('closing', 'Thanks')], { left_out: [7, 8, 9], document_language: 'en' }));
  const plan = await planOutline(demo, { ...brief, slideCount: 4 }, { llm: model.llm, ...options });
  return (await writeDeck(demo, plan, { llm: scripted(titles).llm, ...options })).deck;
}
/** The plan of a change to a deck, its slides as a model gives them. */
const editOf = (slides: Array<Record<string, unknown>>, more: Record<string, unknown> = {}) => ({
  summary: 'Done.',
  language: 'en',
  title: 'Bikes',
  text_size: 'normal',
  fields,
  slides: slides.map((slide) => ({ change: '', images: false, parts: [], shows_main: 'auto', side_by_side: 0, ...slide })),
  ...more,
});

test('planEdit shows the document once, and the writer the text of each slide it writes', async () => {
  const deck = await reportDeck();
  const doc = readDocument(REPORT, 'en');
  const p7 = doc.passages[6]?.text as string;
  const withLayout: DeckSpec = { ...deck, slides: [...deck.slides.slice(0, 3), { pattern: 'text-only', passages: [7], slots: { title: 'Plans', body: [p7] } }, deck.slides[3] as DeckSpec['slides'][number]] };
  const model = scripted(editOf([{ from: 1, pattern: 'cover' }, { from: 2, pattern: 'free' }, { from: 3, pattern: 'free' }, { from: 4, pattern: 'text-only', change: 'A shorter title' }, { from: 5, pattern: 'closing' }]), titles);
  await planEdit(demo, withLayout, 'A shorter title on slide 4', { llm: model.llm, ...options });
  const plan = model.requests[0]?.prompt ?? '';
  // The document once, after the brief, and each slide that carries passages by their numbers, its long texts cut short.
  assert.ok(plan.indexOf('The deck was written from this brief:') < plan.indexOf('The document the deck is made from, in 9 numbered passages') && plan.indexOf('The document the deck is made from') < plan.indexOf('The deck as it is now'));
  assert.equal(plan.split(doc.passages[0]?.text as string).length - 1, 1);
  assert.equal(plan.split(p7).length - 1, 1);
  assert.ok(plan.includes(`\n4. [text-only, p7] title: "Plans"; body: ${JSON.stringify(`${p7.slice(0, 157).trimEnd()}...`)}\n`), plan.slice(plan.indexOf('The deck as it is now')));
  assert.match(plan, /\n2\. \[free: a text; [^\]]*p1-p3; \d+ words; at most about 720 words\] /);
  assert.match(plan, /\np8 \(\d+, on no slide\): /);
  assert.ok(plan.includes(`\n- ${DOCUMENT_EDIT_RULE}\n`));
  assert.equal(model.requests[0]?.system, systemPrompt('en', 'deck', true));
  const slide = (schemaOf(model.requests[0]).properties.slides as unknown as { items: Schema }).items;
  assert.ok(!slide.required?.includes('passages'));
  assert.equal(slide.properties.passages?.description, 'The numbers of the passages of the document this slide carries after the change. Leave it out to keep those it has; [] when it gives them all up.');
  // The writer is shown the text of the slide it writes, and only that.
  const writer = model.requests[1]?.prompt ?? '';
  assert.ok(writer.includes(`s4 (text-only). Change: A shorter title\nText of s4 from the document (p7):\n[Plans]\n${p7}\n`), writer);
  assert.ok(!writer.includes(doc.passages[0]?.text as string));
  assert.ok(writer.includes('Slides with text from the document say that text.'));
  assert.equal(model.requests[1]?.system, systemPrompt('en', 'deck', true));
  // It keeps its layout, which is a little short of its text, as it was when it was planned.
  assert.match(String(schemaOf(model.requests[1]).properties.s4?.properties.body?.description), /^The slide's text from the document \(p7\) shortened to \d+ to \d+ characters, the most this slide holds/);
});

test('planEdit: named passages move, inherited ones stay, and a slide that loses some is written', async () => {
  const deck = await reportDeck();
  const doc = readDocument(REPORT, 'en');
  const model = scripted(editOf([{ from: 1, pattern: 'cover' }, { from: 2, pattern: 'free', change: 'Add the electric bikes', passages: [1, 2, 3, 4] }, { from: 3, pattern: 'free' }, { from: 4, pattern: 'closing' }]), titles);
  const edited = await planEdit(demo, deck, 'Move the part on electric bikes to slide 2', { llm: model.llm, ...options });
  assert.deepEqual(
    edited.deck.slides.map((slide) => slide.passages ?? null),
    [null, [1, 2, 3, 4], [5, 6], null],
  );
  assert.deepEqual(edited.rewritten, [2, 3]);
  // Slide 3 lost p4: it is written from its text, which code sets as it is, and so is the text of slide 2.
  assert.ok((model.requests[1]?.prompt ?? '').includes('s3 (free: b1: a text). Change: Write the slide from its text from the document.\nText of s3 from the document (p5-p6):\n[Costs]\n'));
  assert.deepEqual(edited.deck.slides[1]?.slots.b1Text, fillParagraphs(doc.passages, [1, 2, 3, 4]));
  assert.deepEqual(edited.deck.slides[2]?.slots.b1Text, fillParagraphs(doc.passages, [5, 6]));
  assert.ok((model.requests[1]?.prompt ?? '').includes('A slide whose running text the software sets from the document asks you only for its title and short slots'));
  assert.deepEqual([edited.deck.slides[0], edited.deck.slides[3]], [deck.slides[0], deck.slides[3]]);
  assert.equal(edited.deck.document, REPORT);
  assert.equal(edited.deck.documentLanguage, 'en');
  const kept = documentKept(edited.deck);
  assert.equal(kept?.kept_words, doc.words - tokensOf(fillParagraphs(doc.passages, [7, 8, 9]).join(' ')).length);
  assert.equal(kept?.left_out, 'p7-p9');
  assert.deepEqual(edited.warnings, []);
});

test('planEdit never adds slides for a longer slide', async () => {
  const deck = await reportDeck();
  const doc = readDocument(REPORT, 'en');
  const model = scripted(editOf([{ from: 1, pattern: 'cover' }, { from: 2, pattern: 'free', change: 'Add the plans', passages: [1, 2, 3, 7, 8] }, { from: 3, pattern: 'free' }, { from: 4, pattern: 'closing' }]), titles);
  const edited = await planEdit(demo, deck, 'Add the part about the plans to slide 2', { llm: model.llm, ...options });
  assert.equal(edited.deck.slides.length, 4);
  assert.deepEqual(edited.deck.slides[1]?.passages, [1, 2, 3, 7, 8]);
  assert.equal(edited.deck.slides[1]?.arrange?.dense, true);
  assert.deepEqual(edited.deck.slides[1]?.slots.b1Text, fillParagraphs(doc.passages, [1, 2, 3, 7, 8]));
  assert.deepEqual(edited.rewritten, [2]);
  assert.deepEqual(edited.warnings, []);
  assert.equal(documentKept(edited.deck)?.left_out, 'p9');
  // Its text is set by code: the writer is asked for the rest of the slide alone.
  assert.equal('b1Text' in (schemaOf(model.requests[1]).properties.s2?.properties ?? {}), false);
});

test('planEdit says before writing that a slide would be cut', async () => {
  const long = hebrewDocument(7000);
  const doc = readDocument(long, 'he');
  const every = doc.passages.map((passage) => passage.number);
  const half = Math.ceil(every.length / 2);
  const slide = (passages: number[]): DeckSpec['slides'][number] => ({ pattern: 'free', blocks: ['text'], arrange: { dense: true }, passages, slots: { title: 'דוח', b1Text: ['טקסט'] } });
  const deck: DeckSpec = { version: 1, language: 'he', direction: 'rtl', title: 'דוח', slides: [slide(every.slice(0, half)), slide(every.slice(half))], documentLanguage: 'he', document: long };
  const all = { from: 1, pattern: 'free', change: 'All of it here', passages: every };
  const plan = editOf([all, { from: 2, pattern: 'free', change: 'Another title', passages: [] }], { language: 'he', fields: {} });
  const stopped = scripted(plan);
  const error = await planEdit(rtl, deck, 'Put all of it on slide 1', { llm: stopped.llm, ...rtlOptions, stopBeforeCut: true }).then(
    () => undefined,
    (thrown: unknown) => thrown,
  );
  assert.ok(error instanceof TextCutError, String(error));
  assert.equal(stopped.requests.length, 1);
  assert.match(error.message, /^Nothing was changed: slide 1 would hold about \d+ of its \d+ words \(text at 7 pt\)\. Ask for more slides or a shorter text, or accept that the rest is cut\.$/);
  assert.equal(error.fits.length, 1);
  assert.deepEqual({ ...error.fits[0], holds: 0 }, { slide: 1, words: doc.words, holds: 0, size: 7, dense: true, passages: `p1-p${every.length}` });
  assert.ok((error.fits[0]?.holds ?? 0) < doc.words);
  // Without the stop, the slide is written and the deck is told that it is cut.
  const writer = (request: JsonRequest) => ({ s1: { title: 'הכול', b1Text: ['טקסט'], ...(titles(request) as { s1?: object }).s1 }, s2: { title: 'עוד' } });
  const written = await planEdit(rtl, deck, 'Put all of it on slide 1', { llm: scripted(plan, writer, writer, writer).llm, ...rtlOptions });
  const cut = written.warnings.find((issue) => issue.message.startsWith('the text from the document'));
  assert.match(cut?.message ?? '', new RegExp(`^the text from the document for this slide \\(p1-p${every.length}\\) is ${needOf(doc, every, 'he')} characters, and the slide holds about \\d+ even with its text set small, at 7 pt: it is cut to fit\\. Ask for more slides to keep it whole$`));
  assert.equal(cut?.slide, 1);
  // Slide 2 gave up its passages: it carries none, and keeps no length of the text that went.
  assert.equal(written.deck.slides[1]?.passages, undefined);
  assert.ok(!written.warnings.some((issue) => /words were asked for this slide/.test(issue.message)), JSON.stringify(written.warnings));
});

test('planEdit: a table moved to a slide with no table is running text there, set by code', async () => {
  const text = ['Costs', paragraph(7, 150), '| Item | 2023 | 2024 |', '|---|---|---|', '| Bikes | 120 | 140 |', '| Staff | 300 | 320 |', '', 'Plans', paragraph(11, 150)].join('\n');
  const doc = readDocument(text, 'en');
  assert.deepEqual(
    doc.passages.map((passage) => passage.kind),
    ['text', 'table', 'text'],
  );
  const model = scripted(outlineOf([entry('free', 'Costs', [1], freeText), entry('table', 'Figures', [2]), entry('free', 'Plans', [3], freeText)], { document_language: 'en' }));
  const plan = await planOutline(demo, { ...brief, document: text, slideCount: 3 }, { llm: model.llm, ...options });
  const tabled = (request: JsonRequest) => {
    const answer = titles(request) as Record<string, Record<string, unknown>>;
    if (answer.s2) answer.s2.table = { header: ['Item', '2023', '2024'], rows: [['Bikes', '120', '140']] };
    return answer;
  };
  const { deck } = await writeDeck(demo, plan, { llm: scripted(tabled).llm, ...options });
  const edit = scripted(editOf([{ from: 1, pattern: 'free', change: 'Add the figures', passages: [1, 2] }, { from: 3, pattern: 'free' }]), titles);
  const edited = await planEdit(demo, deck, 'Put the figures on the first slide, under the text', { llm: edit.llm, ...options });
  assert.deepEqual(edited.deck.slides[0]?.passages, [1, 2]);
  assert.deepEqual(edited.deck.slides[0]?.slots.b1Text, fillParagraphs(doc.passages, [1, 2]));
  assert.equal('b1Text' in (schemaOf(edit.requests[1]).properties.s1?.properties ?? {}), false);
});

test('planEdit sets a text as it is only where the document is in the language of the deck', async () => {
  const deck = await reportDeck();
  const run = async (given: DeckSpec) => {
    const model = scripted(editOf([{ from: 1, pattern: 'cover' }, { from: 2, pattern: 'free', change: 'Add the plans', passages: [1, 2, 3, 7, 8] }, { from: 3, pattern: 'free' }, { from: 4, pattern: 'closing' }]), titles, titles, titles);
    await planEdit(demo, given, 'Add the part about the plans to slide 2', { llm: model.llm, ...options });
    return 'b1Text' in (schemaOf(model.requests[1]).properties.s2?.properties ?? {});
  };
  assert.equal(await run(deck), false);
  // A deck whose document is in another language, by the planner's word: the writer writes the text.
  assert.equal(await run({ ...deck, documentLanguage: 'fr' }), true);
  // A deck stored without it goes by the script of the document.
  const { documentLanguage: _language, ...stored } = deck;
  assert.equal(await run(stored), false);
});

/** What a change that stops before writing says of its slides, or `undefined` where it did not stop. */
const stoppedAt = (edit: Promise<unknown>): Promise<TextCutError['fits'] | undefined> =>
  edit.then(
    () => undefined,
    (error: unknown) => {
      if (error instanceof TextCutError) return error.fits;
      throw error;
    },
  );

test('planEdit stops before writing for the room a slide is written in, a heading it keeps counted', async () => {
  const long = hebrewDocument(1500);
  const doc = readDocument(long, 'he');
  const every = doc.passages.map((passage) => passage.number);
  const half = Math.ceil(every.length / 2);
  const slide = (passages: number[], heading: boolean): DeckSpec['slides'][number] => ({ pattern: 'free', blocks: ['text'], arrange: { dense: true }, passages, slots: { title: 'דוח', ...(heading ? { b1Heading: 'רקע' } : {}), b1Text: ['טקסט'] } });
  const deck = (heading: boolean): DeckSpec => ({ version: 1, language: 'he', direction: 'rtl', title: 'דוח', slides: [slide(every.slice(0, half), heading), slide(every.slice(half), false)], documentLanguage: 'he', document: long });
  const plan = editOf([{ from: 1, pattern: 'free', change: 'All of it here', passages: every }, { from: 2, pattern: 'free', change: 'Another title', passages: [] }], { language: 'he', fields: {} });
  const writer = (request: JsonRequest) => ({ s1: { title: 'הכול', ...(schemaOf(request).properties.s1?.properties.b1Heading ? { b1Heading: 'רקע' } : {}), b1Text: ['טקסט'] }, s2: { title: 'עוד' } });
  // Without its heading the slide holds every passage, set by code.
  const held = await planEdit(rtl, deck(false), 'Put all of it on slide 1', { llm: scripted(plan, writer).llm, ...rtlOptions, stopBeforeCut: true });
  assert.deepEqual(held.deck.slides[0]?.slots.b1Text, fillParagraphs(doc.passages, every));
  // The heading it keeps takes some of that room: the change stops after its plan.
  const stop = scripted(plan);
  const fits = await stoppedAt(planEdit(rtl, deck(true), 'Put all of it on slide 1', { llm: stop.llm, ...rtlOptions, stopBeforeCut: true }));
  assert.equal(stop.requests.length, 1);
  assert.equal(fits?.length, 1);
  const [fit] = fits ?? [];
  assert.deepEqual({ ...fit, holds: 0 }, { slide: 1, words: doc.words, holds: 0, size: 7, dense: true, passages: `p1-p${every.length}` });
  // Written all the same, the slide holds as many words as the stop said.
  const written = await planEdit(rtl, deck(true), 'Put all of it on slide 1', { llm: scripted(plan, writer, writer, writer).llm, ...rtlOptions });
  const cut = /is (\d+) characters, and the slide holds about (\d+) even with its text set small/.exec(written.warnings.find((issue) => issue.message.startsWith('the text from the document'))?.message ?? '');
  assert.ok(cut, JSON.stringify(written.warnings));
  assert.equal(fit?.holds, Math.floor((doc.words * Number(cut[2])) / Number(cut[1])));
});

test('planEdit: a part added beside a text from the document leaves it less room, and the cut is said before writing', async () => {
  const deck = await reportDeck();
  // Slide 2 takes in the text of slide 3: p1 to p6, set by code on a text alone.
  const merged = await planEdit(demo, deck, 'Put p1 to p6 on slide 2', { llm: scripted(editOf([{ from: 1, pattern: 'cover' }, { from: 2, pattern: 'free', change: 'Take in slide 3', passages: [1, 2, 3, 4, 5, 6] }, { from: 4, pattern: 'closing' }]), titles).llm, ...options });
  const doc = readDocument(REPORT, 'en');
  assert.deepEqual(merged.deck.slides[1]?.slots.b1Text, fillParagraphs(doc.passages, [1, 2, 3, 4, 5, 6]));
  const change = editOf([
    { from: 1, pattern: 'cover' },
    {
      from: 2,
      pattern: 'free',
      change: 'Add three key points beside the text',
      parts: [
        { from: 1, kind: 'text', rewrite: false },
        { from: 0, kind: 'list', rewrite: true },
      ],
    },
    { from: 3, pattern: 'closing' },
  ]);
  const stop = scripted(change);
  const fits = await stoppedAt(planEdit(demo, merged.deck, 'Add three key points beside the text of slide 2', { llm: stop.llm, ...options, stopBeforeCut: true }));
  assert.equal(stop.requests.length, 1);
  const [fit] = fits ?? [];
  assert.deepEqual({ ...fit, holds: 0 }, { slide: 2, words: passageWords(doc.passages, [1, 2, 3, 4, 5, 6]), holds: 0, passages: 'p1-p6' });
  // Without the stop the writer writes the text for the room it has beside the list, and the deck
  // is told. Kept as it was, it ran three times over its box.
  const model = scripted(change, titles, titles, titles);
  const edited = await planEdit(demo, merged.deck, 'Add three key points beside the text of slide 2', { llm: model.llm, ...options });
  assert.match(String(schemaOf(model.requests[1]).properties.s2?.properties.b1Text?.description), /^The slide's text from the document \(p1-p6\) shortened to \d+ to \d+ characters, the most this slide holds/);
  assert.deepEqual([edited.deck.slides[1]?.blocks, edited.deck.slides[1]?.arrange, edited.deck.slides[1]?.passages], [['text', 'list'], undefined, [1, 2, 3, 4, 5, 6]]);
  assert.ok(edited.warnings.some((issue) => issue.slide === 2 && issue.message.startsWith('the text from the document for this slide (p1-p6)')), JSON.stringify(edited.warnings));

  // A short text that still holds beside the list stays as it is: as the document has it, or as it
  // was written before (in simpler words, say). Nothing of it is asked of the writer.
  const text = ['Costs', paragraph(7, 150), '', 'Plans', paragraph(11, 150)].join('\n');
  const outline = scripted(outlineOf([entry('free', 'Costs', [1], freeText), entry('free', 'Plans', [2], freeText)], { document_language: 'en' }));
  const plan = await planOutline(demo, { ...brief, document: text, slideCount: 2 }, { llm: outline.llm, ...options });
  const short = (await writeDeck(demo, plan, { llm: scripted(titles).llm, ...options })).deck;
  const simpler: DeckSpec = { ...short, slides: short.slides.map((slide, at) => (at === 0 ? { ...slide, slots: { ...slide.slots, b1Text: ['Costs went up a little.'] } } : slide)) };
  const beside = editOf([{ from: 1, pattern: 'free', change: 'Add three key points beside the text', parts: [{ from: 1, kind: 'text', rewrite: false }, { from: 0, kind: 'list', rewrite: true }] }, { from: 2, pattern: 'free' }]);
  for (const given of [short, simpler]) {
    const kept = scripted(beside, titles);
    const added = await planEdit(demo, given, 'Add three key points beside the text of slide 1', { llm: kept.llm, ...options, stopBeforeCut: true });
    assert.deepEqual([added.deck.slides[0]?.blocks, added.deck.slides[0]?.slots.b1Text], [['text', 'list'], given.slides[0]?.slots.b1Text]);
    assert.equal('b1Text' in (schemaOf(kept.requests[1]).properties.s1?.properties ?? {}), false);
    assert.deepEqual(added.warnings, []);
  }
});

test('planEdit: a text from the document asked to be shorter keeps its length when a part is added beside it', async () => {
  // Six passages of 380 characters on one slide, asked for 100 words, and written so.
  const text = ['Background', ...Array.from({ length: 6 }, (_, at) => paragraph(at, 380))].join('\n');
  const outline = scripted(outlineOf([entry('free', 'All of it', [1, 2, 3, 4, 5, 6], { ...freeText, words: 100 })], { document_language: 'en' }));
  const plan = await planOutline(demo, { ...brief, document: text, slideCount: 1 }, { llm: outline.llm, ...options });
  assert.equal(plan.slides[0]?.words, 100);
  const shorter = ['The city opened its bike share in 2019 with forty stations, and riders took more trips every year.', 'Students and commuters make up most members. Repairs take two days, and electric bikes came in 2022.'];
  const { deck } = await writeDeck(demo, plan, { llm: scripted({ s1: { title: 'Bikes', b1Text: shorter } }).llm, ...options });
  assert.deepEqual(deck.slides[0]?.slots.b1Text, shorter);
  const assets = [{ id: 'bikes', kind: 'photo' as const, description: 'Bikes at a station', data: new Uint8Array([1]) }];
  const writer = (request: JsonRequest) => {
    const asked = Object.keys(schemaOf(request).properties.s1?.properties ?? {});
    return { s1: Object.fromEntries(asked.map((name) => [name, name === 'title' ? 'Bikes' : name.endsWith('Items') ? ['One', 'Two', 'Three'] : name.endsWith('Photo') ? { asset: 'bikes', alt: 'Bikes' } : ['Written again.']])) };
  };
  // Beside a list or a photo it holds as it is: the slide is not planned for the whole of its
  // passages, the change does not stop on a cut, and the photo is not left out.
  for (const kind of ['list', 'photo'] as const) {
    const change = editOf([{ from: 1, pattern: 'free', change: `Add ${kind === 'list' ? 'three key points' : 'a photo of bikes'}`, images: kind === 'photo', parts: [{ from: 1, kind: 'text', rewrite: false }, { from: 0, kind, rewrite: true }] }]);
    const model = scripted(change, writer, writer);
    const edited = await planEdit(demo, deck, `Add ${kind === 'list' ? 'three key points' : 'a photo of bikes'} to slide 1`, { llm: model.llm, ...options, assets, stopBeforeCut: true });
    const slide = edited.deck.slides[0];
    assert.deepEqual([slide?.blocks, slide?.slots.b1Text, slide?.passages], [['text', kind], shorter, [1, 2, 3, 4, 5, 6]], kind);
    assert.equal('b1Text' in (schemaOf(model.requests[1]).properties.s1?.properties ?? {}), false, kind);
    assert.deepEqual(edited.warnings, [], kind);
  }
});

test('planEdit: a deck from a document put into another language is not stopped as a cut on a slide that held its text', async () => {
  // Eight paragraphs of English on one slide, set dense. In Hebrew their words are counted at 6.5
  // characters each: more than the slide holds, though it held the English text.
  const text = ['Background', ...Array.from({ length: 8 }, (_, at) => paragraph(at, 520))].join('\n');
  const outline = scripted(outlineOf([entry('free', 'All of it', [1, 2, 3, 4, 5, 6, 7, 8], freeText)], { document_language: 'en' }));
  const plan = await planOutline(demo, { ...brief, document: text, slideCount: 1 }, { llm: outline.llm, ...options });
  const { deck, warnings } = await writeDeck(demo, plan, { llm: scripted(titles).llm, ...options });
  assert.deepEqual([deck.slides[0]?.arrange?.dense, warnings], [true, []]);
  const doc = readDocument(text, 'en');
  assert.ok(charsOfWords(doc.words, 'he') > (deck.slides[0]?.slots.b1Text as string[]).join('').length);
  const change = editOf([{ from: 1, pattern: 'free', change: 'Translate into Hebrew' }], { language: 'he', title: 'אופניים' });
  const hebrew = { s1: { title: 'אופניים', b1Text: Array.from({ length: 8 }, (_, at) => paragraph(at, 480, HEBREW)) } };
  const model = scripted(change, hebrew, hebrew);
  const edited = await planEdit(demo, deck, 'Translate the deck into Hebrew', { llm: model.llm, ...options, stopBeforeCut: true });
  // The writer is told to write the text to what the slide holds, not to shorten it, and no cut is said.
  assert.match(String(schemaOf(model.requests[1]).properties.s1?.properties.b1Text?.description), /^The slide's text from the document \(p1-p\d+\), in its own words: \d+ to \d+ characters in all/);
  assert.deepEqual(edited.deck.slides[0]?.slots.b1Text, hebrew.s1.b1Text);
  assert.deepEqual(edited.warnings, []);
});

test('planEdit: a slide whose text the request asks for in other words is written, also where its passages change', async () => {
  const deck = await reportDeck();
  const doc = readDocument(REPORT, 'en');
  const plan = (reword: boolean) =>
    editOf([{ from: 1, pattern: 'cover' }, { from: 2, pattern: 'free', change: 'Simpler words, and add the electric bikes', passages: [1, 2, 3, 4], ...(reword ? { reword: true } : {}) }, { from: 3, pattern: 'free' }, { from: 4, pattern: 'closing' }]);
  const simpler = { s2: { title: 'Bikes', b1Text: ['The city has had bikes to share since 2019.', 'People ride more each year.'] }, s3: { title: 'Costs' } };
  const model = scripted(plan(true), simpler, simpler);
  const edited = await planEdit(demo, deck, 'Slide 2 in simpler words, and move the electric bikes onto it', { llm: model.llm, ...options });
  // The plan may say so for each slide.
  const entry = (schemaOf(model.requests[0]).properties.slides as unknown as { items: Schema }).items;
  assert.equal(entry.properties.reword?.type, 'boolean');
  assert.ok(!entry.required?.includes('reword'));
  assert.ok((model.requests[0]?.prompt ?? '').includes('\n- reword: true only when the request asks for the text of this slide in other words'));
  // Slide 2 is written from its passages, and kept as the writer wrote it: it is not held to their words.
  assert.match(String(schemaOf(model.requests[1]).properties.s2?.properties.b1Text?.description), /^The slide's text from the document \(p1-p4\)/);
  assert.deepEqual(edited.deck.slides[1]?.slots.b1Text, simpler.s2.b1Text);
  assert.equal(model.requests.length, 2);
  assert.deepEqual(edited.warnings, []);
  // Slide 3, which only lost p4, is set by code.
  assert.deepEqual(edited.deck.slides[2]?.slots.b1Text, fillParagraphs(doc.passages, [5, 6]));
  // Without it, the passages are set as the document has them.
  const moved = scripted(plan(false), titles);
  const filled = await planEdit(demo, deck, 'Move the electric bikes onto slide 2', { llm: moved.llm, ...options });
  assert.deepEqual(filled.deck.slides[1]?.slots.b1Text, fillParagraphs(doc.passages, [1, 2, 3, 4]));
});
