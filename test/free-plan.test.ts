// Free slides in the planner and the chat tools: when one is on offer, how a plan names its
// parts, how its content is written (with the other slides, or by calls of its own for a slide of
// many parts), what becomes of a part that comes back without content, and how a deck that has
// one is changed part by part. No test calls a live model.
import assert from 'node:assert/strict';
import { readFileSync } from 'node:fs';
import { dirname, join } from 'node:path';
import { test } from 'node:test';
import { fileURLToPath } from 'node:url';
import { placeholderPhoto } from '../examples/placeholder.ts';
import { PlanError, TOOL_NAMES, editPlan, freeSlide, generateDeck, loadTemplate, memoryStore, planDeck, planEdit, planOutline, presentationTools, renderDeck, resolveAssets, storeKeys, validateDeck, writeDeck } from '../src/index.ts';
import type { Brief, ChartValue, DeckPlan, DeckSpec, FreeBlockKind, Issue, JsonRequest, LlmProvider, OutlineEntry, SlideSpec, TemplateManifest } from '../src/index.ts';
import { freePattern } from '../src/free/offer.ts';
import { slotNote } from '../src/plan/catalog.ts';
import type { CatalogContext } from '../src/plan/catalog.ts';
import { freeWritePattern } from '../src/plan/free.ts';
import { partRanges, rangeSchema, readRange } from '../src/plan/parts.ts';
import type { FreeJob } from '../src/plan/parts.ts';
import { heldParts, roomAsHeld } from '../src/plan/reshape.ts';
import { buildFixturePptx, cardsManifest } from './fixture.ts';
import { withoutCodes } from './issues.ts';

const pptx = buildFixturePptx();
// Three patterns: `content` (with a photo), `closing` (its text is a deck field) and `cards`. A free slide stands on the frame of `content`.
const template = loadTemplate(pptx, cardsManifest);
const fields = { closing: 'Thank you' };
const brief: Brief = { prompt: 'A deck about dogs', language: 'en', fields, slideCount: 2 };

// The demo template has a chart of its own, so its free slides take charts in its colours. Its frame shows a text of the deck.
const root = join(dirname(fileURLToPath(import.meta.url)), '..');
const demoManifest = JSON.parse(readFileSync(join(root, 'templates', 'demo.manifest.json'), 'utf8')) as TemplateManifest;
const demoPptx = new Uint8Array(readFileSync(join(root, 'templates', 'demo.pptx')));
const demo = loadTemplate(demoPptx, demoManifest);
const stores: Brief = { prompt: 'Sales of every store by quarter in 2025: 10, 14, 12 and 18 thousand.', language: 'en', fields: { organizationDepartment: 'Retail' } };

/** A model that replays prepared answers and records what it was asked. An answer may be worked out from the request. */
function scripted(...answers: unknown[]): { llm: LlmProvider; requests: JsonRequest[] } {
  const requests: JsonRequest[] = [];
  return {
    requests,
    llm: {
      async generateJson(request) {
        requests.push(request);
        if (answers.length === 0) throw new Error('the scripted model has no answer left');
        const answer = answers.shift();
        return typeof answer === 'function' ? (answer as (request: JsonRequest) => unknown)(request) : answer;
      },
    },
  };
}

type Schema = { properties: Record<string, Schema>; required?: string[]; enum?: string[]; items?: Schema; minItems?: number; maxItems?: number; description?: string; type?: string };
const schemaOf = (request: JsonRequest | undefined) => request?.schema as unknown as Schema;
const entryOf = (request: JsonRequest | undefined) => schemaOf(request).properties.slides?.items as Schema;
const offered = (request: JsonRequest | undefined) => entryOf(request).properties.pattern?.enum;
const kindsOf = (request: JsonRequest | undefined) => entryOf(request).properties.shows?.items?.properties.kind?.enum;
const promptOf = (request: JsonRequest | undefined) => request?.prompt ?? '';
/** The rules of a planning prompt: its last block, one rule to a line. */
const rulesOf = (request: JsonRequest | undefined) => (promptOf(request).split('\n\n').at(-1) ?? '').split('\n');
/** The line of the list of patterns that offers free slides. */
const freeLineOf = (request: JsonRequest | undefined) => promptOf(request).split('\n').find((line) => line.startsWith('- free: '));

/** An entry of an outline as a model answers it where free slides are on offer: every property, with nothing in the three that are for free slides alone. */
const entry = (pattern: string, purpose: string, more: Record<string, unknown> = {}) => ({ pattern, purpose, shows: [], shows_main: 'auto', side_by_side: 0, ...more });
const freeEntry = (purpose: string, shows: Array<[kind: string, number: number]>, main = 'auto', side = 0) =>
  entry('free', purpose, { shows: shows.map(([kind, number]) => ({ kind, number })), shows_main: main, side_by_side: side });
const outlineOf = (...slides: unknown[]) => ({ title: 'Dogs', slides });

const KINDS = 'text (running text, with a heading when there are several), list (a list of short points), card (a heading and a short text in a box), step (a numbered step: a heading and a short text), quote (a quotation), note (one highlighted line across the slide, for the bottom line), photo (a photo)';
// The brief of the fixture has no number, so neither charts nor tables are on offer: of the three, the line names pictures alone.
const FREE_LINE =
  `- free: A slide the software lays out itself from the parts you name in "shows", any number of each: ${KINDS}. ` +
  'Only for content that no layout above shows: several pictures on one slide, more items than a layout takes, or parts that no layout combines. ' +
  "A slide that a layout above shows well uses that layout. About 8 parts read well at the template's text size; more are set smaller, so plan more only when the brief asks for them.";
const PARTS_RULE =
  '- A slide with the layout free names its parts in "shows", in reading order. A number of things the brief asks to see on one slide (twelve charts, forty logos) is that number of parts, however many.';
const KEEP_PARTS_RULE = '- A slide with the layout free names its parts in "shows". Repeat them exactly as the plan has them unless the request changes what that slide shows.';

// What every plan is told of an amount of text asked for a slide (0.6.1), and the property it answers in. A change to a
// deck says it as the rule for "words"; a change to a plan keeps the words its slides have.
const WORDS_RULE = (asked: 'brief' | 'request' | 'deck', keep = false) =>
  `- ${asked === 'deck' ? 'words: ' : ''}When the ${asked === 'deck' ? 'request' : asked} asks for an amount of text on a slide ("200 words", "a paragraph of about 80 words"), set "words" on that slide to that number of words, ` +
  `and choose a pattern whose slots hold that much: a word takes about 6 characters with its space. Leave "words" out of every other slide${keep ? ', and keep the words a slide has unless the request changes them' : ''}.`;
const WORDS = { type: 'integer', minimum: 1, maximum: 3000, description: 'Only when an amount of text is asked for on this slide: the number of words. Leave it out otherwise.' };

// The outline of the fixture as it was asked before there were free slides: the list of patterns, the rules and the
// schema, with the rule and the property on asked lengths that every plan has since 0.6.1.
const PROMPT_BEFORE =
  'Brief:\nA deck about dogs\n\nPlan a deck of exactly 2 slides. Choose one pattern for each slide from this list ("?" marks an optional slot):\n\n' +
  '- content: Title, body, photo and icon.\n  slots: title (at most 20 characters); body (at most 3 items); photo (photo); note (text)\n' +
  '- closing: Closing words. Plain text.\n  slots: \n' +
  '- cards: A row of cards.\n  slots: title (text); cardNText (at most 20 characters)\n  items: 1 to 5 cards side by side in one row, numbered N = 1, 2, ...\n\n' +
  "- What the brief asks about the form of the deck comes first, before variety, before the usual choices below and before what a pattern's description advises: how many slides, which kinds of slides, how many things on a slide, how many side by side. Do exactly that, even when it means one slide or one pattern for every slide.\n" +
  '- For a pattern with a row of items, set "items" to the number the brief asks for, or else the number the content really has. Do not pad a list to fill a row.\n' +
  '- When the brief asks for a number of columns, or for items side by side, use a pattern that takes that number of items and whose items are "side by side in one row" or "in lines of" that number or more: either sets them on one line. Items "in lines of" a smaller number are broken over several lines. When the list has no such pattern, use the one that comes closest.\n' +
  `${WORDS_RULE('brief')}\n` +
  '- Choose a pattern only when the content can fill every required slot within its limits.\n' +
  '- A wish about the kind of slides (text only, no tables, only cards) decides which patterns may be used, and comes before variety. "Text only" means the patterns marked "Plain text.": a title with paragraphs or a list, and no cards, figures, tables, charts, icons or pictures. Using one pattern for many slides is right when the wish leaves no other.\n' +
  '- Otherwise vary the patterns across the deck.\n' +
  '- Unless the brief says how many pictures it wants: use a pattern with a photo on about half of the slides.';
const ENTRY_BEFORE = {
  type: 'object',
  properties: {
    pattern: { type: 'string', enum: ['content', 'closing', 'cards'] },
    purpose: { type: 'string', description: 'The one point this slide makes: a single sentence of at most 25 words' },
    items: { type: 'integer', minimum: 1, maximum: 5, description: 'For a pattern with a row of items: how many items the content calls for. Leave it out for other patterns.' },
    words: WORDS,
  },
  required: ['pattern', 'purpose'],
};

// A change to a deck of the fixture (two cards, then the closing slide) as it was asked before there were free slides:
// the plan of the change with its schema, and the call that writes the one slide that changes.
const EDIT_PROMPT_BEFORE =
  'The deck was written from this brief:\nA deck about dogs\n\n' +
  'The deck as it is now ("Dogs"), in the language with BCP 47 tag "en", body text size "normal":\n1. [cards, 2 items] title: "Two"; card1Text: "One"; card2Text: "Two"\n2. [closing] \n\n' +
  'Deck-level texts, shown on several slides:\nclosing: "Thank you"\n\nRequested change:\nShorten the title of slide 1\n\n' +
  'Return the slides of the deck after the change, in their new order.\n' +
  "- What the request asks about the form of the deck comes first, before variety, before the usual choices below and before what a pattern's description advises: how many slides, which kinds of slides, how many things on a slide, how many side by side. Do exactly that, even when it means one slide or one pattern for every slide.\n" +
  '- from: the number of the existing slide a slide is based on, or 0 for a new slide. Leave a slide out to delete it.\n' +
  '- pattern: keep the pattern of an existing slide unless the request calls for a different layout.\n' +
  '  A wish about the kind of slides (text only, no tables, only cards) decides which patterns may be used, and comes before variety. "Text only" means the patterns marked "Plain text.": a title with paragraphs or a list, and no cards, figures, tables, charts, icons or pictures. Using one pattern for many slides is right when the wish leaves no other.\n' +
  '  When the request asks for a number of columns, or for items side by side, use a pattern that takes that number of items and whose items are "side by side in one row" or "in lines of" that number or more: either sets them on one line. Items "in lines of" a smaller number are broken over several lines. When the list has no such pattern, use the one that comes closest.\n' +
  '- change: what to write or change in the content of that slide. Leave it empty for every slide whose content the request leaves alone.\n' +
  '- Colours are set by the software, not by you: for a wish about colour leave change empty.\n' +
  '- No colour was given with this change: if the request asks for one, say in summary that the colour was not changed.\n' +
  '- A wish about how the deck looks (more modern, less empty, less text) is met with another pattern for a slide, another number of items, or the text size. It changes layouts and settings, not wording: leave change empty unless the slide gets another pattern.\n' +
  '- A request to shorten, cut or rewrite the text of a slide is a change of its content: say it in change.\n' +
  '- images: true only when the pictures or icons of that slide should change.\n' +
  `${WORDS_RULE('deck')}\n` +
  '- items: for a pattern with a row of items, the number the slide has after the change.\n\n' +
  'Patterns, for new slides and for a slide that changes its layout ("?" marks an optional slot):\n' +
  '- content: Title, body, photo and icon.\n  slots: title (at most 20 characters); body (at most 3 items); photo (photo); note (text)\n' +
  '- closing: Closing words. Plain text.\n  slots: \n' +
  '- cards: A row of cards.\n  slots: title (text); cardNText (at most 20 characters)\n  items: 1 to 5 cards side by side in one row, numbered N = 1, 2, ...';
const EDIT_SCHEMA_BEFORE = {
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
          from: { type: 'integer', minimum: 0, maximum: 2, description: 'Number of the existing slide this slide is based on; 0 for a new slide' },
          pattern: { type: 'string', enum: ['content', 'closing', 'cards'] },
          change: {
            type: 'string',
            description:
              'What to write or change in the content of this slide: its text, numbers, pictures or parts. Empty when the content stays as it is; a new colour or a new place in the deck is not a change of content.',
          },
          images: { type: 'boolean', description: 'true only when the pictures or icons of this slide should change' },
          words: WORDS,
          items: { type: 'integer', minimum: 1, maximum: 5, description: 'For a pattern with a row of items: how many items the slide has after the change. Leave it out for other patterns.' },
        },
        required: ['from', 'pattern', 'change', 'images'],
      },
    },
    text_size: { type: 'string', enum: ['normal', 'large', 'x-large'], description: 'Size of body text after the change. Repeat the current size unless the request is about text size.' },
    fields: { type: 'object', description: 'The deck-level texts after the change. Repeat a value exactly unless the request affects it.', properties: { closing: { type: 'string' } }, required: ['closing'] },
  },
  required: ['summary', 'language', 'title', 'slides', 'text_size', 'fields'],
};
const WRITE_PROMPT_BEFORE =
  'The deck was written from this brief:\nA deck about dogs\n\nRequested change:\nShorten the title of slide 1\n\nThe deck after the change:\n1. cards, to write\n2. closing, unchanged \n\n' +
  'Write only the slides below. When a slide has current content, start from it and change only what is asked: keep every other text word for word.\n\n' +
  's1 (cards). Change: A shorter title\nCurrent content: {"title":"Two","card1Text":"One","card2Text":"Two"}\n\n' +
  'Fill every required property. Leave an optional property out when the slide is better without it.';

// --- when free slides are on offer, and what the outline says of them ---------------------------

test('the outline offers free slides after the template\'s patterns, and reads the parts a slide names', async () => {
  const { llm, requests } = scripted(
    // A slide of a pattern that names parts all the same: they are passed over.
    outlineOf(entry('cards', 'Two things', { items: 2, shows: [{ kind: 'card', number: 9 }], side_by_side: 4 }), freeEntry('Why dogs', [['card', 3], ['note', 1]], 'top', 9)),
  );
  const plan = await planOutline(template, brief, { llm });

  // The list: the template's patterns as they were, then one line for free slides, with no slots to list.
  const prompt = promptOf(requests[0]);
  assert.ok(prompt.includes(`  items: 1 to 5 cards side by side in one row, numbered N = 1, 2, ...\n${FREE_LINE}\n\n- What the brief asks`), freeLineOf(requests[0]) ?? 'no line for free slides');
  // It is no plain text, and has no slots to list before its parts are chosen.
  assert.doesNotMatch(freeLineOf(requests[0]) ?? '', /Plain text\.|slots:/);
  assert.equal(rulesOf(requests[0])[1], PARTS_RULE, 'right after the rule about wishes');
  assert.equal(rulesOf(requests[0]).at(-1), '- Unless the brief says how many pictures it wants: use a pattern with a photo on about half of the slides (a free slide with a photo counts).');

  // The entry: what it was, and three properties that every slide answers, in alphabetical order after the pattern and the purpose.
  assert.deepEqual(entryOf(requests[0]), {
    type: 'object',
    properties: {
      pattern: { type: 'string', enum: ['content', 'closing', 'cards', 'free'] },
      purpose: ENTRY_BEFORE.properties.purpose,
      items: { ...ENTRY_BEFORE.properties.items, description: `${ENTRY_BEFORE.properties.items.description} Not for free: its parts go in "shows".` },
      shows: {
        type: 'array',
        description: 'Only for the layout free: its parts in reading order, like parts together. [] for every other layout.',
        items: {
          type: 'object',
          properties: {
            count: { type: 'integer', minimum: 1, description: 'How many parts of this kind stand here, one after another: 8 for eight charts. One entry for them all, not one entry for each.' },
            kind: { type: 'string', enum: ['text', 'list', 'card', 'step', 'quote', 'note', 'photo'] },
          },
          required: ['count', 'kind'],
        },
      },
      shows_main: {
        type: 'string',
        enum: ['auto', 'left', 'right', 'top', 'none'],
        description: 'Only for free: where the first part stands when it is the main one and larger. auto lets the software decide; none sets all parts alike.',
      },
      side_by_side: { type: 'integer', minimum: 0, description: 'Only for free, when the brief asks for a number of columns; 0 lets the software decide.' },
      words: WORDS,
    },
    required: ['pattern', 'purpose', 'shows', 'shows_main', 'side_by_side'],
  });
  assert.deepEqual(Object.keys(entryOf(requests[0]).properties), ['pattern', 'purpose', 'items', 'shows', 'shows_main', 'side_by_side', 'words']);
  const declared = JSON.stringify(entryOf(requests[0]));
  for (const word of ['anyOf', 'oneOf', '"pattern":{"type":"string","pattern"', '"format"', '"enum":[]', 'minItems', 'maxItems']) assert.ok(!declared.includes(word), word);
  // Room for a plan of two slides, and for the parts each may name.
  assert.equal(requests[0]?.maxOutputTokens, 400 + 2 * 160 + 2 * 12);

  // The plan: a free slide lists its parts, and how they are set out where the answer says so. No other slide gains a key.
  assert.deepEqual(plan.slides, [
    { pattern: 'cards', purpose: 'Two things', items: 2 },
    // Nine side by side is more than the slide has parts: four it is.
    { pattern: 'free', purpose: 'Why dogs', blocks: ['card', 'card', 'card', 'note'], arrange: { lead: 'top', columns: 4 } },
  ]);
  assert.deepEqual(Object.keys(plan.slides[0] as OutlineEntry), ['pattern', 'purpose', 'items']);
  assert.deepEqual(JSON.parse(JSON.stringify(plan)), plan, 'a plan is JSON an app stores');

  // What a slide may be made of follows what the deck has: numbers in the brief, pictures that can be had, graphics that were supplied.
  const numbers = scripted(outlineOf(freeEntry('Sales', [['chart', 2]], 'auto', 0)));
  const graphic = { id: 'map', kind: 'graphic' as const, description: 'A map', data: new Uint8Array([1]) };
  const charted = await planOutline(template, { ...brief, prompt: 'Sales grew from 10 to 14', imageDensity: 'none', slideCount: 1 }, { llm: numbers.llm, assets: [graphic] });
  assert.deepEqual(kindsOf(numbers.requests[0]), ['text', 'list', 'card', 'figure', 'step', 'quote', 'note', 'chart', 'table', 'graphic']);
  assert.match(freeLineOf(numbers.requests[0]) ?? '', /figure \(one big number and what it counts\), .*chart \(a chart of numbers the brief states\), table \(a table\), graphic \(one of the supplied graphics, shown whole\)\. Only for content/);
  assert.deepEqual(charted.slides, [{ pattern: 'free', purpose: 'Sales', blocks: ['chart', 'chart'] }], 'auto and 0 leave the arrangement to the software');
  assert.equal(rulesOf(numbers.requests[0]).at(-1), '- Do not plan any photos.', 'the sentence it always was');

  // The line says what a free slide is for with the kinds that are on offer, and with no other: "several charts" is not
  // said to a deck whose brief has no number, nor "pictures" where none can be had.
  const forWhat = async (more: Partial<Brief>, options: { assets?: Array<typeof graphic> } = {}) => {
    const { llm, requests } = scripted(outlineOf(freeEntry('Sales', [['text', 1]])));
    await planOutline(template, { ...brief, slideCount: 1, ...more }, { llm, ...options });
    return /Only for content that no layout above shows: (.*?)\. A slide that a layout above shows well uses that layout\./.exec(freeLineOf(requests[0]) ?? '')?.[1];
  };
  const rest = 'more items than a layout takes, or parts that no layout combines';
  assert.equal(await forWhat({ prompt: 'Sales grew from 10 to 14' }), `several charts, tables or pictures on one slide, ${rest}`);
  assert.equal(await forWhat({ prompt: 'Sales grew from 10 to 14', imageDensity: 'none' }), `several charts or tables on one slide, ${rest}`);
  assert.equal(await forWhat({ prompt: 'Sales grew from 10 to 14', imageDensity: 'none' }, { assets: [graphic] }), `several charts, tables or pictures on one slide, ${rest}`, 'a supplied graphic is a picture');
  assert.equal(await forWhat({}), `several pictures on one slide, ${rest}`);
  assert.equal(await forWhat({ imageDensity: 'none' }), rest, 'none of the three is on offer');

  // A model that gives one entry for each part numbers them, chart 1 to chart 8 (seen with the production model, in
  // Hebrew). That is eight charts, not thirty-six. Counts that do not run up from 1 are counts: 2 and 2 are four.
  const each = (kind: string, counts: number[]) => counts.map((count) => ({ kind, count }));
  const shown = async (shows: unknown[]) => {
    const { llm } = scripted(outlineOf({ pattern: 'free', purpose: 'Sales', shows, shows_main: 'none', side_by_side: 4 }));
    return (await planOutline(template, { ...brief, prompt: 'Sales grew from 10 to 14', imageDensity: 'none', slideCount: 1 }, { llm })).slides[0]?.blocks?.join(' ');
  };
  assert.equal(await shown([...each('chart', [1, 2, 3, 4, 5, 6, 7, 8]), ...each('note', [1])]), 'chart chart chart chart chart chart chart chart note');
  assert.equal(await shown([...each('chart', [2, 2]), ...each('note', [1])]), 'chart chart chart chart note');
  assert.equal(await shown([...each('card', [1]), ...each('chart', [1, 2]), ...each('card', [1, 2, 3])]), 'card chart chart card card card');
  // The name the count had before is still read.
  assert.equal(await shown([{ kind: 'chart', number: 3 }]), 'chart chart chart');
});

test('a free slide whose parts cannot be used is asked for once more, and then put right by code', async () => {
  const cards = entry('cards', 'Two things', { items: 2 });
  const again = (problem: string) => `\n\nYour previous answer could not be used: ${problem}. Answer again using only the listed patterns.`;

  // No parts at all: asked once, then one text. Planning does not fail over the parts of a free slide.
  const none = scripted(outlineOf(cards, freeEntry('Why dogs', [])), outlineOf(cards, freeEntry('Why dogs', [])));
  const plain = await planOutline(template, brief, { llm: none.llm });
  assert.equal(none.requests.length, 2);
  assert.ok(promptOf(none.requests[1]).endsWith(again('slide 2: a free slide needs "shows", such as 6 charts and 1 note')));
  assert.deepEqual(plain.slides[1], { pattern: 'free', purpose: 'Why dogs', blocks: ['text'] });

  // A kind that is not on offer (a chart, in a brief without a number) is no part; what is left of the second answer is taken.
  const charts = scripted(outlineOf(cards, freeEntry('Why dogs', [['chart', 3]])), outlineOf(cards, freeEntry('Why dogs', [['chart', 3], ['note', 1], ['poem', 2]])));
  assert.deepEqual((await planOutline(template, brief, { llm: charts.llm })).slides[1]?.blocks, ['note']);
  assert.equal(charts.requests.length, 2);
  // The retry that answers well is taken as it is.
  const mended = scripted(outlineOf(cards, freeEntry('Why dogs', [])), outlineOf(cards, freeEntry('Why dogs', [['step', 2]], 'none', 1)));
  assert.deepEqual((await planOutline(template, brief, { llm: mended.llm })).slides[1], { pattern: 'free', purpose: 'Why dogs', blocks: ['step', 'step'], arrange: { lead: 'none', columns: 1 } });

  // More parts than any slide is asked for in earnest: said once, then cut.
  const many = scripted(outlineOf(freeEntry('Logos', [['card', 401]])), outlineOf(freeEntry('Logos', [['card', 401]])));
  const cut = await planOutline(template, { ...brief, slideCount: 1 }, { llm: many.llm });
  assert.ok(promptOf(many.requests[1]).endsWith(again('slide 1: a free slide takes at most 400 parts, not 401')));
  assert.equal(cut.slides[0]?.blocks?.length, 400);
  // A number no slide could hold is said as it was given, and never becomes a list of that length.
  const absurd = scripted(outlineOf(freeEntry('Logos', [['card', 1e9], ['note', 5]])), outlineOf(freeEntry('Logos', [['card', 1e9], ['note', 5]])));
  assert.equal((await planOutline(template, { ...brief, slideCount: 1 }, { llm: absurd.llm })).slides[0]?.blocks?.length, 400);
  assert.ok(promptOf(absurd.requests[1]).endsWith(again('slide 1: a free slide takes at most 400 parts, not 1000000005')));
  // Four hundred is no problem.
  const most = scripted(outlineOf(freeEntry('Logos', [['card', 400]])));
  assert.equal((await planOutline(template, { ...brief, slideCount: 1 }, { llm: most.llm })).slides[0]?.blocks?.length, 400);
  assert.equal(most.requests.length, 1);

  // Every photo of a slide is a picture to generate: six at most, and one more for each photo that was supplied.
  const gallery = scripted(outlineOf(freeEntry('Dogs', [['photo', 7], ['note', 1]])), outlineOf(freeEntry('Dogs', [['photo', 7], ['note', 1]])));
  const six = await planOutline(template, { ...brief, slideCount: 1 }, { llm: gallery.llm });
  assert.ok(promptOf(gallery.requests[1]).endsWith(again('slide 1: at most 6 photos can be made for one slide; supply pictures for more')));
  assert.deepEqual(six.slides[0]?.blocks, ['photo', 'photo', 'photo', 'photo', 'photo', 'photo', 'note']);
  const supplied = scripted(outlineOf(freeEntry('Dogs', [['photo', 8]])));
  const photo = (id: string) => ({ id, kind: 'photo' as const, description: 'A dog', data: new Uint8Array([1]) });
  assert.equal((await planOutline(template, { ...brief, slideCount: 1 }, { llm: supplied.llm, assets: [photo('rex'), photo('fido')] })).slides[0]?.blocks?.length, 8);
  assert.equal(supplied.requests.length, 1);

  // A row that cannot take the number asked of it: with free slides on offer, the retry says where such a slide can go.
  const nine = entry('cards', 'Nine things', { items: 9 });
  const row = scripted(outlineOf(nine, cards), outlineOf(freeEntry('Nine things', [['card', 9]]), cards));
  const moved = await planOutline(template, brief, { llm: row.llm });
  assert.ok(promptOf(row.requests[1]).endsWith(`${again('slide 1: pattern "cards" takes 1 to 5 cards, not 9').slice(0, -1)}, or use free for a slide that needs more items than a pattern takes.`));
  assert.equal(moved.slides[0]?.blocks?.length, 9);
  const off = scripted(outlineOf(nine, cards), outlineOf(cards, cards));
  await planOutline(template, brief, { llm: off.llm, free: false });
  assert.ok(promptOf(off.requests[1]).endsWith(again('slide 1: pattern "cards" takes 1 to 5 cards, not 9')));
});

test('with free slides turned off, or under a list that does not name them, the outline is asked as it was before', async () => {
  const answer = outlineOf({ pattern: 'cards', purpose: 'Two things', items: 2 });
  const ask = async (use: ReturnType<typeof loadTemplate>, more: Partial<Brief>, options: { free?: boolean; patterns?: string[] }) => {
    const { llm, requests } = scripted(answer);
    await planOutline(use, { ...brief, ...more }, { llm, ...options });
    return requests[0] as JsonRequest;
  };

  // The option: the prompt and the schema of before, character for character, in the room of before.
  const off = await ask(template, {}, { free: false });
  assert.equal(off.prompt, PROMPT_BEFORE);
  assert.deepEqual(entryOf(off), ENTRY_BEFORE);
  assert.equal(JSON.stringify(off.schema), JSON.stringify({ type: 'object', properties: { title: { type: 'string', description: 'Title of the deck' }, slides: { type: 'array', minItems: 2, maxItems: 2, items: ENTRY_BEFORE } }, required: ['title', 'slides'] }));
  assert.equal(off.maxOutputTokens, 400 + 2 * 160);

  // The app's own list, which an app may have passed for years: it names no free slides, so it gets none by upgrading.
  assert.deepEqual(await ask(template, {}, { patterns: ['content', 'closing', 'cards'] }), off);
  // A template whose manifest turns them off.
  assert.deepEqual(await ask(loadTemplate(pptx, { ...cardsManifest, free: false }), {}, {}), off);
  // The person's list: the patterns it names and nothing else.
  const listed = await ask(template, { patterns: ['cards', 'closing'] }, {});
  assert.deepEqual(offered(listed), ['closing', 'cards']);
  assert.doesNotMatch(JSON.stringify(listed), /free|shows|side_by_side/);

  // A change to such a plan says nothing of free slides either.
  const plan: DeckPlan = { version: 1, title: 'Dogs', brief, slides: [{ pattern: 'cards', purpose: 'Two things', items: 2 }] };
  const change = scripted({ summary: 'Done.', language: 'en', title: 'Dogs', slides: plan.slides });
  await editPlan(template, plan, 'A better title', { llm: change.llm, free: false });
  assert.doesNotMatch(JSON.stringify(change.requests[0]), /free|shows|side_by_side/);
  assert.deepEqual(entryOf(change.requests[0]), ENTRY_BEFORE);
  assert.equal(change.requests[0]?.maxOutputTokens, 400 + 11 * 160);
});

test('a list may name free slides, alone or with patterns, and a list that cannot have them says why', async () => {
  // Alone: the deck is made of free slides, and the line does not speak of layouts above it.
  const alone = scripted(outlineOf(freeEntry('Why dogs', [['card', 3]]), freeEntry('More', [['text', 1]])));
  const plan = await planOutline(template, { ...brief, patterns: ['free'] }, { llm: alone.llm });
  assert.deepEqual(offered(alone.requests[0]), ['free']);
  assert.equal(
    freeLineOf(alone.requests[0]),
    `- free: A slide the software lays out itself from the parts you name in "shows", any number of each: ${KINDS}. About 8 parts read well at the template's text size; more are set smaller, so plan more only when the brief asks for them.`,
  );
  assert.doesNotMatch(promptOf(alone.requests[0]), /layout above|- content: |- cards: /);
  assert.ok(!('items' in entryOf(alone.requests[0]).properties), 'no row of items is on offer');
  assert.ok(!rulesOf(alone.requests[0]).includes('- Otherwise vary the patterns across the deck.'), 'a list is a list: the rule about variety is left out');
  assert.deepEqual(plan.brief.patterns, ['free']);
  // The list goes to the deck, like any list.
  const written = await writeDeck(template, plan, { llm: scripted({ s1: { b1Heading: 'One', b1Text: 'A', b2Heading: 'Two', b2Text: 'B', b3Heading: 'Three', b3Text: 'C' }, s2: { b1Text: ['Because.'] } }).llm });
  assert.deepEqual(written.deck.patterns, ['free']);
  assert.deepEqual(written.deck.slides.map((slide) => slide.pattern), ['free', 'free']);

  // With patterns: those and free slides.
  const some = scripted(outlineOf(entry('cards', 'Two things', { items: 2 })));
  await planOutline(template, { ...brief, patterns: ['free', 'cards'] }, { llm: some.llm });
  assert.deepEqual(offered(some.requests[0]), ['cards', 'free']);
  // The app's list obeys the same rule: free slides where it names them.
  const app = scripted(outlineOf(entry('cards', 'Two things', { items: 2 })));
  await planOutline(template, brief, { llm: app.llm, patterns: ['cards', 'free'] });
  assert.deepEqual(offered(app.requests[0]), ['cards', 'free']);
  // Both lists: free slides when both name them.
  const both = scripted(outlineOf({ pattern: 'cards', purpose: 'Two things', items: 2 }));
  await planOutline(template, { ...brief, patterns: ['free', 'cards'] }, { llm: both.llm, patterns: ['cards', 'closing'] });
  assert.deepEqual(offered(both.requests[0]), ['cards']);

  // Named where none can be made: an error that says why, before any model is asked.
  const none = scripted();
  const refused = async (use: ReturnType<typeof loadTemplate>, given: Brief, options: { free?: boolean; patterns?: string[] } = {}) => {
    try {
      await planOutline(use, given, { llm: none.llm, ...options });
    } catch (error) {
      assert.ok(error instanceof PlanError, String(error));
      return error.message;
    }
    return 'no error';
  };
  const cannot = 'The deck cannot be limited to a pattern that cannot be filled here: ';
  assert.equal(await refused(template, { ...brief, patterns: ['free'] }, { free: false }), `${cannot}free (free slides are turned off)`);
  assert.equal(await refused(loadTemplate(pptx, { ...cardsManifest, free: false }), { ...brief, patterns: ['cards', 'free'] }), `${cannot}free (the template's manifest turns free slides off)`);
  const cardsOnly = loadTemplate(pptx, { ...cardsManifest, patterns: cardsManifest.patterns.filter((pattern) => pattern.id === 'cards') });
  assert.equal(await refused(cardsOnly, { ...brief, patterns: ['free'] }), `${cannot}free (the template has no layout with a title to stand on)`);
  // The frame of the demo template shows a text of the deck, which this brief does not give.
  assert.equal(await refused(demo, { prompt: 'Dogs', patterns: ['free', 'section'] }), `${cannot}free needs the deck field "organizationDepartment"`);
  assert.equal(await refused(template, { ...brief, patterns: ['free'] }, { patterns: ['cards'] }), 'The deck is limited to the patterns free, and the options allow only cards: no pattern is in both lists');
  assert.equal(none.requests.length, 0);

  // Without a list, a deck on that template is planned from what can be filled: no free slides, and no error.
  const open = scripted(outlineOf({ pattern: 'section', purpose: 'Part one' }));
  await planOutline(demo, { prompt: 'Dogs', slideCount: 1 }, { llm: open.llm, layout: 'adaptive' });
  assert.deepEqual(offered(open.requests[0]), ['section']);
});

test('a text-only deck is offered free slides of text alone', async () => {
  const { llm, requests } = scripted(outlineOf(freeEntry('Why dogs', [['text', 2], ['card', 3], ['note', 1]])));
  const plan = await planOutline(template, { ...brief, slideCount: 1, textOnly: true }, { llm });
  assert.deepEqual(offered(requests[0]), ['closing', 'free']);
  assert.deepEqual(kindsOf(requests[0]), ['text', 'list', 'quote', 'note']);
  assert.match(freeLineOf(requests[0]) ?? '', /any number of each: text \([^)]*\), list \([^)]*\), quote \([^)]*\), note \([^)]*\)\. Only for content that no layout above shows.* asks for them\. Plain text\.$/);
  assert.ok(rulesOf(requests[0]).includes('- The deck is text only, and every pattern in the list is plain text. Using the same pattern for several slides in a row is expected.'));
  assert.doesNotMatch(rulesOf(requests[0]).join('\n'), /photo/);
  // Cards are no plain text: what the answer names of them is passed over.
  assert.deepEqual(plan.slides[0]?.blocks, ['text', 'text', 'note']);
});

// --- changing a plan ------------------------------------------------------------------------------

const chartsPlan = (more: Partial<OutlineEntry> = {}): DeckPlan => ({
  version: 1,
  title: 'Stores',
  brief: { ...brief, prompt: 'Sales of our 6 stores' },
  slides: [
    { pattern: 'cards', purpose: 'Two things', items: 2 },
    { pattern: 'free', purpose: 'Six store charts', blocks: ['chart', 'chart', 'chart', 'chart', 'chart', 'chart', 'note'], arrange: { columns: 3, lead: 'left' }, ...more },
  ],
});
const revised = (...slides: unknown[]) => ({ summary: 'Done.', language: 'en', title: 'Stores', slides });

test('a change to a plan lists a free slide with its parts, and keeps them unless the answer names others', async () => {
  const plan = chartsPlan();
  const cards = entry('cards', 'Two things', { items: 2 });
  const closing = entry('closing', 'Thanks');

  // The answer says nothing of the parts: the slide with that purpose keeps what it had, wherever it stands now.
  const kept = scripted(revised(closing, freeEntry('Six store charts', []), cards));
  const first = await editPlan(template, plan, 'Open with the closing slide', { llm: kept.llm });
  assert.match(promptOf(kept.requests[0]), /\n1\. cards \(2 items\): Two things\n2\. free \(6 charts and a note; 3 side by side; main part at the left\): Six store charts\n\nRequested change:/);
  assert.equal(rulesOf(kept.requests[0])[2], KEEP_PARTS_RULE);
  assert.match(freeLineOf(kept.requests[0]) ?? '', /so plan more only when the request asks for them\.$/);
  assert.deepEqual(first.plan.slides, [{ pattern: 'closing', purpose: 'Thanks' }, plan.slides[1], plan.slides[0]]);
  assert.equal(kept.requests.length, 1, 'no second call for it');
  // One schema for a slide of a plan, new or changed.
  const outline = scripted(outlineOf(cards));
  await planOutline(template, { ...plan.brief, instructions: 'Open with the closing slide', slideCount: 1 }, { llm: outline.llm });
  assert.deepEqual(entryOf(kept.requests[0]), entryOf(outline.requests[0]));
  // Room for the plan, ten slides more, and what each may say of its parts; the two runs of parts the plan has are counted too.
  assert.equal(kept.requests[0]?.maxOutputTokens, 400 + 12 * 160 + (12 + 2) * 12);

  // The answer repeats the parts, as it is told to, and says nothing of how they are set out: that stays as planned too.
  const repeated = scripted(revised(cards, freeEntry('Six store charts', [['chart', 6], ['note', 1]])));
  assert.deepEqual((await editPlan(template, plan, 'A better title', { llm: repeated.llm })).plan.slides, plan.slides);
  // What it does say of the arrangement is taken, and the rest is kept.
  const wider = scripted(revised(cards, freeEntry('Six store charts', [['chart', 6], ['note', 1]], 'auto', 2)));
  assert.deepEqual((await editPlan(template, plan, 'Two charts side by side on slide 2', { llm: wider.llm })).plan.slides[1]?.arrange, { columns: 2, lead: 'left' });

  // A new purpose for the slide in the same place: its parts stay too.
  const reworded = scripted(revised(cards, freeEntry('Every store at a glance', [])));
  assert.deepEqual((await editPlan(template, plan, 'A better sentence for slide 2', { llm: reworded.llm })).plan.slides[1], { ...plan.slides[1], purpose: 'Every store at a glance' });

  // The request changes what the slide shows: the answer names the parts, and they are taken.
  const fewer = scripted(revised(cards, freeEntry('Six store charts', [['chart', 4], ['note', 1]], 'none', 2)));
  assert.deepEqual((await editPlan(template, plan, 'Only four charts on slide 2, two side by side, all alike', { llm: fewer.llm })).plan.slides[1], {
    pattern: 'free',
    purpose: 'Six store charts',
    blocks: ['chart', 'chart', 'chart', 'chart', 'note'],
    arrange: { lead: 'none', columns: 2 },
  });
  // The request adds a chart and says nothing of how the slide is set out. "auto" and 0 are what the schema asks for then,
  // and they do not undo what a person asked for earlier: three side by side, the main part at the left.
  const seventh = scripted(revised(cards, freeEntry('Six store charts', [['chart', 7], ['note', 1]])));
  assert.deepEqual((await editPlan(template, plan, 'A seventh chart on slide 2', { llm: seventh.llm })).plan.slides[1], {
    pattern: 'free',
    purpose: 'Six store charts',
    blocks: [...Array<string>(7).fill('chart'), 'note'],
    arrange: { lead: 'left', columns: 3 },
  });
  // No more side by side than the slide has parts after the change.
  const pair = scripted(revised(cards, freeEntry('Six store charts', [['chart', 2]])));
  assert.deepEqual((await editPlan(template, plan, 'Only two charts on slide 2', { llm: pair.llm })).plan.slides[1]?.arrange, { lead: 'left', columns: 2 });

  // A free slide the change adds has no parts to keep: it is asked for once more, and then is one text.
  const added = scripted(revised(cards, freeEntry('Six store charts', []), freeEntry('Our team', [])), revised(cards, freeEntry('Six store charts', []), freeEntry('Our team', [])));
  const grown = await editPlan(template, plan, 'Add a slide about the team', { llm: added.llm });
  assert.ok(promptOf(added.requests[1]).endsWith('\n\nYour previous answer could not be used: slide 3: a free slide needs "shows", such as 6 charts and 1 note. Answer again using only the listed patterns.'));
  assert.deepEqual(grown.plan.slides.slice(1), [plan.slides[1], { pattern: 'free', purpose: 'Our team', blocks: ['text'] }]);
});

test('a plan keeps its free slide where free slides are no longer on offer', async () => {
  const plan = chartsPlan();
  const cards = entry('cards', 'Two things', { items: 2 });
  // Turned off by the app since the plan was made. The slide is still listed, may still be named, and keeps its parts.
  const { llm, requests } = scripted(revised(cards, freeEntry('Six store charts', [])));
  const changed = await editPlan(template, plan, 'A better title', { llm, free: false });
  assert.deepEqual(offered(requests[0]), ['content', 'closing', 'cards', 'free']);
  assert.deepEqual(kindsOf(requests[0]), ['note', 'chart'], 'the kinds the plan uses, and no other');
  assert.equal(
    freeLineOf(requests[0]),
    '- free: A slide the software lays out itself from the parts you name in "shows", any number of each: note (one highlighted line across the slide, for the bottom line), chart (a chart of numbers the brief states). Only for the slides of the plan that have it already.',
  );
  assert.ok(rulesOf(requests[0]).includes(KEEP_PARTS_RULE));
  assert.doesNotMatch(rulesOf(requests[0]).at(-1) ?? '', /free slide/, 'photos are counted as before');
  assert.deepEqual(changed.plan.slides, plan.slides);
  // It gets no second one there: the option decides what a new slide may be. Said once, and then the change fails like any plan that cannot be used.
  const two = revised(cards, freeEntry('Six store charts', []), freeEntry('Our team', [['note', 1]]));
  const more = scripted(two, two);
  await assert.rejects(editPlan(template, plan, 'Add a slide about the team', { llm: more.llm, free: false }), /^PlanError: The model did not produce a usable plan: the layout free is only for the slide of the plan that has it already, and the answer gives it to 2$/);
  assert.ok(promptOf(more.requests[1]).endsWith('\n\nYour previous answer could not be used: the layout free is only for the slide of the plan that has it already, and the answer gives it to 2. Answer again using only the listed patterns.'));
  // The slide it has may still move.
  const movedFree = scripted(revised(freeEntry('Six store charts', []), cards));
  assert.deepEqual((await editPlan(template, plan, 'Charts first', { llm: movedFree.llm, free: false })).plan.slides, [plan.slides[1], plan.slides[0]]);

  // Under a list that does not name free slides it is the same.
  const listed = scripted(revised(cards, freeEntry('Six store charts', [])));
  const limited = await editPlan(template, plan, 'A better title', { llm: listed.llm, deckPatterns: ['cards'] });
  assert.match(freeLineOf(listed.requests[0]) ?? '', /Only for the slides of the plan that have it already\.$/);
  assert.ok(rulesOf(listed.requests[0]).includes('- This deck was limited to these layouts: cards. Keep to them unless the requested change asks for another.'));
  assert.deepEqual(limited.plan.slides, plan.slides);
  // A list may name free slides when it is set with a change, as when it comes with a brief.
  const named = scripted(revised(cards, freeEntry('Six store charts', [])));
  const both = await editPlan(template, plan, 'Only cards and free slides from now on', { llm: named.llm, deckPatterns: ['cards', 'free', 'tables'] });
  assert.deepEqual(both.plan.brief.patterns, ['cards', 'free']);
  assert.ok(rulesOf(named.requests[0]).includes('- This deck was limited to these layouts: cards, free. Keep to them unless the requested change asks for another.'));
  assert.match(freeLineOf(named.requests[0]) ?? '', /About 8 parts read well/);
});

// --- writing --------------------------------------------------------------------------------------

const cardsContent = { title: 'Two', card1Text: 'One', card2Text: 'Two' };
const threeCards = { title: 'Three ways in', b1Heading: 'Longer hours', b1Text: 'Bikes unlock from five.', b2Heading: 'Safer docks', b2Text: 'Every station gets a light.', b3Heading: 'Group passes', b3Text: 'One pass for five riders.' };
const cardsPlan = (more: Partial<OutlineEntry> = {}): DeckPlan => ({
  version: 1,
  title: 'Dogs',
  brief,
  slides: [
    { pattern: 'cards', purpose: 'Two things', items: 2 },
    { pattern: 'free', purpose: 'Three ways in', blocks: ['card', 'card', 'card'], ...more },
    { pattern: 'closing', purpose: 'Thanks' },
  ],
});

test('a free slide of a few parts is written with the other slides, a property for each slot', async () => {
  const { llm, requests } = scripted({ s1: cardsContent, s2: threeCards, s3: {} });
  const written = await writeDeck(template, cardsPlan({ arrange: { columns: 3 } }), { llm });
  assert.equal(requests.length, 1, 'one call for the deck, as before');

  // The writer is told which of the slide's slots are what: their names say only their number.
  assert.match(promptOf(requests[0]), /\ns1: cards\. Two things\ns2: free \(b1 to b3: 3 cards\)\. Three ways in\ns3: closing\. Thanks\n\nFill every required property\./);
  const s2 = schemaOf(requests[0]).properties.s2 as Schema;
  assert.deepEqual(Object.keys(s2.properties), ['title', 'b1Heading', 'b1Text', 'b2Heading', 'b2Text', 'b3Heading', 'b3Text']);
  // In strict layout too, which is the default, a part that may be left out is not required: the title here.
  assert.deepEqual(s2.required, ['b1Heading', 'b1Text', 'b2Heading', 'b2Text', 'b3Heading', 'b3Text']);
  assert.match(s2.properties.b2Text?.description ?? '', /^Write \d+ to \d+ characters\. at most \d+ characters$/);
  assert.match(s2.properties.b2Heading?.description ?? '', /^A heading of a few words\. at most \d+ characters$/);

  // The slide: what makes it a free slide, the frame it was planned on, and its content.
  assert.deepEqual(written.deck.slides[1], { pattern: 'free', blocks: ['card', 'card', 'card'], arrange: { columns: 3 }, base: 'content', slots: threeCards });
  assert.deepEqual(Object.keys(written.deck.slides[1] as SlideSpec), ['pattern', 'blocks', 'arrange', 'base', 'slots']);
  assert.deepEqual(written.deck.slides[0], { pattern: 'cards', items: 2, slots: cardsContent }, 'another slide is what it always was');
  assert.deepEqual(written.warnings, []);
  assert.deepEqual(written.outline, cardsPlan({ arrange: { columns: 3 } }).slides);
  assert.deepEqual(validateDeck(template, written.deck), []);
  assert.ok(renderDeck(template, written.deck).pptx.length > 1000);

  // Without an arrangement in the plan the slide states none.
  const plain = await writeDeck(template, cardsPlan(), { llm: scripted({ s1: cardsContent, s2: threeCards, s3: {} }).llm });
  assert.deepEqual(Object.keys(plain.deck.slides[1] as SlideSpec), ['pattern', 'blocks', 'base', 'slots']);

  // With an icon set a card is offered an icon, as a choice of ids; each kind of part says in its slots what it is.
  const icons = [{ id: 'clock', description: 'a clock', data: new Uint8Array([1]) }];
  const mixed: DeckPlan = { ...cardsPlan(), slides: [{ pattern: 'free', purpose: 'A mix', blocks: ['card', 'quote', 'list'] }] };
  const model = scripted({ s1: { b1Icon: 'clock', b1Heading: 'Hours', b1Text: 'From five.', b2Quote: 'I sold my car.', b3Items: ['One', 'Two'] } });
  const mix = await writeDeck(template, mixed, { llm: model.llm, icons });
  const s1 = schemaOf(model.requests[0]).properties.s1 as Schema;
  assert.deepEqual(Object.keys(s1.properties), ['title', 'b1Icon', 'b1Heading', 'b1Text', 'b2Quote', 'b2Source', 'b3Heading', 'b3Items']);
  assert.deepEqual(s1.required, ['b1Heading', 'b1Text', 'b2Quote', 'b3Items']);
  assert.deepEqual(s1.properties.b1Icon?.enum, ['clock']);
  assert.match(promptOf(model.requests[0]), /s1: free \(b1: a card; b2: a quote; b3: a list\)\. A mix/);
  assert.match(promptOf(model.requests[0]), /Icons to choose from:\n- clock: a clock/);
  assert.deepEqual(mix.deck.slides[0]?.slots, { b1Icon: { asset: 'clock' }, b1Heading: 'Hours', b1Text: 'From five.', b2Quote: 'I sold my car.', b3Items: ['One', 'Two'] });
});

test('text far over the budget the writer was shown is sent back, also where the slide still takes it', async () => {
  // The cards of this slide are shown a budget; the slide, laid out with the text, sets it smaller and has no complaint of its own.
  const first = await writeDeck(template, cardsPlan(), { llm: scripted({ s1: cardsContent, s2: threeCards, s3: {} }).llm });
  const model = scripted({ s1: cardsContent, s2: threeCards, s3: {} });
  await writeDeck(template, cardsPlan(), { llm: model.llm });
  const budget = Number(/at most (\d+) characters$/.exec((schemaOf(model.requests[0]).properties.s2 as Schema).properties.b1Text?.description ?? '')?.[1]);
  const long = 'Bikes unlock from five. '.repeat(40).slice(0, Math.ceil(budget * 1.3));
  const deckWith = (text: string): DeckSpec => ({ ...first.deck, slides: first.deck.slides.map((slide, index) => (index === 1 ? { ...slide, slots: { ...slide.slots, b1Text: text } } : slide)) });
  assert.deepEqual(validateDeck(template, deckWith(long)), [], 'the check of the deck lets it pass');

  const { llm, requests } = scripted({ s1: cardsContent, s2: { ...threeCards, b1Text: long }, s3: {} }, { s1: cardsContent, s2: threeCards, s3: {} });
  const written = await writeDeck(template, cardsPlan(), { llm });
  assert.equal(requests.length, 2);
  assert.ok(promptOf(requests[1]).includes(`\nIt has these problems:\n- s2.b1Text: text is ${long.length} characters; the slot fits about ${budget}\n\nReturn the full content again`), promptOf(requests[1]).slice(-300));
  assert.deepEqual(written.warnings, []);
  // A little over is left alone: the slide sets it a little smaller.
  const near = 'Bikes unlock from five. '.repeat(40).slice(0, Math.floor(budget * 1.2));
  const calm = scripted({ s1: cardsContent, s2: { ...threeCards, b1Text: near }, s3: {} });
  await writeDeck(template, cardsPlan(), { llm: calm.llm });
  assert.equal(calm.requests.length, 1);
  // What still stands after the last round is a warning on that slot, like any text over its budget.
  const stubborn = scripted({ s1: cardsContent, s2: { ...threeCards, b1Text: long }, s3: {} });
  const kept = await writeDeck(template, cardsPlan(), { llm: stubborn.llm, maxRepairs: 0 });
  assert.deepEqual(withoutCodes(kept.warnings), [{ severity: 'warning', slide: 2, slot: 'b1Text', message: `text is ${long.length} characters; the slot fits about ${budget}` }]);
});

const quarters = ['Q1', 'Q2', 'Q3', 'Q4'];
const chartOf = (name: string): ChartValue => ({ type: 'column', categories: quarters, series: [{ name, values: [10, 14, 12, 18] }] });
/** A chart as an item of an array, with several series or with one list of numbers. */
const chartItem = (block: number | undefined, more: Record<string, unknown> = {}) => ({
  ...(block === undefined ? {} : { block }),
  caption: `Store ${block ?? '?'}`,
  categories: quarters,
  format: 'number',
  type: 'column',
  series: [{ name: '2025', values: [10, 14, 12, 18] }],
  ...more,
});
const sparkItem = (block: number) => ({ block, caption: `Store ${block}`, categories: quarters, format: 'number', type: 'line', values: [10, 14, 12, 18] });
const range = (first: number, last: number) => Array.from({ length: last - first + 1 }, (_, index) => first + index);
const chartsDeck = (count: number, more: OutlineEntry[] = []): DeckPlan => ({
  version: 1,
  title: 'Stores',
  brief: stores,
  slides: [{ pattern: 'cover', purpose: 'Open' }, { pattern: 'free', purpose: 'Every store', blocks: [...Array<'chart'>(count).fill('chart'), 'note'] }, ...more],
});
const NOTE = 'Every store sold most in the last quarter of the year.';
const cover = { s1: { title: 'Stores' } };

test('a free slide of many parts is written by a call of its own, as arrays of like parts', async () => {
  // Parts 5, 8 and 10 come without a number, with a number that is taken and with one outside the array.
  const items = range(1, 12).map((block) => (block === 5 ? chartItem(undefined, { caption: 'Store 5' }) : block === 8 ? chartItem(3, { caption: 'Store 8' }) : block === 10 ? chartItem(99, { caption: 'Store 10' }) : chartItem(block)));
  const { llm, requests } = scripted({ r01: [...items, chartItem(13), chartItem(4)], r02: [{ block: 13, text: NOTE }], title: 'Every store' }, cover);
  const written = await writeDeck(demo, chartsDeck(12), { llm, layout: 'adaptive' });
  assert.equal(requests.length, 2, 'the slide of thirteen parts first, then the rest of the deck');

  // The call for the slide: what to write, in lines that name the parts by number.
  const [parts, rest] = requests as [JsonRequest, JsonRequest];
  assert.ok(
    promptOf(parts).endsWith(
      '\n\nThe slides of the deck:\ns1: cover. Open\ns2: free (12 charts and a note). Every store\n\n' +
        'Write parts 1 to 13 of slide 2 (free: 12 charts and a note): Every store\n- r01: parts 1 to 12, 12 charts\n- r02: part 13, a note\n' +
        'Part N is the item whose "block" is N. Where an item has nothing for a property, give "" or [].',
    ),
    promptOf(parts).slice(-420),
  );
  assert.ok(promptOf(parts).startsWith(`Brief:\n${stores.prompt}\n\n`));
  assert.ok((parts.maxOutputTokens ?? Infinity) <= 24000 && (parts.maxOutputTokens ?? 0) > 3000, String(parts.maxOutputTokens));
  assert.equal(parts.effort, undefined, 'a note is text to write');

  // Its schema: an array for each run of like parts and the title. No array states a length, at any depth.
  const schema = schemaOf(parts);
  assert.deepEqual(Object.keys(schema.properties), ['r01', 'r02', 'title']);
  assert.deepEqual(schema.required, ['r01', 'r02', 'title']);
  const charts = schema.properties.r01 as Schema;
  // What the array says of its charts is what their slot says, each thing once: what the part is for, then its limits, then the rule of every chart.
  assert.match(charts.description ?? '', /^Exactly 12 charts: parts 1 to 12\. A small chart\. When the brief gives no numbers for it, leave it empty\. at most \d+ categories, at most 2 series\. A chart of numbers the brief states\. Use only those numbers: never estimate, interpolate or invent a value to complete a series\.$/);
  assert.deepEqual(Object.keys(charts.items?.properties ?? {}), ['block', 'caption', 'categories', 'format', 'series', 'type']);
  assert.deepEqual(charts.items?.required, ['block', 'caption', 'categories', 'format', 'series', 'type']);
  assert.deepEqual(charts.items?.properties.block, { type: 'integer', description: 'The number of the part this item is, counting every part of the slide from 1.' });
  assert.deepEqual(Object.keys(charts.items?.properties.series?.items?.properties ?? {}), ['name', 'values']);
  assert.match(charts.items?.properties.caption?.description ?? '', /^What it shows, in a few words\. Give one to every chart of the slide, or to none\. at most \d+ characters\. "" for none$/);
  const note = schema.properties.r02 as Schema;
  assert.equal(note.description, 'Exactly 1 note: part 13.');
  assert.deepEqual(note.items?.required, ['block', 'text']);
  assert.match(note.items?.properties.text?.description ?? '', /^The one line to take away\. Write \d+ to \d+ characters\. at most \d+ characters$/);
  assert.match(schema.properties.title?.description ?? '', /^The title of the slide\. at most \d+ characters\. "" for a slide without one$/);
  const declared = JSON.stringify(schema);
  for (const word of ['anyOf', 'oneOf', '"format":"', 'minItems', 'maxItems', '"enum":[]']) assert.ok(!declared.includes(word), word);

  // The call for the rest: the slides it writes, and a word on the one that is written apart.
  assert.deepEqual(Object.keys(schemaOf(rest).properties), ['s1']);
  assert.ok(promptOf(rest).includes('\n\nWrite the content of each slide of this deck:\ns1: cover. Open\n\nThese slides of the deck are written apart, not here:\ns2: free (12 charts and a note). Every store\n\nFill every required property.'));

  // The slide: every item is the part its number says, or else the next part of its array that has none. Items beyond the twelve are dropped.
  const slide = written.deck.slides[1] as SlideSpec;
  assert.deepEqual(slide.blocks, [...Array<string>(12).fill('chart'), 'note']);
  assert.equal(slide.base, 'list-with-photo');
  assert.deepEqual(range(1, 12).map((block) => slide.slots[`b${block}Caption`]), range(1, 12).map((block) => `Store ${block}`));
  assert.deepEqual(slide.slots.b7Chart, chartOf('2025'));
  assert.deepEqual([slide.slots.title, slide.slots.b13Text], ['Every store', NOTE]);
  assert.equal(Object.keys(slide.slots).length, 26);
  assert.deepEqual(written.warnings.filter((issue) => issue.severity === 'error' || issue.slot !== undefined), []);
  assert.ok(renderDeck(demo, written.deck, { layout: 'adaptive' }).pptx.length > 1000);
});

test('what is wrong with a part written apart is said by its number, and a short array in one line', async () => {
  // Ten charts of twelve; then all twelve, with one whose numbers do not match its labels; then all well.
  const twelve = range(1, 12).map((block) => chartItem(block));
  const short = { r01: twelve.slice(0, 10), r02: [{ block: 13, text: NOTE }], title: 'Every store' };
  const uneven = { ...short, r01: twelve.map((item) => (item.block === 3 ? { ...item, series: [{ name: '2025', values: [10, 14, 12] }] } : item)), r02: [{ block: 13, text: `${NOTE} `.repeat(8) }] };
  const { llm, requests } = scripted(short, uneven, { ...short, r01: twelve }, cover);
  const events: string[] = [];
  const written = await writeDeck(demo, chartsDeck(12), { llm, layout: 'adaptive', onProgress: (event) => events.push(event.stage === 'repair' ? `repair ${event.attempt}: ${event.issues}` : event.stage) });
  assert.equal(requests.length, 4);

  // The first repair: the answer as it was given, and one line for the array that is short. No line for each chart that is missing.
  const one = promptOf(requests[1]);
  assert.ok(one.startsWith(promptOf(requests[0])));
  assert.ok(one.includes(`\n\nYour previous answer:\n${JSON.stringify(short)}\n\nIt has these problems:\n- s2.r01: 10 items; it needs exactly 12, parts 1 to 12\n\nReturn the full content again with these problems fixed. Change nothing else.`), one.slice(-400));
  assert.doesNotMatch(one, /this slot is required/);
  assert.deepEqual(requests[1]?.schema, requests[0]?.schema);
  // The second: a part is named by its number and its kind, where the slot's own name (b3Chart) is one the model never saw.
  const two = promptOf(requests[2]);
  assert.match(two, /\nIt has these problems:\n- s2 part 3 \(chart\): series "2025" has 3 values; the chart has 4 categories\n- s2 part 13 \(note\): text is \d+ characters; the slot fits about \d+\n\nReturn the full content again/);
  assert.deepEqual(events, ['write', 'repair 1: 1', 'repair 2: 2']);
  assert.equal((written.deck.slides[1] as SlideSpec).blocks?.length, 13);
  assert.deepEqual(written.warnings.filter((issue) => issue.slot !== undefined), []);
});

test('sixty charts are written by two calls side by side, and a call is asked again by itself', async () => {
  let inFlight = 0;
  let most = 0;
  const requests: JsonRequest[] = [];
  // The second call is three charts short, also when asked again.
  const answer = (request: JsonRequest): unknown => {
    if ('s1' in schemaOf(request).properties) return cover;
    const first = Number(/Write parts (\d+) to/.exec(request.prompt)?.[1]);
    return first === 1 ? { r01: range(1, 30).map(sparkItem), title: 'Every store' } : { r01: range(31, 57).map(sparkItem) };
  };
  const llm: LlmProvider = {
    async generateJson(request) {
      requests.push(request);
      most = Math.max(most, ++inFlight);
      await new Promise((resolve) => setTimeout(resolve, 5));
      inFlight--;
      return answer(request);
    },
  };
  const plan: DeckPlan = { ...chartsDeck(0), slides: [{ pattern: 'cover', purpose: 'Open' }, { pattern: 'free', purpose: 'Every store', blocks: Array<'chart'>(60).fill('chart') }] };
  const written = await writeDeck(demo, plan, { llm, layout: 'adaptive', maxRepairs: 1 });

  // Two calls for the slide, each with its own parts, sent together; one of them again; then the rest of the deck.
  assert.deepEqual(requests.map((request) => /Write (parts \d+ to \d+) of slide 2 \(free: 60 charts\): Every store/.exec(request.prompt)?.[1] ?? 'the rest'), ['parts 1 to 30', 'parts 31 to 60', 'parts 31 to 60', 'the rest']);
  assert.equal(most, 2);
  const [first, second, again] = requests as [JsonRequest, JsonRequest, JsonRequest];
  assert.deepEqual(Object.keys(schemaOf(first).properties), ['r01', 'title']);
  assert.deepEqual(Object.keys(schemaOf(second).properties), ['r01'], 'the title is written once, in the first call');
  // A chart this small shows one series: its numbers are one list. An item has no series to count, so no number of series is said.
  assert.equal(
    schemaOf(second).properties.r01?.description,
    'Exactly 30 charts: parts 31 to 60. A small chart. When the brief gives no numbers for it, leave it empty. at most 12 categories. ' +
      'A chart of numbers the brief states. Use only those numbers: never estimate, interpolate or invent a value to complete a series.',
  );
  assert.deepEqual(Object.keys(schemaOf(second).properties.r01?.items?.properties ?? {}), ['block', 'caption', 'categories', 'format', 'type', 'values']);
  // Charts alone are numbers to set down: little reasoning, and a ceiling that no call goes over.
  for (const request of [first, second, again]) {
    assert.equal(request.effort, 'low');
    assert.ok((request.maxOutputTokens ?? Infinity) <= 24000, String(request.maxOutputTokens));
  }
  assert.ok(promptOf(again).includes('\nIt has these problems:\n- s2.r01: 27 items; it needs exactly 30, parts 31 to 60\n\nReturn the full content again'));
  assert.ok(promptOf(again).startsWith(promptOf(second)));
  // The second call is told which of the slide's charts its parts are. The first has no chart before it, and is told nothing.
  assert.ok(promptOf(first).includes('\n- r01: parts 1 to 30, 30 charts\nPart N is the item'), promptOf(first).slice(-400));
  assert.ok(promptOf(second).includes('\n- r01: parts 31 to 60, 30 charts (the 31st to the 60th chart of the slide)\nPart N is the item'), promptOf(second).slice(-400));
  // A call sees its own parts alone. It is told that the others are written apart, and how the parts take what the brief
  // lists, so that two calls do not each chart the brief's first thirty things.
  for (const request of [first, second]) {
    assert.ok(
      promptOf(request).endsWith(
        '\nPart N is the item whose "block" is N. Where an item has nothing for a property, give "" or [].\n' +
          "The other parts of the slide are written apart in the same way. Where the brief lists what the parts show, like parts take it in the brief's order.",
      ),
      promptOf(request).slice(-300),
    );
  }

  // Still three short after the last round: the deck is made with the fifty-seven, and says so. A free slide stops no deck.
  const slide = written.deck.slides[1] as SlideSpec;
  assert.equal(slide.blocks?.length, 57);
  assert.deepEqual(slide.slots.b57Chart, { type: 'line', categories: quarters, series: [{ name: 'Store 57', values: [10, 14, 12, 18] }] });
  assert.ok(!('b58Chart' in slide.slots));
  const told: Issue = { severity: 'warning', message: 'slide 2: 3 parts had no content and were left out' };
  assert.ok(withoutCodes(written.warnings).some((issue) => JSON.stringify(issue) === JSON.stringify(told)), JSON.stringify(written.warnings));
  assert.deepEqual(written.warnings.filter((issue) => issue.severity === 'error'), []);
  assert.ok(renderDeck(demo, written.deck, { layout: 'adaptive' }).pptx.length > 1000);
});

test('a deck whose every slide is written apart asks for nothing else', async () => {
  const plan: DeckPlan = { ...chartsDeck(0), slides: [{ pattern: 'free', purpose: 'Every store', blocks: Array<'chart'>(10).fill('chart') }] };
  const { llm, requests } = scripted({ r01: range(1, 10).map((block) => chartItem(block)), title: '' });
  const written = await writeDeck(demo, plan, { llm, layout: 'adaptive' });
  assert.equal(requests.length, 1);
  assert.equal(written.deck.slides[0]?.blocks?.length, 10);
  assert.ok(!('title' in (written.deck.slides[0] as SlideSpec).slots), 'an empty title is no title');
});

const numbersPlan = (blocks: OutlineEntry['blocks'], purpose = 'Sales by store'): DeckPlan => ({ version: 1, title: 'Stores', brief: stores, slides: [{ pattern: 'free', purpose, blocks: blocks ?? [] }] });
const kinds = (...runs: Array<[kind: NonNullable<OutlineEntry['blocks']>[number], number: number]>) => runs.flatMap(([kind, number]) => Array<typeof kind>(number).fill(kind));
/** Every description a schema holds, at any depth. */
const descriptionsOf = (schema: unknown): string[] =>
  typeof schema !== 'object' || schema === null ? [] : Object.entries(schema).flatMap(([name, value]) => (name === 'description' && typeof value === 'string' ? [value] : descriptionsOf(value)));

test('what a writer is told of a slot states its limit once, and a card of an array may leave its icon out', async () => {
  const icons = [{ id: 'clock', description: 'a clock', data: new Uint8Array([1]) }];
  // The software that lays a slide out says how much a figure, a quotation or a list holds; the limit itself is joined on
  // after it. Said in both, it would be said twice, and for a list with two different numbers.
  const once = /^(?:[A-Z][^]*?[a-z)]\. )?(?:Write \d+ to \d+ characters(?: in all)?\. )?at most \d+ characters(?: in total, at most \d+ items)?(?:\. "" for none)?$/;
  const apart = scripted({});
  await writeDeck(demo, numbersPlan(kinds(['figure', 4], ['quote', 4], ['list', 2], ['card', 2], ['table', 2], ['note', 1])), { llm: apart.llm, layout: 'adaptive', icons, maxRepairs: 0 });
  const arrays = schemaOf(apart.requests[0]);
  assert.deepEqual(Object.keys(arrays.properties), ['r01', 'r02', 'r03', 'r04', 'r05', 'r06', 'title']);
  const item = (name: string) => arrays.properties[name]?.items as Schema;
  for (const [array, property, start] of [
    ['r01', 'label', 'What the figure counts. '],
    ['r01', 'value', 'The figure itself, such as 87% or 2.4M. When the brief gives no number for it, leave it empty. '],
    ['r02', 'quote', 'The quotation, word for word, without quotation marks. '],
    ['r03', 'items', 'One short point in each item. '],
    ['r04', 'heading', 'A heading of a few words. '],
    ['r04', 'text', ''],
    ['r06', 'text', 'The one line to take away. '],
  ] as const) {
    const said = item(array).properties[property]?.description ?? '';
    assert.ok(said.startsWith(start) && once.test(said), `${array}.${property}: ${said}`);
  }
  for (const said of descriptionsOf(apart.requests[0]?.schema)) assert.doesNotMatch(said, /At most \d+ character|\.\. |at most \d+ characters.*at most \d+ characters/, said);
  // An icon is a choice of ids, which has no value for none: a card may leave it out, as its description allows.
  assert.deepEqual(item('r04').properties.icon, { type: 'string', enum: ['clock'], description: 'An icon that suits the card. Give one to every card of the slide, or to none' });
  assert.deepEqual(item('r04').required, ['block', 'heading', 'text']);
  // A table may be left empty like a chart, and the writer is told so.
  assert.match(arrays.properties.r05?.description ?? '', /^Exactly 2 tables: parts 13 and 14\. A table; a word or a short phrase in each cell\. When the brief gives nothing for it, leave it empty\. at most \d+ columns, at most \d+ rows, at most \d+ characters in a cell\.$/);

  // The same slots as properties of a slide that is written with the deck.
  const flat = scripted({ s1: {} });
  await writeDeck(demo, numbersPlan(['figure', 'quote', 'list', 'card', 'table', 'note']), { llm: flat.llm, layout: 'adaptive', icons, maxRepairs: 0 });
  const s1 = schemaOf(flat.requests[0]).properties.s1 as Schema;
  for (const name of ['title', 'b1Label', 'b2Quote', 'b3Items', 'b4Heading', 'b4Text', 'b6Text']) assert.match(s1.properties[name]?.description ?? '', once, name);
  for (const said of descriptionsOf(s1)) assert.doesNotMatch(said, /At most \d+ character|\.\. |at most \d+ characters.*at most \d+ characters/, said);
  assert.match(s1.properties.b5Table?.description ?? '', /^A table; a word or a short phrase in each cell\. When the brief gives nothing for it, leave it empty\. at most \d+ columns, at most \d+ rows, at most \d+ characters in a cell$/);
  // The table is asked for like any part, and may be given empty: its schema states no least number of columns or rows.
  assert.ok(s1.required?.includes('b5Table') && !s1.required?.includes('b4Icon'));
  assert.doesNotMatch(JSON.stringify(s1.properties.b5Table), /minItems/);
});

test('an item of an array is read as a slot of a slide written with the deck is: a number as text, a part number given as a string', async () => {
  // Ten figures come back last first, their part numbers as strings and their values as numbers: an answer of a model whose
  // schema is not enforced. Each goes to the part it names, and 87 is the text "87", as it is for a slide of nine figures.
  const figures = range(1, 10)
    .reverse()
    .map((block) => ({ block: String(block), label: ['Riders', `in store ${block}`], value: 80 + block }));
  const apart = scripted({ r01: figures, title: 'Riders' });
  const written = await writeDeck(demo, numbersPlan(kinds(['figure', 10])), { llm: apart.llm, layout: 'adaptive' });
  const slots = written.deck.slides[0]?.slots ?? {};
  assert.equal(apart.requests.length, 1);
  assert.deepEqual(range(1, 10).map((block) => slots[`b${block}Value`]), range(1, 10).map((block) => String(80 + block)));
  assert.deepEqual([slots.b1Label, slots.b10Label], ['Riders in store 1', 'Riders in store 10']);
  assert.deepEqual(written.warnings.filter((issue) => /left out/.test(issue.message)), []);

  const flat = scripted({ s1: Object.fromEntries(range(1, 9).flatMap((block) => [[`b${block}Value`, 80 + block], [`b${block}Label`, ['Riders', `in store ${block}`]]])) });
  const together = await writeDeck(demo, numbersPlan(kinds(['figure', 9])), { llm: flat.llm, layout: 'adaptive' });
  assert.deepEqual([together.deck.slides[0]?.slots.b1Value, together.deck.slides[0]?.slots.b1Label], [slots.b1Value, slots.b1Label]);
});

test('a repair answer that holds only what was to be mended is laid over the answer before it', async () => {
  // Ten charts of twelve come first. Asked again, the model returns the two that were missing and nothing else:
  // the ten are not lost for having been right the first time, and neither is the title.
  const twelve = range(1, 12).map((block) => chartItem(block));
  const { llm, requests } = scripted({ r01: twelve.slice(0, 10), r02: [{ block: 13, text: NOTE }], title: 'Every store' }, { r01: twelve.slice(10) }, cover);
  const written = await writeDeck(demo, chartsDeck(12), { llm, layout: 'adaptive' });
  assert.equal(requests.length, 3, 'one repair round, and then the rest of the deck');
  const slide = written.deck.slides[1] as SlideSpec;
  assert.deepEqual(slide.blocks, [...Array<string>(12).fill('chart'), 'note']);
  assert.deepEqual(range(1, 12).map((block) => slide.slots[`b${block}Caption`]), range(1, 12).map((block) => `Store ${block}`));
  assert.deepEqual([slide.slots.title, slide.slots.b13Text], ['Every store', NOTE]);
  assert.deepEqual(written.warnings.filter((issue) => /left out/.test(issue.message)), []);

  // Items that come back without a number take the parts that have none yet, and not the places of those that were right.
  const blind = scripted({ r01: twelve.slice(0, 10), r02: [{ block: 13, text: NOTE }], title: 'Every store' }, { r01: [chartItem(undefined, { caption: 'Eleven' }), chartItem(undefined, { caption: 'Twelve' })] }, cover);
  const placed = (await writeDeck(demo, chartsDeck(12), { llm: blind.llm, layout: 'adaptive' })).deck.slides[1] as SlideSpec;
  assert.deepEqual([placed.slots.b1Caption, placed.slots.b10Caption, placed.slots.b11Caption, placed.slots.b12Caption], ['Store 1', 'Store 10', 'Eleven', 'Twelve']);
  // What the next round is shown as the answer before is the whole of it, each item with its number.
  const still = scripted({ r01: twelve.slice(0, 10), r02: [{ block: 13, text: NOTE }], title: 'Every store' }, { r01: [chartItem(11)] }, { r01: twelve }, cover);
  await writeDeck(demo, chartsDeck(12), { llm: still.llm, layout: 'adaptive' });
  assert.ok(promptOf(still.requests[2]).includes(`\n\nYour previous answer:\n${JSON.stringify({ r01: twelve.slice(0, 11), r02: [{ block: 13, text: NOTE }], title: 'Every store' })}\n\nIt has these problems:\n- s2.r01: 11 items; it needs exactly 12, parts 1 to 12\n\n`), promptOf(still.requests[2]).slice(-300));

  // A slide written with the deck: one card is far too long, and the repair answer gives that card's text alone for the slide.
  const first = scripted({ s1: cardsContent, s2: threeCards, s3: {} });
  await writeDeck(template, cardsPlan(), { llm: first.llm });
  const budget = Number(/at most (\d+) characters$/.exec((schemaOf(first.requests[0]).properties.s2 as Schema).properties.b1Text?.description ?? '')?.[1]);
  const long = 'Bikes unlock from five. '.repeat(40).slice(0, Math.ceil(budget * 1.3));
  const flat = scripted({ s1: cardsContent, s2: { ...threeCards, b1Text: long }, s3: {} }, { s1: cardsContent, s2: { b1Text: 'Bikes unlock from five.' }, s3: {} });
  const mended = await writeDeck(template, cardsPlan(), { llm: flat.llm });
  assert.equal(flat.requests.length, 2);
  assert.deepEqual(mended.deck.slides[1]?.slots, threeCards);
  assert.deepEqual(mended.warnings, []);
  // A slot the problems were about and that the new answer leaves out is left out: that is an answer to the problem.
  const sourced = { title: 'Said', b1Quote: 'I sold my car.', b1Source: 'A rider of many years who now cycles to work every single day of the week, whatever the weather is like', b2Text: NOTE };
  const dropped = scripted({ s1: sourced }, { s1: { title: 'Said', b1Quote: 'I sold my car.', b2Text: NOTE } });
  const quoted = await writeDeck(demo, numbersPlan(['quote', 'note']), { llm: dropped.llm, layout: 'adaptive' });
  assert.match(promptOf(dropped.requests[1]), /\nIt has these problems:\n- s1\.b1Source: text is \d+ characters; the slot fits about \d+\n/);
  assert.deepEqual(quoted.deck.slides[0]?.slots, { title: 'Said', b1Quote: 'I sold my car.', b2Text: NOTE });
});

test('text far over its budget is sent back whatever the budget, from twelve characters up', async () => {
  // A planned slide is set out for texts of the usual length, and each slot is told what its box holds then. So a budget
  // is one to write to on a slide of sixteen cards as on a slide of three, and text half as long again goes back there
  // too. Before, a budget under forty characters was a count of a few words that the slide did not hold a writer to.
  const budgets: number[] = [];
  for (const [use, count] of [[demo, 12], [demo, 16], [template, 12], [template, 24]] as const) {
    let budget = 0;
    const answer = (request: JsonRequest): unknown => {
      budget = Number(/at most (\d+) characters$/.exec((schemaOf(request).properties.r01?.items as Schema).properties.text?.description ?? '')?.[1]);
      const text = 'Open from five in the morning at every dock. '.repeat(20).slice(0, Math.ceil(budget * 1.5)).trim();
      return { r01: range(1, count).map((block) => ({ block, heading: 'Hours', text })), title: 'Why' };
    };
    const { llm, requests } = scripted(answer, answer, answer);
    const plan: DeckPlan = { version: 1, title: 'T', brief: use === demo ? stores : { ...brief, slideCount: 1 }, slides: [{ pattern: 'free', purpose: 'Why', blocks: kinds(['card', count]) }] };
    const written = await writeDeck(use, plan, { llm, layout: 'adaptive' });
    const sentBack = written.warnings.filter((issue) => issue.slot !== undefined && /characters/.test(issue.message));
    assert.deepEqual([requests.length, sentBack.length], [3, count], `${count} cards with a budget of ${budget} characters`);
    assert.match(promptOf(requests[1]), /\nIt has these problems:\n- s1 part 1 \(card\) text: text is \d+ characters; the slot fits about \d+\n/);
    budgets.push(budget);
  }
  // No card is told less than a dozen characters, and some of these budgets are under the forty that text was once held from.
  assert.ok(budgets.every((budget) => budget >= 12) && budgets.some((budget) => budget < 40), budgets.join(', '));
});

test('what a slot of a free slide says of itself reaches the writer as the engine says it, with its limits after it', async () => {
  const icons = [{ id: 'clock', description: 'a clock', data: new Uint8Array([1]) }];
  const context: CatalogContext = { layout: 'adaptive', fields: stores.fields as Record<string, string>, icons, assets: [], imageDensity: 'medium', figures: true, textScale: 1 };
  const RULE = 'A chart of numbers the brief states. Use only those numbers: never estimate, interpolate or invent a value to complete a series.';
  /** What the schema says of a slot: its own description, a full stop, its limits. Each once, and nothing else. */
  const saidOf = (slot: TemplateManifest['patterns'][number]['slots'][number]): string => {
    const limits =
      slot.kind === 'text'
        ? `at most ${slot.maxChars} characters${(slot.maxParagraphs ?? 1) > 1 ? ` in total, at most ${slot.maxParagraphs} items` : ''}`
        : slot.kind === 'chart'
          ? `at most ${slot.maxCategories} categories, at most ${slot.maxSeries} series`
          : slot.kind === 'table'
            ? `at most ${slot.maxColumns} columns, at most ${slot.maxRows} rows, at most ${slot.maxCellChars} characters in a cell`
            : '';
    return [slot.description, limits].filter(Boolean).join('. ');
  };

  // A slide written with the deck: a property for each slot.
  const blocks = kinds(['figure', 1], ['quote', 1], ['list', 1], ['card', 2], ['table', 1], ['chart', 1], ['note', 1]);
  const engine = freePattern(demo, { direction: 'ltr', textScale: 1 }, { blocks }, true);
  const made = freeWritePattern(demo, { direction: 'ltr', textScale: 1 }, { blocks }, context);
  assert.ok(typeof engine !== 'string' && 'pattern' in made);
  const stated = new Map(engine.pattern.slots.map((slot) => [slot.name, slot]));
  // The planner takes nothing from a description and puts nothing into one: no limit to strip, no full stop, no sentence to add.
  for (const { slot } of made.pattern.slots) assert.equal(slot.description, stated.get(slot.name)?.description, slot.name);
  assert.equal(stated.get('b6Table')?.description, 'A table; a word or a short phrase in each cell. When the brief gives nothing for it, leave it empty');
  for (const slot of engine.pattern.slots) assert.doesNotMatch(slot.description ?? '', /\.$|[Aa]t most \d|up to \d/, slot.name);
  const flat = scripted({ s1: {} });
  await writeDeck(demo, numbersPlan(blocks), { llm: flat.llm, layout: 'adaptive', icons, maxRepairs: 0 });
  const s1 = schemaOf(flat.requests[0]).properties.s1 as Schema;
  for (const name of ['b1Value', 'b1Label', 'b2Quote', 'b2Source', 'b3Heading', 'b3Items', 'b4Heading', 'b4Text', 'b6Table', 'b8Text']) {
    assert.equal(s1.properties[name]?.description, saidOf(stated.get(name) as never), name);
  }
  assert.equal(s1.properties.b7Chart?.description, `${saidOf(stated.get('b7Chart') as never)}. ${RULE}`);
  assert.equal(s1.properties.b4Icon?.description, 'An icon that suits the card. Give one to every card of the slide, or to none');
  // A chart and a table are required and may be left empty. Their lists say how, where the property has to be there.
  assert.equal(s1.properties.b7Chart?.properties.series?.description, 'One or more series. A pie takes exactly one. [] for a chart that is left empty.');
  assert.equal(s1.properties.b7Chart?.properties.categories?.description, 'Labels along the axis, or the names of the pie slices. [] for a chart that is left empty');
  assert.equal(s1.properties.b6Table?.properties.header?.description, 'Column headings, a word or two each. [] for a table that is left empty');
  assert.equal(s1.properties.b6Table?.properties.rows?.description, 'Body rows. Every row has one cell per column, in column order. [] for a table that is left empty.');
  // The chart of a pattern has to be filled, and is asked for as it always was.
  const pattern = scripted({ s1: {} });
  await writeDeck(demo, { version: 1, title: 'Stores', brief: stores, slides: [{ pattern: 'chart-native', purpose: 'Sales' }] }, { llm: pattern.llm, layout: 'adaptive', maxRepairs: 0 }).catch(() => undefined);
  const chart = (schemaOf(pattern.requests[0]).properties.s1 as Schema).properties.chart as Schema;
  assert.deepEqual([chart.properties.series?.description, chart.properties.series?.minItems, chart.properties.categories?.description], ['One or more series. A pie takes exactly one.', 1, 'Labels along the axis, or the names of the pie slices']);

  // A text whose budget is too small for a sentence on how much to write has no description at all: the schema says its limit, and no more.
  const sixteen = kinds(['card', 16]);
  const cards = freeWritePattern(demo, { direction: 'ltr', textScale: 1 }, { blocks: sixteen }, { ...context, icons: [] });
  assert.ok('pattern' in cards);
  const text = cards.pattern.slots.find((planned) => planned.slot.name === 'b1Text')?.slot;
  assert.ok(text?.kind === 'text' && text.description === undefined && (text.maxChars ?? 0) >= 12 && (text.maxChars ?? 0) < 40, JSON.stringify(text));
  const apart = scripted({});
  await writeDeck(demo, numbersPlan(sixteen), { llm: apart.llm, layout: 'adaptive', maxRepairs: 0 });
  assert.equal((schemaOf(apart.requests[0]).properties.r01?.items as Schema).properties.text?.description, `at most ${text.maxChars} characters`);

  // A limit of one is said in the singular, whatever slot states it.
  assert.equal(slotNote({ name: 'b1Table', kind: 'table', shapeId: '1', maxColumns: 1, maxRows: 1, maxCellChars: 1 }, context), 'at most 1 column, at most 1 row, at most 1 character in a cell');
  assert.equal(slotNote({ name: 'b1Chart', kind: 'chart', shapeId: '1', maxCategories: 1, maxSeries: 1, description: 'A small chart' }, context), 'A small chart. at most 1 category, at most 1 series');
  assert.equal(slotNote({ name: 'b1Text', kind: 'text', shapeId: '1', maxChars: 1 }, context), 'at most 1 character');
  assert.equal(slotNote({ name: 'b1Items', kind: 'text', shapeId: '1', maxChars: 40, maxParagraphs: 2 }, context), 'at most 40 characters in total, at most 2 items');
});

test('an array says of its parts what the smallest of them holds, and a call has room for the charts its slots say they show', () => {
  const context: CatalogContext = { layout: 'adaptive', fields: stores.fields as Record<string, string>, icons: [], assets: [], imageDensity: 'medium', figures: true, textScale: 1 };
  const jobOf = (blocks: FreeBlockKind[], arrange?: OutlineEntry['arrange']): FreeJob => {
    const made = freeWritePattern(demo, { direction: 'ltr', textScale: 1 }, { blocks, arrange }, context);
    assert.ok('pattern' in made);
    return { number: 1, pattern: made.pattern, purpose: 'Sales', free: { blocks, ...(arrange ? { arrange } : {}), base: made.base } };
  };
  const chartsOf = (job: FreeJob) => job.pattern.slots.flatMap(({ slot }) => (slot.kind === 'chart' ? [slot] : []));
  const said = (job: FreeJob) => (rangeSchema(job, partRanges(job, context)[0] as never, context) as unknown as Schema).properties.r01?.description ?? '';

  // Twelve charts, the first the main one: it is drawn large and the eleven beside it small. One array holds them all,
  // and is told the limits of the small ones. What fits the smallest part fits every one.
  const led = jobOf(kinds(['chart', 12]), { lead: 'left' });
  const [main, ...small] = chartsOf(led);
  assert.deepEqual([main?.detail ?? 'full', main?.maxCategories, main?.maxSeries], ['full', 8, 3]);
  const least = { categories: Math.min(...small.map((slot) => slot.maxCategories ?? 99)), series: Math.min(...small.map((slot) => slot.maxSeries ?? 99)) };
  assert.deepEqual(least, { categories: 12, series: 1 });
  // The eleven beside the main one are drawn bare (their text would be under 7 pt): the array is told the lowest of each limit, the main chart's categories and the small ones' series.
  assert.equal(said(led), `Exactly 12 charts: parts 1 to 12. A small chart. When the brief gives no numbers for it, leave it empty. at most ${main?.maxCategories} categories, at most ${least.series} series. A chart of numbers the brief states. Use only those numbers: never estimate, interpolate or invent a value to complete a series.`);
  // Tables likewise: of each limit the lowest that a table of the array states.
  const tables = jobOf(kinds(['table', 10]), { lead: 'top' });
  const limits = tables.pattern.slots.flatMap(({ slot }) => (slot.kind === 'table' ? [slot] : []));
  const lowest = (pick: (slot: (typeof limits)[number]) => number | undefined) => Math.min(...limits.map((slot) => pick(slot) ?? 99));
  assert.ok(new Set(limits.map((slot) => slot.maxColumns)).size > 1, 'the tables of this slide differ in size');
  assert.match(said(tables), new RegExp(`\\. at most ${lowest((slot) => slot.maxColumns)} columns, at most ${lowest((slot) => slot.maxRows)} rows, at most ${lowest((slot) => slot.maxCellChars)} characters in a cell\\.$`));

  // The ceiling of a call follows the limits too: a chart is as long as the categories and series its slot says it shows.
  const room = (job: FreeJob) => {
    const title = job.pattern.slots.find((planned) => planned.slot.name === 'title')?.slot;
    // For each part: a dozen tokens, its chart by the numbers it may show, its caption by its budget; half as much again for each.
    const parts = job.free.blocks.map((_, index) =>
      job.pattern.slots
        .filter((planned) => new RegExp(`^b${index + 1}[A-Z]`).test(planned.slot.name))
        .reduce((sum, { slot }) => sum + (slot.kind === 'chart' ? (40 + 12 * (slot.maxCategories ?? 0) * (1 + (slot.maxSeries ?? 0))) * 1.5 : slot.kind === 'text' ? ((slot.maxChars ?? 0) + 40) * 1.5 : 0), 12),
    );
    return Math.ceil(parts.reduce((sum, tokens) => sum + tokens, 0) + 300 + (title?.kind === 'text' ? ((title.maxChars ?? 0) + 40) * 1.5 : 0));
  };
  // Ten and twelve charts are small charts with names and numbers; twenty-four would have text under 7 pt and are drawn bare.
  for (const count of [10, 12]) {
    const job = jobOf(kinds(['chart', count]));
    assert.ok(chartsOf(job).every((slot) => slot.detail === 'compact' && slot.maxSeries === 2 && slot.maxCategories !== undefined), `${count} charts`);
    const ranges = partRanges(job, context);
    assert.deepEqual([ranges.length, ranges[0]?.room], [1, room(job)], `${count} charts`);
  }
  // Twelve charts of four categories and two series take less room than twenty-five drawn as bars alone, which may show twelve.
  assert.ok((partRanges(jobOf(kinds(['chart', 12])), context)[0]?.room ?? 0) < (partRanges(jobOf(kinds(['chart', 25])), context)[0]?.room ?? 0));
  // Whether a chart is asked for as one list of numbers is its slot's limit of one series, whatever the chart is drawn as.
  const bare = jobOf(kinds(['chart', 25]));
  assert.ok(chartsOf(bare).every((slot) => slot.maxSeries === 1));
  assert.match(said(bare), /^Exactly 25 charts: parts 1 to 25\. A small chart\. When the brief gives no numbers for it, leave it empty\. at most 12 categories\. A chart of numbers/);
  const one: FreeJob = { ...bare, pattern: { ...bare.pattern, slots: bare.pattern.slots.map((planned) => (planned.mode === 'chart' ? { ...planned, slot: { ...planned.slot, detail: 'compact' as const } } : planned)) } };
  const two: FreeJob = { ...bare, pattern: { ...bare.pattern, slots: bare.pattern.slots.map((planned) => (planned.mode === 'chart' ? { ...planned, slot: { ...planned.slot, maxSeries: 2 } } : planned)) } };
  const items = (job: FreeJob) => Object.keys(((rangeSchema(job, partRanges(job, context)[0] as never, context) as unknown as Schema).properties.r01?.items as Schema).properties);
  assert.ok(items(one).includes('values') && !items(one).includes('series'));
  assert.ok(items(two).includes('series') && !items(two).includes('values'));
});

test('parts of two kinds that take turns are written by calls of at most twelve arrays each', async () => {
  // A figure and its chart, twelve times: an array for every part. One schema of twenty-four arrays would be as large as
  // the schema of properties that arrays are there to avoid, so the slide is cut where a call would have a thirteenth.
  const blocks = range(1, 12).flatMap((): Array<'figure' | 'chart'> => ['figure', 'chart']);
  const answer = (request: JsonRequest): unknown =>
    Object.fromEntries(
      Object.entries(schemaOf(request).properties).map(([name, schema]) => {
        if (name === 'title') return [name, 'Every store'];
        const block = Number(/: part (\d+)\./.exec(schema.description ?? '')?.[1]);
        return [name, [block % 2 === 1 ? { block, label: 'thousand', value: '18' } : chartItem(block)]];
      }),
    );
  const { llm, requests } = scripted(answer, answer);
  const written = await writeDeck(demo, numbersPlan(blocks), { llm, layout: 'adaptive' });
  assert.deepEqual(requests.map((request) => Object.keys(schemaOf(request).properties)), [[...range(1, 12).map((run) => `r${String(run).padStart(2, '0')}`), 'title'], range(13, 24).map((run) => `r${run}`)]);
  assert.deepEqual(requests.map((request) => /Write (parts \d+ to \d+) of slide 1/.exec(request.prompt)?.[1]), ['parts 1 to 12', 'parts 13 to 24']);
  assert.equal(written.deck.slides[0]?.blocks?.length, 24);
  assert.deepEqual([written.deck.slides[0]?.slots.b23Value, written.deck.slides[0]?.slots.b24Caption], ['18', 'Store 24']);
  // The second call sees parts 13 to 24 alone: each of its lines says which figure or chart of the slide the part is.
  assert.doesNotMatch(promptOf(requests[0]), /of the slide\)/, 'no like part stands before the parts of the first call');
  assert.ok(promptOf(requests[1]).includes('\n- r13: part 13, a figure (the 7th figure of the slide)\n- r14: part 14, a chart (the 7th chart of the slide)\n'), promptOf(requests[1]).slice(-900));
});

test('a call that writes a later range of a slide is told which of their kind its parts are', async () => {
  // The parts a call is asked for, read from its schema: "Exactly 2 cards: parts 50 and 51."
  const numbersIn = (words: string) =>
    words.split(/, | and /).flatMap((span) => {
      const [from, to] = span.split(' to ').map(Number) as [number, number | undefined];
      return range(from, to ?? from);
    });
  const answer = (request: JsonRequest): unknown =>
    Object.fromEntries(
      Object.entries(schemaOf(request).properties).map(([name, schema]) => {
        if (name === 'title') return [name, 'Every store'];
        const [, kind, parts] = /^Exactly \d+ ([a-z]+?)s?: parts? ([^.]*)\./.exec(schema.description ?? '') ?? [];
        return [name, numbersIn(parts ?? '').map((block) => (kind === 'chart' ? sparkItem(block) : kind === 'figure' ? { block, label: 'thousand', value: '18' } : kind === 'note' ? { block, text: NOTE } : { block, heading: 'Hours', text: 'From five.' }))];
      }),
    );
  const written = async (blocks: NonNullable<OutlineEntry['blocks']>) => {
    const { llm, requests } = scripted(...Array<typeof answer>(40).fill(answer));
    const { deck } = await writeDeck(demo, numbersPlan(blocks), { llm, layout: 'adaptive', maxRepairs: 0 });
    assert.equal(deck.slides[0]?.blocks?.length, blocks.length);
    return requests.map(promptOf);
  };

  // A note, then sixty charts: part 32 is the 31st chart. A call that took "the charts" from the top of the brief would
  // write the first call's charts again.
  const charts = await written(['note', ...kinds(['chart', 60])]);
  assert.equal(charts.length, 2);
  assert.match(charts[0] ?? '', /\n- r01: part 1, a note\n- r02: parts 2 to \d+, \d+ charts\nPart N is the item/);
  const [, from, count, place] = /\n- r02: parts (\d+) to 61, (\d+) charts \(the (\d+)(?:st|nd|rd|th) to the 60th chart of the slide\)\nPart N is the item/.exec(charts[1] ?? '') ?? [];
  assert.deepEqual([Number(place), Number(count)], [Number(from) - 1, 62 - Number(from)], (charts[1] ?? '').slice(-400));

  // A card, a figure and its chart twenty-four times, and two cards: every part an array, so the slide takes several
  // calls. Each place is said as a person says it: 2nd, 3rd, 11th, 12th, 13th, 21st, 22nd, 23rd.
  const turns = await written(['card', ...range(1, 24).flatMap((): Array<'figure' | 'chart'> => ['figure', 'chart']), 'card', 'card']);
  assert.ok(turns.length >= 5, String(turns.length));
  assert.doesNotMatch(turns[0] ?? '', /of the slide\)/);
  assert.match(turns[0] ?? '', /\n- r01: part 1, a card\n- r02: part 2, a figure\n- r03: part 3, a chart\n- r04: part 4, a figure\n/);
  const later = turns.slice(1).join('\n');
  for (const line of [
    '- r22: part 22, a figure (the 11th figure of the slide)',
    '- r24: part 24, a figure (the 12th figure of the slide)',
    '- r26: part 26, a figure (the 13th figure of the slide)',
    '- r42: part 42, a figure (the 21st figure of the slide)',
    '- r43: part 43, a chart (the 21st chart of the slide)',
    '- r44: part 44, a figure (the 22nd figure of the slide)',
    '- r46: part 46, a figure (the 23rd figure of the slide)',
    '- r50: parts 50 and 51, 2 cards (the 2nd and the 3rd card of the slide)',
  ]) {
    assert.ok(later.includes(`\n${line}\n`), line);
  }
  // Every such call is still told how the parts take what the brief lists.
  for (const prompt of turns) assert.ok(prompt.endsWith("The other parts of the slide are written apart in the same way. Where the brief lists what the parts show, like parts take it in the brief's order."));

  // A slide that one call writes says nothing of it, wherever its parts stand.
  const one = await written(['note', ...kinds(['chart', 12])]);
  assert.equal(one.length, 1);
  assert.doesNotMatch(one[0] ?? '', /of the slide\)/);
});

test('of the photos a writer describes for one slide, six are made', async () => {
  const png = placeholderPhoto(40, 30, [14, 124, 123], [244, 162, 89]);
  const rex = { id: 'rex', kind: 'photo' as const, description: 'A dog called Rex', data: png };
  // Seven photos may be planned where one was supplied. The writer describes all seven, which would be seven pictures to generate.
  const plan: DeckPlan = { version: 1, title: 'Dogs', brief: { ...brief, slideCount: 1 }, slides: [{ pattern: 'free', purpose: 'Seven dogs', blocks: kinds(['photo', 7]) }] };
  const described = (supplied?: number) => ({ s1: { title: 'Seven dogs', ...Object.fromEntries(range(1, 7).map((block) => [`b${block}Photo`, { prompt: `Dog number ${block} in a park`, alt: 'A dog', ...(block === supplied ? { asset: 'rex' } : {}) }])) } });
  const { llm, requests } = scripted(described(), described(4));
  const written = await writeDeck(template, plan, { llm, assets: [rex] });
  // The writer is told the limit where a slide has more photos than are made for one, and told again when it goes over it.
  const s1 = schemaOf(requests[0]).properties.s1 as Schema;
  assert.equal(s1.properties.b7Photo?.description, 'Of the photos of this slide at most 6 can be made from a description: give the others a supplied photo');
  assert.ok(promptOf(requests[1]).includes('\nIt has these problems:\n- s1: at most 6 photos can be made for one slide; use a supplied photo for the others\n\nReturn the full content again'), promptOf(requests[1]).slice(-300));
  const slots = written.deck.slides[0]?.slots ?? {};
  assert.deepEqual([written.deck.slides[0]?.blocks?.length, slots.b4Photo, written.warnings], [7, { asset: 'rex', alt: 'A dog' }, []]);
  assert.equal(Object.values(slots).filter((value) => typeof value === 'object' && value !== null && 'prompt' in value).length, 6);

  // Still seven after the last round: the seventh is left out, and the deck says so. Six pictures are generated, not seven.
  const stubborn = scripted(described(), described());
  const six = await writeDeck(template, plan, { llm: stubborn.llm, assets: [rex], maxRepairs: 1 });
  assert.deepEqual(six.deck.slides[0]?.blocks, kinds(['photo', 6]));
  assert.deepEqual(withoutCodes(six.warnings), [{ severity: 'warning', message: 'slide 1: 1 photo was left out, since at most 6 are made for one slide' }]);
  // A slide of up to six photos is told nothing of it.
  const few = scripted({ s1: {} });
  await writeDeck(template, { ...plan, slides: [{ pattern: 'free', purpose: 'Two dogs', blocks: kinds(['photo', 2]) }] }, { llm: few.llm, assets: [rex], maxRepairs: 0 }).catch(() => undefined);
  assert.doesNotMatch(JSON.stringify(few.requests[0]?.schema), /can be made from a description/);
});

// --- parts without content --------------------------------------------------------------------------

const emptyChart = { type: 'column', categories: [], series: [], format: 'number' };

test('a chart the brief has no numbers for is taken off the slide, without asking again', async () => {
  const { llm, requests } = scripted({ s1: { title: 'Sales', b1Caption: 'North', b1Chart: chartOf('2025'), b2Caption: 'South', b2Chart: emptyChart, b3Text: NOTE } });
  const written = await writeDeck(demo, numbersPlan(['chart', 'chart', 'note']), { llm, layout: 'adaptive' });
  assert.equal(requests.length, 1, 'no repair round: asking again brings no numbers');
  // Every chart is asked for: left optional, a model leaves the charts of a slide out. One the brief has no numbers for
  // is given empty, as its description says, and no least number of categories forces numbers to be made up.
  const s1 = schemaOf(requests[0]).properties.s1 as Schema;
  assert.deepEqual(s1.required, ['b1Chart', 'b2Chart', 'b3Text']);
  assert.doesNotMatch(JSON.stringify(s1.properties.b2Chart), /minItems/);
  assert.match(s1.properties.b2Chart?.description ?? '', /^A chart\. When the brief gives no numbers for it, leave it empty\. at most 8 categories, at most 3 series\. A chart of numbers the brief states\. Use only those numbers: never estimate, interpolate or invent a value to complete a series\.$/);
  // The limits are also in the schema, as for the chart of any pattern, and state no least number.
  assert.deepEqual([s1.properties.b2Chart?.properties.categories?.maxItems, s1.properties.b2Chart?.properties.series?.maxItems], [8, 3]);
  // The parts that stay are numbered anew, and the deck is told what was left out: a note on the deck, with no slide to rewrite.
  assert.deepEqual(written.deck.slides[0], { pattern: 'free', blocks: ['chart', 'note'], base: 'list-with-photo', slots: { title: 'Sales', b1Caption: 'North', b1Chart: chartOf('2025'), b2Text: NOTE } });
  assert.deepEqual(withoutCodes(written.warnings.filter((issue) => /left out/.test(issue.message))), [{ severity: 'warning', message: 'slide 1: 1 part had no content and was left out' }]);
  assert.doesNotMatch(written.warnings.map((issue) => issue.message).join(' '), /characters|paragraphs|lines/, 'not a message a rewrite would be started for');
  assert.deepEqual(validateDeck(demo, written.deck, { layout: 'adaptive' }).filter((issue) => issue.severity === 'error'), []);

  // A figure without its number and a table without columns go the same way. A slide left with nothing says its purpose.
  const bare = scripted({ s1: { b1Label: 'riders', b2Table: { header: [], rows: [] }, b3Chart: emptyChart } });
  const said = await writeDeck(demo, numbersPlan(['figure', 'table', 'chart'], 'Our stores did well this year.'), { llm: bare.llm, layout: 'adaptive' });
  assert.equal(bare.requests.length, 1);
  assert.deepEqual(said.deck.slides[0], { pattern: 'free', blocks: ['text'], base: 'list-with-photo', slots: { b1Text: 'Our stores did well this year.' } });
  assert.deepEqual(withoutCodes(said.warnings.filter((issue) => /left out/.test(issue.message))), [{ severity: 'warning', message: 'slide 1: 3 parts had no content and were left out' }]);

  // A table that comes back with its headings and no row shows nothing: it is left empty too. Written by arrays, where
  // every property is asked of every item, that is how a table without numbers comes back.
  const headed = scripted({ s1: { title: 'Sales', b1Table: { header: ['Store', 'Sales'], rows: [] }, b2Text: NOTE } });
  const noted = await writeDeck(demo, numbersPlan(['table', 'note']), { llm: headed.llm, layout: 'adaptive' });
  assert.equal(headed.requests.length, 1);
  assert.deepEqual(noted.deck.slides[0], { pattern: 'free', blocks: ['note'], base: 'list-with-photo', slots: { title: 'Sales', b1Text: NOTE } });
  const rows = scripted({ r01: range(1, 10).map((block) => ({ block, caption: '', header: ['Store', 'Sales'], rows: [] })), r02: [{ block: 11, text: NOTE }], title: 'Sales' });
  const bare10 = await writeDeck(demo, numbersPlan(kinds(['table', 10], ['note', 1])), { llm: rows.llm, layout: 'adaptive' });
  assert.deepEqual([rows.requests.length, bare10.deck.slides[0]?.blocks], [1, ['note']]);
  assert.deepEqual(withoutCodes(bare10.warnings.filter((issue) => /left out/.test(issue.message))), [{ severity: 'warning', message: 'slide 1: 10 parts had no content and were left out' }]);
});

test('a slide left with no part and no purpose still says something, and stops no deck', async () => {
  // A purpose may be empty: an outline is asked for a sentence and may give none. The slide then says the title that
  // was written for it, once, as its text.
  const titled = scripted({ s1: { title: 'Sales in 2025', b1Chart: emptyChart, b2Chart: emptyChart } });
  const one = await writeDeck(demo, numbersPlan(['chart', 'chart'], ''), { llm: titled.llm, layout: 'adaptive' });
  assert.deepEqual(one.deck.slides[0], { pattern: 'free', blocks: ['text'], base: 'list-with-photo', slots: { b1Text: 'Sales in 2025' } });
  assert.deepEqual(validateDeck(demo, one.deck, { layout: 'adaptive' }).filter((issue) => issue.severity === 'error'), []);
  // Without a title either, it says the title of the deck; and a deck without one, the first words of its brief.
  const silent = await writeDeck(demo, numbersPlan(['chart', 'chart'], ' '), { llm: scripted({ s1: {} }).llm, layout: 'adaptive' });
  assert.deepEqual(silent.deck.slides[0]?.slots, { b1Text: 'Stores' });
  const untitled = await writeDeck(demo, { ...numbersPlan(kinds(['figure', 12]), ''), title: '' }, { llm: scripted({ r01: [], title: '' }).llm, layout: 'adaptive', maxRepairs: 0 });
  assert.deepEqual(untitled.deck.slides[0]?.slots, { b1Text: stores.prompt });
  assert.deepEqual(withoutCodes(untitled.warnings.filter((issue) => /left out/.test(issue.message))), [{ severity: 'warning', message: 'slide 1: 12 parts had no content and were left out' }]);
  // Through planDeck, where the outline gives the empty purpose.
  const planned = scripted({ title: 'Stores', slides: [freeEntry('', [['chart', 2]])] }, { s1: { b1Chart: emptyChart } });
  const whole = await planDeck(demo, { ...stores, slideCount: 1 }, { llm: planned.llm, layout: 'adaptive' });
  assert.deepEqual(whole.deck.slides[0]?.slots, { b1Text: 'Stores' });
});

test('a part whose content is still missing after the last repair is taken off: a free slide stops no deck', async () => {
  // The second card never gets its text. It is asked for as any missing slot is, by the name the writer knows it by.
  const answer = { s1: { title: 'Two ways', b1Heading: 'Hours', b1Text: 'From five.', b2Heading: 'Docks', b3Chart: chartOf('2025') } };
  const { llm, requests } = scripted(answer, answer, answer);
  const written = await writeDeck(demo, numbersPlan(['card', 'card', 'chart']), { llm, layout: 'adaptive' });
  assert.equal(requests.length, 3, 'two repair rounds, as for any slide');
  assert.match(promptOf(requests[1]), /\nIt has these problems:\n- s1\.b2Text: this slot is required\n\nReturn the full content again/);
  assert.deepEqual(written.deck.slides[0], { pattern: 'free', blocks: ['card', 'chart'], base: 'list-with-photo', slots: { title: 'Two ways', b1Heading: 'Hours', b1Text: 'From five.', b2Chart: chartOf('2025') } });
  assert.deepEqual(withoutCodes(written.warnings.filter((issue) => /left out/.test(issue.message))), [{ severity: 'warning', message: 'slide 1: 1 part had no content and was left out' }]);

  // A chart that cannot be drawn goes too, and is told apart from one that was never written. A problem on a part
  // that moved up when the empty chart before it went is still told under the name the writer gave it.
  const broken = { s1: { b1Chart: emptyChart, b2Chart: { ...chartOf('2025'), series: [{ name: '2025', values: [1, 2] }] }, b3Text: NOTE } };
  const model = scripted(broken, broken);
  const kept = await writeDeck(demo, numbersPlan(['chart', 'chart', 'note']), { llm: model.llm, layout: 'adaptive', maxRepairs: 1 });
  assert.match(promptOf(model.requests[1]), /\nIt has these problems:\n- s1\.b2Chart: series "2025" has 2 values; the chart has 4 categories\n\nReturn/);
  assert.deepEqual(kept.deck.slides[0], { pattern: 'free', blocks: ['note'], base: 'list-with-photo', slots: { b1Text: NOTE } });
  assert.deepEqual(kept.warnings.filter((issue) => /left out/.test(issue.message)).map((issue) => issue.message), ['slide 1: 1 part had no content and was left out', 'slide 1: 1 part could not be shown and was left out']);
});

test('a plan with a free slide is written whatever is on offer for a new one, and one the template cannot carry says so', async () => {
  // Turned off since the plan was made, or outside the app's list: the slide is written all the same.
  for (const options of [{ free: false }, { patterns: ['cards', 'closing'] }]) {
    const written = await writeDeck(template, cardsPlan(), { llm: scripted({ s1: cardsContent, s2: threeCards, s3: {} }).llm, ...options });
    assert.deepEqual(written.deck.slides[1]?.blocks, ['card', 'card', 'card']);
  }
  // A frame that shows a text of the deck needs that text.
  const plan = numbersPlan(['card', 'card']);
  await assert.rejects(writeDeck(demo, { ...plan, brief: { prompt: 'Dogs' } }, { llm: scripted().llm, layout: 'adaptive' }), (error: unknown) => {
    assert.ok(error instanceof PlanError);
    assert.equal(error.message, 'Slide 1: the free slide needs the deck field "organizationDepartment"');
    return true;
  });
  // A photo where none can be had.
  await assert.rejects(writeDeck(demo, { ...numbersPlan(['photo', 'note']), brief: { ...stores, imageDensity: 'none' } }, { llm: scripted().llm, layout: 'adaptive' }), /^PlanError: Slide 1: the free slide needs a photo, and images are turned off$/);
});

test('the counts of a plan are checked before anything is built for them or written, and those of a deck before its pictures are gathered', async () => {
  // A plan is kept by an app and may be shown for editing: a number in it is not taken at its word.
  // The row of the fixture takes 1 to 5 cards. A number it does not take, or one that is not
  // whole, ends the call before a model is asked; a million items are never worked out.
  const rowOf = (items: number): DeckPlan => ({ ...cardsPlan(), slides: [{ pattern: 'cards', purpose: 'Two things', items }, { pattern: 'closing', purpose: 'Thanks' }] });
  for (const [items, said] of [[6, '6'], [0, '0'], [-2, '-2'], [2.5, '2.5'], [1e6, '1000000'], [1e9, '1000000000'], [Number.NaN, 'NaN'], [Number.POSITIVE_INFINITY, 'Infinity']] as const) {
    const model = scripted();
    await assert.rejects(writeDeck(template, rowOf(items), { llm: model.llm }), (error: unknown) => {
      assert.ok(error instanceof PlanError, String(error));
      assert.equal(error.message, `Slide 1: pattern "cards" takes 1 to 5 cards, not ${said}`);
      return true;
    });
    assert.equal(model.requests.length, 0, 'no model is asked');
  }
  // Every number the row takes is written as it was.
  for (const items of [1, 5]) {
    const cards = Object.fromEntries([['title', 'Two'], ...Array.from({ length: items }, (_, index) => [`card${index + 1}Text`, 'One'])]);
    const written = await writeDeck(template, rowOf(items), { llm: scripted({ s1: cards, s2: {} }).llm });
    assert.deepEqual(written.deck.slides[0], { pattern: 'cards', items, slots: cards });
  }
  // A pattern without a row has no number to check: what a plan says there is passed over, as it was.
  const noRow: DeckPlan = { ...cardsPlan(), slides: [{ pattern: 'closing', purpose: 'Thanks', items: 1e9 }] };
  assert.deepEqual((await writeDeck(template, noRow, { llm: scripted({ s1: {} }).llm })).deck.slides, [{ pattern: 'closing', slots: {} }]);

  // A free slide is held to the parts an answer of a model is held to.
  const many = scripted();
  const full = (count: number): DeckPlan => ({ ...cardsPlan(), slides: [{ pattern: 'closing', purpose: 'Thanks' }, { pattern: 'free', purpose: 'Every logo', blocks: Array<FreeBlockKind>(count).fill('card') }] });
  await assert.rejects(writeDeck(template, full(401), { llm: many.llm }), /^PlanError: Slide 2: a free slide takes at most 400 parts, not 401$/);
  await assert.rejects(writeDeck(template, full(100000), { llm: many.llm }), /^PlanError: Slide 2: a free slide takes at most 400 parts, not 100000$/);
  assert.equal(many.requests.length, 0);

  // The page on generating quotes both messages as the code gives them, the first for the demo template's row of figures.
  const page = readFileSync(join(root, 'docs', 'generating.md'), 'utf8').replace(/\s+/g, ' ');
  const stats: DeckPlan = { version: 1, title: 'Stores', brief: stores, slides: [{ pattern: 'cover', purpose: 'Open' }, { pattern: 'stat-cards', purpose: 'The year in numbers', items: 7 }] };
  const told = await writeDeck(demo, stats, { llm: many.llm, layout: 'adaptive' }).then(() => '', (error: Error) => error.message);
  assert.equal(told, 'Slide 2: pattern "stat-cards" takes 2 to 6 stats, not 7');
  assert.ok(page.includes(`\`${told}\``));
  const crowded = await writeDeck(demo, { ...stats, slides: [{ pattern: 'cover', purpose: 'Open' }, { pattern: 'free', purpose: 'Every store', blocks: Array<FreeBlockKind>(2000).fill('figure') }] }, { llm: many.llm, layout: 'adaptive' }).then(() => '', (error: Error) => error.message);
  assert.ok(page.includes(`\`${crowded}\``), crowded);
  assert.equal(many.requests.length, 0);

  // A deck is gathered pictures for before it is checked. A number of items its row does not take
  // is left for rendering to report: no slots are worked out for it, and no picture is made.
  const deck = (items: number): DeckSpec => ({ version: 1, slides: [{ pattern: 'cards', items, slots: { title: 'Two', card1Text: 'One', card9Icon: { asset: 'gen-1', prompt: 'A dog' } } }] });
  let made = 0;
  const images = { generateImage: async () => (made++, placeholderPhoto(20, 10, [0, 0, 0], [255, 255, 255])) };
  for (const items of [9, 2.5, 1e9]) {
    assert.deepEqual(await resolveAssets(template, deck(items), { images }), {});
    assert.throws(() => renderDeck(template, deck(items)), new RegExp(`pattern "cards" takes 1 to 5 cards, not ${items}`));
  }
  assert.equal(made, 0);
});

test('planDeck goes from a brief to a deck with a free slide in two calls', async () => {
  const { llm, requests } = scripted(outlineOf(entry('cards', 'Two things', { items: 2 }), freeEntry('Three ways in', [['card', 3]], 'auto', 3)), { s1: cardsContent, s2: threeCards });
  const { deck, outline } = await planDeck(template, brief, { llm });
  assert.equal(requests.length, 2);
  assert.deepEqual(outline[1], { pattern: 'free', purpose: 'Three ways in', blocks: ['card', 'card', 'card'], arrange: { columns: 3 } });
  assert.deepEqual(deck.slides[1], { pattern: 'free', blocks: ['card', 'card', 'card'], arrange: { columns: 3 }, base: 'content', slots: threeCards });
});

test('a free slide with pictures is written, gets its pictures and renders like any slide', async () => {
  const png = placeholderPhoto(40, 30, [14, 124, 123], [244, 162, 89]);
  const rex = { id: 'rex', kind: 'photo' as const, description: 'A dog called Rex', data: png };
  const { llm, requests } = scripted(outlineOf(freeEntry('Two dogs', [['photo', 2], ['note', 1]])), {
    s1: { title: 'Two dogs', b1Photo: { prompt: 'A dog in a park', alt: 'A dog' }, b2Photo: { prompt: 'unused', alt: 'Rex', asset: 'rex' }, b3Text: 'Both are looking for a home.' },
  });
  const prompts: string[] = [];
  const result = await generateDeck(template, { ...brief, slideCount: 1 }, { llm, assets: [rex], images: { generateImage: async (request) => (prompts.push(request.prompt), png) } });

  // A photo of a free slide is asked for like the photo of a pattern: a description, or one of the photos supplied.
  const s1 = schemaOf(requests[1]).properties.s1 as Schema;
  // Two photos are told apart by a caption each, as two charts are: a caption is offered, and may be left out.
  assert.deepEqual(Object.keys(s1.properties), ['title', 'b1Photo', 'b1Caption', 'b2Photo', 'b2Caption', 'b3Text']);
  assert.deepEqual(s1.required, ['b1Photo', 'b2Photo', 'b3Text']);
  assert.deepEqual(s1.properties.b1Photo?.required, ['prompt', 'alt']);
  assert.deepEqual(s1.properties.b1Photo?.properties.asset?.enum, ['rex']);
  assert.match(promptOf(requests[1]), /s1: free \(b1 to b2: 2 photos; b3: a note\)\. Two dogs/);
  assert.match(promptOf(requests[1]), /Supplied pictures\./);
  const slots = result.deck.slides[0]?.slots ?? {};
  assert.deepEqual(slots.b2Photo, { asset: 'rex', alt: 'Rex' });
  assert.deepEqual(prompts, ['A dog in a park'], 'one picture is generated; the other was supplied');
  assert.equal(Object.keys(result.assets).length, 2);
  assert.deepEqual(result.warnings, []);
  assert.ok(result.pptx.length > 1000);
});

// --- changing a deck ----------------------------------------------------------------------------------

const storeNames = ['North', 'South', 'East', 'West', 'Centre', 'Online'];
/** A slide of six charts and a note, three side by side, as a deck holds it. */
const storesSlide = (): SlideSpec => ({ ...freeSlide({ title: 'Sales by store', blocks: [...storeNames.map((name) => ({ chart: chartOf(name), caption: name })), { note: NOTE }], columns: 3 }), base: 'list-with-photo' });
const storesDeck = (slide: SlideSpec = storesSlide()): DeckSpec => ({
  version: 1,
  title: 'Stores',
  language: 'en',
  direction: 'ltr',
  brief: stores.prompt,
  fields: stores.fields as Record<string, string>,
  slides: [{ pattern: 'cover', slots: { title: 'Stores' } }, slide, { pattern: 'closing', slots: { title: 'Thanks' } }],
});
/** An entry of the plan of a change as a model answers it where free slides are in play: every property, with nothing in the three that are for free slides alone. */
const stay = (from: number, pattern: string, more: Record<string, unknown> = {}) => ({ from, pattern, change: '', images: false, parts: [], shows_main: 'auto', side_by_side: 0, ...more });
const part = (from: number, kind: string, rewrite = false) => ({ from, kind, rewrite });
const changeOf = (...slides: unknown[]) => ({ summary: 'Done.', language: 'en', title: 'Stores', text_size: 'normal', fields: stores.fields, slides });
/** The six charts and the note as they are, by their numbers. */
const sixAndNote = (more: Record<number, unknown> = {}) => [1, 2, 3, 4, 5, 6, 7].map((number) => more[number] ?? part(number, number === 7 ? 'note' : 'chart'));
const adaptive = { layout: 'adaptive' as const };
const PARTS_NOTE =
  'Only for the layout free, and only when the request changes which parts the slide has, their order, or the content of some of them: every part of the slide after the change, in order. ' +
  "from: the number of the existing part it is, 0 for a new part. rewrite: true when the request changes that part's content. [] when the parts stay as they are, and for every other layout.";
const REPEAT = 'Repeat the current value unless the request is about the arrangement.';
const STORES_LISTED =
  '2. [free: 6 charts and a note; 3 side by side] title: "Sales by store"; ' +
  storeNames.map((name, index) => `${index + 1} chart "${name}" (column, 4 categories)`).join('; ') +
  `; 7 note "${NOTE}"`;

test('a change to a deck lists a free slide by its parts, without their data, and copies one it leaves alone', async () => {
  const deck = storesDeck();
  const free = JSON.stringify(deck.slides[1]);
  const { llm, requests } = scripted(changeOf(stay(3, 'closing'), stay(2, 'free', { side_by_side: 3 }), stay(1, 'cover')));
  const moved = await planEdit(demo, deck, 'Open with the closing slide', { llm, ...adaptive });

  // The deck as the model is shown it: the parts of the free slide by number, each in a few words. No number of any chart.
  const prompt = promptOf(requests[0]);
  assert.ok(prompt.includes(`\n1. [cover] title: "Stores"\n${STORES_LISTED}\n3. [closing] title: "Thanks"\n`), prompt.slice(0, 900));
  assert.doesNotMatch(prompt.split('Requested change:')[0] ?? '', /"values"|\[10,14,12,18\]|Q1/);
  // The rules gain three lines, the last: what "parts" is for, that parts taken off or moved need no words, and what a
  // slide of a pattern hands on when it becomes a free slide. The list of patterns ends with the line for free slides,
  // whose parts a change names in "parts".
  assert.ok(
    prompt.includes(
      '\n- items: for a pattern with a row of items, the number the slide has after the change.\n' +
        '- parts: for a slide with the layout free whose parts change. Leave it [] otherwise.\n' +
        '- A slide with the layout free whose parts are only taken off or moved needs nothing written: leave its change empty unless something on the slide is to be written too, such as another title.\n' +
        '- A slide of a pattern that takes the layout free hands its chart, its table and its photos to the first parts of those kinds that it names. Give such a part rewrite false unless the request changes what it shows.\n\n' +
        'Patterns, for new slides',
    ),
    prompt.slice(-1600),
  );
  assert.match(freeLineOf(requests[0]) ?? '', /^- free: A slide the software lays out itself from the parts you name in "parts", any number of each: text \(.* so plan more only when the request asks for them\.$/);
  assert.ok(prompt.endsWith(freeLineOf(requests[0]) ?? 'no line for free slides'));

  // An entry: what it was, and three properties that every slide answers. Its parts are named one by one.
  const entry = entryOf(requests[0]);
  assert.deepEqual(Object.keys(entry.properties), ['from', 'pattern', 'change', 'images', 'words', 'items', 'parts', 'shows_main', 'side_by_side']);
  assert.deepEqual(entry.required, ['from', 'pattern', 'change', 'images', 'parts', 'shows_main', 'side_by_side']);
  assert.deepEqual(entry.properties.parts, {
    type: 'array',
    description: PARTS_NOTE,
    items: {
      type: 'object',
      properties: { from: { type: 'integer', minimum: 0 }, kind: { type: 'string', enum: ['text', 'list', 'card', 'figure', 'step', 'quote', 'note', 'chart', 'table', 'photo'] }, rewrite: { type: 'boolean' } },
      required: ['from', 'kind', 'rewrite'],
    },
  });
  assert.deepEqual(entry.properties.shows_main, {
    type: 'string',
    enum: ['auto', 'left', 'right', 'top', 'none'],
    description: `Only for free: where the first part stands when it is the main one and larger. auto lets the software decide; none sets all parts alike. ${REPEAT}`,
  });
  assert.deepEqual(entry.properties.side_by_side, { type: 'integer', minimum: 0, description: `Only for free, when the request asks for a number of columns; 0 lets the software decide. ${REPEAT}` });
  assert.ok(entry.properties.pattern?.enum?.includes('free'));
  const declared = JSON.stringify(requests[0]?.schema);
  for (const word of ['anyOf', 'oneOf', '"format"', '"enum":[]', 'maxItems']) assert.ok(!declared.includes(word), word);
  // Room: the plan of a change, what each slide may say of its arrangement, and the parts it may name:
  // those the deck has twice over, and forty that a request may add.
  assert.equal(requests[0]?.maxOutputTokens, (400 + 13 * 160) * 1.5 + 13 * 12 + 16 * (2 * 7 + 40));

  // Left alone, the slide is copied by code: its parts, its arrangement and its frame, byte for byte.
  assert.equal(requests.length, 1, 'nothing is written');
  assert.deepEqual(moved.deck.slides.map((slide) => slide.pattern), ['closing', 'free', 'cover']);
  assert.equal(JSON.stringify(moved.deck.slides[1]), free);
  assert.deepEqual(moved.rewritten, []);

  // An answer that says nothing of free slides at all, as a model without the three properties would give, changes nothing either.
  const bare = scripted(changeOf(...[1, 2, 3].map((from) => ({ from, pattern: deck.slides[from - 1]?.pattern, change: '', images: false }))));
  const same = await planEdit(demo, deck, 'Nothing to change', { llm: bare.llm, ...adaptive });
  assert.equal(JSON.stringify(same.deck.slides), JSON.stringify(deck.slides));

  // Another slide is rewritten: the free slide beside it stays as it is, and the deck is checked with it in place.
  const rewrite = scripted(changeOf(stay(1, 'cover', { change: 'A shorter title' }), stay(2, 'free', { side_by_side: 3 }), stay(3, 'closing')), { s1: { title: 'Shops' } });
  const rewritten = await planEdit(demo, deck, 'Shorten the title of slide 1', { llm: rewrite.llm, ...adaptive });
  assert.deepEqual(rewritten.rewritten, [1]);
  assert.equal(rewritten.deck.slides[0]?.slots.title, 'Shops');
  assert.equal(JSON.stringify(rewritten.deck.slides[1]), free);
  assert.deepEqual(Object.keys(schemaOf(rewrite.requests[1]).properties), ['s1'], 'only the slide that changes is asked for');
  assert.match(promptOf(rewrite.requests[1]), /\n1\. cover, to write\n2\. free, unchanged "Sales by store"\n3\. closing, unchanged "Thanks"\n/);
});

test('a change rewrites one part of a free slide and leaves the others as they are', async () => {
  const deck = storesDeck();
  const worded = 'Online grew fastest of all the stores this year.';
  const { llm, requests } = scripted(
    changeOf(stay(1, 'cover'), stay(2, 'free', { change: 'Say in the note that Online grew fastest', parts: sixAndNote({ 7: part(7, 'note', true) }), side_by_side: 3 }), stay(3, 'closing')),
    { s2: { title: 'Sales by store', b7Text: worded } },
  );
  const result = await planEdit(demo, deck, 'Say in the note that Online grew fastest', { llm, ...adaptive });
  assert.equal(requests.length, 2);
  assert.deepEqual(result.rewritten, [2]);

  // The writer is asked for the note and the title, shown what they say now, and told in a line each what stays.
  const s2 = schemaOf(requests[1]).properties.s2 as Schema;
  assert.deepEqual(Object.keys(s2.properties), ['title', 'b7Text']);
  assert.deepEqual(s2.required, ['b7Text']);
  const task = promptOf(requests[1]);
  assert.ok(
    task.includes(
      '\n\ns2 (free: b1 to b6: 6 charts; b7: a note). Change: Say in the note that Online grew fastest\n' +
        `What stays as it is and is not to be written: ${storeNames.map((name, index) => `${index + 1} chart "${name}" (column, 4 categories)`).join('; ')}\n` +
        `Current content: ${JSON.stringify({ title: 'Sales by store', b7Text: NOTE })}\n\nFill every required property.`,
    ),
    task.slice(-700),
  );
  assert.doesNotMatch(task, /\[10,14,12,18\]|"values"/, 'no number of a chart that stays');

  // The slide: its charts exactly as they were, and the new note.
  const before = deck.slides[1] as SlideSpec;
  const after = result.deck.slides[1] as SlideSpec;
  for (const name of Object.keys(before.slots).filter((slot) => slot.endsWith('Chart'))) assert.equal(JSON.stringify(after.slots[name]), JSON.stringify(before.slots[name]), name);
  assert.deepEqual(after, { pattern: 'free', blocks: before.blocks, arrange: { columns: 3 }, base: 'list-with-photo', slots: { ...before.slots, b7Text: worded } });
  assert.deepEqual(result.warnings.filter((issue) => issue.slot !== undefined || issue.severity === 'error'), []);

  // A change that says nothing of the parts rewrites the content of all of them, each from what it holds: a translation does.
  const french = { ...Object.fromEntries(Object.entries(before.slots)), title: 'Ventes par magasin', b7Text: 'Chaque magasin a vendu le plus au dernier trimestre.' };
  const translate = scripted(
    { ...changeOf(stay(1, 'cover'), stay(2, 'free', { side_by_side: 3 }), stay(3, 'closing')), language: 'fr' },
    { s1: { title: 'Magasins' }, s2: french, s3: { title: 'Merci' } },
  );
  const translated = await planEdit(demo, deck, 'In French', { llm: translate.llm, ...adaptive });
  assert.deepEqual(translated.rewritten, [1, 2, 3]);
  const all = schemaOf(translate.requests[1]).properties.s2 as Schema;
  assert.deepEqual(Object.keys(all.properties), ['title', ...[1, 2, 3, 4, 5, 6].flatMap((number) => [`b${number}Caption`, `b${number}Chart`]), 'b7Text']);
  assert.ok(promptOf(translate.requests[1]).includes(`\ns2 (free: b1 to b6: 6 charts; b7: a note). Change: Translate into the new deck language.\nCurrent content: ${JSON.stringify(before.slots)}\n`));
  assert.equal(translated.deck.slides[1]?.slots.title, 'Ventes par magasin');
  assert.deepEqual(translated.deck.slides[1]?.arrange, { columns: 3 });
});

test('charts are added to a free slide without the model seeing or returning the numbers of the others', async () => {
  const deck = storesDeck();
  const before = deck.slides[1] as SlideSpec;
  const airport = { type: 'column', categories: quarters, series: [{ name: 'Airport', values: [10, 12, 14, 18] }] };
  const harbour = { type: 'column', categories: quarters, series: [{ name: 'Harbour', values: [18, 14, 12, 10] }] };
  // The two new charts go before the note.
  const parts = [...sixAndNote().slice(0, 6), part(0, 'chart', true), part(0, 'chart', true), part(7, 'note')];
  const { llm, requests } = scripted(
    changeOf(stay(1, 'cover'), stay(2, 'free', { change: 'Add charts for Airport and Harbour', parts, side_by_side: 4 }), stay(3, 'closing')),
    { s2: { title: 'Sales by store', b7Caption: 'Airport', b7Chart: airport, b8Caption: 'Harbour', b8Chart: harbour } },
  );
  const result = await planEdit(demo, deck, 'Add charts for Airport (10, 12, 14, 18) and Harbour (18, 14, 12, 10), four side by side', { llm, ...adaptive });

  const s2 = schemaOf(requests[1]).properties.s2 as Schema;
  assert.deepEqual(Object.keys(s2.properties), ['title', 'b7Caption', 'b7Chart', 'b8Caption', 'b8Chart']);
  const asked = JSON.stringify(requests[1]);
  assert.ok(!asked.includes('[10,14,12,18]') && !asked.includes('b1Chart') && !asked.includes('b9Text'), 'nothing of the charts and the note that stay');
  assert.match(promptOf(requests[1]), /\ns2 \(free: b1 to b8: 8 charts; b9: a note\)\. Change: Add charts for Airport and Harbour\nWhat stays as it is and is not to be written: 1 chart "North" \(column, 4 categories\); .*; 6 chart "Online" \(column, 4 categories\); 9 note "Every store/);

  // Nine parts: the six charts under their numbers, the two new ones, and the note, which is part 9 now.
  const after = result.deck.slides[1] as SlideSpec;
  assert.deepEqual(after.blocks, [...Array<string>(8).fill('chart'), 'note']);
  assert.deepEqual(after.arrange, { columns: 4 });
  for (const number of [1, 2, 3, 4, 5, 6]) {
    assert.equal(JSON.stringify(after.slots[`b${number}Chart`]), JSON.stringify(before.slots[`b${number}Chart`]));
    assert.equal(after.slots[`b${number}Caption`], before.slots[`b${number}Caption`]);
  }
  assert.deepEqual([after.slots.b7Chart, after.slots.b8Caption, after.slots.b9Text, after.slots.title], [airport, 'Harbour', NOTE, 'Sales by store']);
  assert.equal(Object.keys(after.slots).length, 18);
  // The numbers of a chart are held against the brief where the chart was written now: the request joined it, and they are in it.
  assert.deepEqual(result.warnings.filter((issue) => issue.slot !== undefined), []);
  assert.ok(renderDeck(demo, result.deck, adaptive).pptx.length > 1000);

  // A chart whose numbers the request does not give is said, and of the charts that stay nothing is said again.
  const unknown = scripted(
    changeOf(stay(1, 'cover'), stay(2, 'free', { change: 'Add a chart for Airport', parts: [...sixAndNote().slice(0, 6), part(0, 'chart', true), part(7, 'note')], side_by_side: 3 }), stay(3, 'closing')),
    { s2: { title: 'Sales by store', b7Caption: 'Airport', b7Chart: { ...airport, series: [{ name: 'Airport', values: [1, 2, 3, 4] }] } } },
  );
  const guessed = await planEdit(demo, { ...deck, brief: 'Sales of every store' }, 'Add a chart for Airport', { llm: unknown.llm, ...adaptive });
  assert.deepEqual(withoutCodes(guessed.warnings.filter((issue) => issue.slot !== undefined)), [{ severity: 'warning', slide: 2, slot: 'b7Chart', message: 'chart values not stated in the brief: 1, 2, 3, 4. Check them before using the deck.' }]);
});

test('a free slide of many parts is listed in short, and a change to it is written by calls of its own for the parts that change', async () => {
  const sixty = freeSlide({ title: 'Every store', blocks: range(1, 60).map((number) => ({ chart: chartOf(`Store ${number}`), caption: `Store ${number}` })) });
  const deck = storesDeck({ ...sixty, base: 'list-with-photo' });
  const parts = [...range(1, 60).map((number) => part(number, 'chart', number === 7)), part(0, 'chart', true), part(0, 'chart', true)];
  const { llm, requests } = scripted(changeOf(stay(1, 'cover'), stay(2, 'free', { change: 'Call store 7 "Harbour", and add stores 61 and 62', parts }), stay(3, 'closing')), {
    r01: [{ block: 7, caption: 'Harbour', categories: quarters, format: 'number', type: 'column', values: [10, 14, 12, 18] }, sparkItem(61), sparkItem(62)],
    title: 'Every store',
  });
  const result = await planEdit(demo, deck, 'Call store 7 "Harbour", and add stores 61 and 62 with the numbers 10, 14, 12 and 18', { llm, ...adaptive });
  assert.equal(requests.length, 2, 'the plan of the change, and one call for the three parts');

  // Sixty parts are not listed one by one: the first three and the last of the run, and how many lie between.
  assert.ok(
    promptOf(requests[0]).includes(
      '\n2. [free: 60 charts] title: "Every store"; 1 chart "Store 1" (column, 4 categories); 2 chart "Store 2" (column, 4 categories); 3 chart "Store 3" (column, 4 categories); ... parts 4 to 59: 56 more charts; 60 chart "Store 60" (column, 4 categories)\n',
    ),
  );
  // Room for the plan to name every part of the slide, twice over, and those the request speaks of.
  assert.equal(requests[0]?.maxOutputTokens, (400 + 13 * 160) * 1.5 + 13 * 12 + 16 * (2 * 60 + (7 + 61 + 62 + 10 + 14 + 12 + 18)));

  // The call writes the parts that change, and no other. Like parts among them are one array, whatever stands between them on the slide.
  const [, call] = requests as [JsonRequest, JsonRequest];
  assert.deepEqual(Object.keys(schemaOf(call).properties), ['r01', 'title']);
  assert.match(schemaOf(call).properties.r01?.description ?? '', /^Exactly 3 charts: parts 7, 61 and 62\. A small chart\. When the brief gives no numbers for it, leave it empty\. at most 12 categories\. A chart of numbers the brief states\./);
  assert.equal(call.effort, 'low', 'charts alone are numbers to set down');
  assert.ok((call.maxOutputTokens ?? 0) < 3000, 'room for three charts and a title, not for sixty-two');
  const asked = promptOf(call);
  assert.ok(
    asked.endsWith(
      '\n\nWrite only the parts asked for below. Where they have current content, start from it and change only what is asked: keep every other text word for word.\n\n' +
        'Write parts 7, 61 and 62 of slide 2 (free: 62 charts): Call store 7 "Harbour", and add stores 61 and 62\n- r01: parts 7, 61 and 62, 3 charts\n' +
        'Part N is the item whose "block" is N. Where an item has nothing for a property, give "" or [].\n' +
        // What the parts to write hold now, in the shape the call answers in.
        `Current content: ${JSON.stringify({ r01: [{ block: 7, caption: 'Store 7', categories: quarters, format: 'number', type: 'column', values: [10, 14, 12, 18] }], title: 'Every store' })}\n` +
        'What stays as it is and is not to be written: 1 chart "Store 1" (column, 4 categories); 2 chart "Store 2" (column, 4 categories); 3 chart "Store 3" (column, 4 categories); ... parts 4 and 5: 2 more charts; 6 chart "Store 6" (column, 4 categories); ' +
        '8 chart "Store 8" (column, 4 categories); 9 chart "Store 9" (column, 4 categories); 10 chart "Store 10" (column, 4 categories); ... parts 11 to 59: 49 more charts; 60 chart "Store 60" (column, 4 categories)',
    ),
    asked.slice(-1200),
  );
  assert.ok(asked.includes('\n\nThe deck after the change:\n1. cover, unchanged "Stores"\n2. free, to write\n3. closing, unchanged "Thanks"\n\n'));

  const after = result.deck.slides[1] as SlideSpec;
  assert.equal(after.blocks?.length, 62);
  assert.equal(after.slots.b7Caption, 'Harbour');
  assert.deepEqual(after.slots.b62Chart, { type: 'line', categories: quarters, series: [{ name: 'Store 62', values: [10, 14, 12, 18] }] });
  for (const number of [1, 6, 8, 60]) assert.equal(JSON.stringify(after.slots[`b${number}Chart`]), JSON.stringify(sixty.slots[`b${number}Chart`]));
  assert.deepEqual(result.rewritten, [2]);
  assert.ok(renderDeck(demo, result.deck, adaptive).pptx.length > 1000);
});

test('a photo is taken off a free slide by code, and the right ones stay', async () => {
  const photo = (name: string) => ({ asset: `gen-${name}`, prompt: `A dog called ${name}`, alt: name });
  const gallery: SlideSpec = { ...freeSlide({ title: 'Three dogs', blocks: [{ photo: photo('Rex') }, { photo: photo('Fido') }, { photo: photo('Luna') }], lead: 'none' }), base: 'content' };
  const deck: DeckSpec = { version: 1, title: 'Dogs', language: 'en', direction: 'ltr', brief: 'A deck about dogs', fields, slides: [{ pattern: 'cards', items: 2, slots: cardsContent }, gallery, { pattern: 'closing', slots: {} }] };
  const answer = (...slides: unknown[]) => ({ summary: 'Done.', language: 'en', title: 'Dogs', text_size: 'normal', fields, slides });

  // The plan names the photos that stay and has no words for a writer, as its rules ask of it.
  const { llm, requests } = scripted(answer(stay(1, 'cards'), stay(2, 'free', { parts: [part(1, 'photo'), part(3, 'photo')], shows_main: 'none' }), stay(3, 'closing')));
  const result = await planEdit(template, deck, 'Remove the second photo of slide 2', { llm });
  assert.match(promptOf(requests[0]), /\n2\. \[free: 3 photos; all parts alike\] title: "Three dogs"; 1 photo of "A dog called Rex"; 2 photo of "A dog called Fido"; 3 photo of "A dog called Luna"\n/);
  assert.equal(requests.length, 1, 'no part is new or written again, and nothing is said to write: nothing is asked of a writer');
  // The first and the third photo, the third under the number 2 now. The slide counts as changed.
  assert.deepEqual(result.deck.slides[1], { pattern: 'free', blocks: ['photo', 'photo'], arrange: { lead: 'none' }, base: 'content', slots: { title: 'Three dogs', b1Photo: photo('Rex'), b2Photo: photo('Luna') } });
  assert.deepEqual(result.rewritten, [2]);
  assert.deepEqual(result.warnings, []);
  assert.deepEqual(validateDeck(template, result.deck), []);

  // Moved by code too: the last photo first.
  const turn = scripted(answer(stay(1, 'cards'), stay(2, 'free', { parts: [part(3, 'photo'), part(1, 'photo'), part(2, 'photo')], shows_main: 'none' }), stay(3, 'closing')));
  const turned = await planEdit(template, deck, 'Show Luna first on slide 2', { llm: turn.llm });
  assert.equal(turn.requests.length, 1);
  assert.deepEqual([1, 2, 3].map((number) => turned.deck.slides[1]?.slots[`b${number}Photo`]), [photo('Luna'), photo('Rex'), photo('Fido')]);

  // A plan that has words for the slide and names no part to write means its title, which it has no other way to name:
  // the photo goes by code as before, and one small call writes the title alone. No photo is offered or described again.
  const renamed = scripted(answer(stay(1, 'cards'), stay(2, 'free', { change: 'Remove the second photo and call the slide "Two dogs"', parts: [part(1, 'photo'), part(3, 'photo')], shows_main: 'none' }), stay(3, 'closing')), { s2: { title: 'Two dogs' } });
  const two = await planEdit(template, deck, 'Remove the second photo of slide 2 and call the slide "Two dogs"', { llm: renamed.llm });
  assert.equal(renamed.requests.length, 2);
  const titleOnly = schemaOf(renamed.requests[1]).properties.s2 as Schema;
  assert.deepEqual([Object.keys(titleOnly.properties), titleOnly.required], [['title'], []]);
  assert.ok(
    promptOf(renamed.requests[1]).includes(
      '\ns2 (free: b1 to b2: 2 photos). Change: Remove the second photo and call the slide "Two dogs"\n' +
        'What stays as it is and is not to be written: 1 photo of "A dog called Rex"; 2 photo of "A dog called Luna"\nCurrent content: {"title":"Three dogs"}\n',
    ),
    promptOf(renamed.requests[1]).slice(-500),
  );
  assert.deepEqual(two.deck.slides[1], { pattern: 'free', blocks: ['photo', 'photo'], arrange: { lead: 'none' }, base: 'content', slots: { title: 'Two dogs', b1Photo: photo('Rex'), b2Photo: photo('Luna') } });
  assert.deepEqual([two.rewritten, two.warnings], [[2], []]);
  // A writer that says nothing of the title leaves it as it was.
  const silent = scripted(answer(stay(1, 'cards'), stay(2, 'free', { change: 'Remove the second photo', parts: [part(1, 'photo'), part(3, 'photo')], shows_main: 'none' }), stay(3, 'closing')), { s2: {} });
  assert.equal((await planEdit(template, deck, 'Remove the second photo of slide 2', { llm: silent.llm })).deck.slides[1]?.slots.title, 'Three dogs');

  // A photo is made anew only where the change is about pictures, and then only the one that is to be written again.
  const cat = scripted(answer(stay(1, 'cards'), stay(2, 'free', { change: 'A cat in place of Fido', images: true, parts: [part(1, 'photo'), part(2, 'photo', true), part(3, 'photo')], shows_main: 'none' }), stay(3, 'closing')), {
    s2: { title: 'Two dogs and a cat', b2Photo: { prompt: 'A cat on a sofa', alt: 'A cat' } },
  });
  const swapped = await planEdit(template, deck, 'A cat in place of Fido on slide 2', { llm: cat.llm });
  // The new picture can get a caption: one is offered for the photos of a slide that has several.
  assert.deepEqual(Object.keys((schemaOf(cat.requests[1]).properties.s2 as Schema).properties), ['title', 'b2Photo', 'b2Caption']);
  assert.deepEqual((schemaOf(cat.requests[1]).properties.s2 as Schema).required, ['b2Photo']);
  assert.ok(
    promptOf(cat.requests[1]).includes(
      '\ns2 (free: b1 to b3: 3 photos). Change: A cat in place of Fido\nWhat stays as it is and is not to be written: 1 photo of "A dog called Rex"; 3 photo of "A dog called Luna"\n' +
        'Current content: {"title":"Three dogs","b2Photo":{"prompt":"A dog called Fido","alt":"Fido"}}\n',
    ),
  );
  assert.equal((swapped.deck.slides[1]?.slots.b2Photo as { prompt?: string }).prompt, 'A cat on a sofa');
  assert.deepEqual([swapped.deck.slides[1]?.slots.b1Photo, swapped.deck.slides[1]?.slots.b3Photo, swapped.deck.slides[1]?.slots.title], [photo('Rex'), photo('Luna'), 'Two dogs and a cat']);

  // The slide takes a pattern: it is laid out again like any slide, its first photo goes along, and what it held is the writer's material.
  const back = scripted(answer(stay(1, 'cards'), stay(2, 'content', { change: 'One photo, with the three names as points' }), stay(3, 'closing')), { s2: { title: 'Three dogs', body: ['Rex', 'Fido', 'Luna'], note: 'All three are looking for a home.' } });
  const plain = await planEdit(template, deck, 'Make slide 2 an ordinary slide with one photo', { llm: back.llm });
  assert.deepEqual(plain.deck.slides[1], { pattern: 'content', slots: { photo: photo('Rex'), title: 'Three dogs', body: ['Rex', 'Fido', 'Luna'], note: 'All three are looking for a home.' } });
  assert.match(promptOf(back.requests[1]), /\ns2 \(content\)\. Change: One photo, with the three names as points\nContent of the slide it replaces: \{"title":"Three dogs","b1Photo":\{"prompt":"A dog called Rex","alt":"Rex"\},"b2Photo":/);
});

test('a part that is written again keeps its picture, and keeps what it had where nothing usable comes back for it', async () => {
  const icons = ['clock', 'lamp', 'people'].map((id) => ({ id, description: id, data: new Uint8Array([1]) }));
  const card = (heading: string, text: string, icon: string) => ({ card: { heading, text, icon } });
  const ways: SlideSpec = {
    ...freeSlide({ title: 'Three ways in', blocks: [card('Longer hours', 'Bikes unlock from five.', 'clock'), card('Safer docks', 'Every station gets a light.', 'lamp'), card('Group passes', 'One pass for five riders.', 'people')] }),
    base: 'content',
  };
  const deck: DeckSpec = { version: 1, title: 'Dogs', language: 'en', direction: 'ltr', brief: 'A deck about bikes', fields, slides: [ways, { pattern: 'closing', slots: {} }] };
  const answer = (more: Record<string, unknown>) => ({ summary: 'Done.', language: 'en', title: 'Dogs', text_size: 'normal', fields, slides: [stay(1, 'free', { parts: [part(1, 'card'), part(2, 'card', true), part(3, 'card')], ...more }), stay(2, 'closing')] });
  const brighter = { title: 'Three ways in', b2Heading: 'Brighter docks', b2Text: 'Two lights at every station.' };

  // The second card is written again. Its icon is a picture: it stays, and is not asked for.
  const { llm, requests } = scripted(answer({ change: 'Two lights at every station' }), { s1: brighter });
  const result = await planEdit(template, deck, 'Slide 1: every station gets two lights', { llm, icons });
  assert.match(promptOf(requests[0]), /\n1\. \[free: 3 cards\] title: "Three ways in"; 1 card \(icon "clock"\) "Longer hours": "Bikes unlock from five\."; 2 card \(icon "lamp"\) "Safer docks": "Every station gets a light\."; 3 card /);
  assert.deepEqual(Object.keys((schemaOf(requests[1]).properties.s1 as Schema).properties), ['title', 'b2Heading', 'b2Text']);
  // The card that is written is not among what "is not to be written": what it keeps is said in a line of its own.
  assert.ok(
    promptOf(requests[1]).includes(
      '\nWhat stays as it is and is not to be written: 1 card (icon "clock") "Longer hours": "Bikes unlock from five."; 3 card (icon "people") "Group passes": "One pass for five riders."\n' +
        'Of the parts that are written, this stays as it is: 2 card (icon "lamp")\n' +
        'Current content: {"title":"Three ways in","b2Heading":"Safer docks","b2Text":"Every station gets a light."}\n',
    ),
    promptOf(requests[1]).slice(-600),
  );
  assert.deepEqual(result.deck.slides[0]?.slots, { ...ways.slots, b2Heading: 'Brighter docks', b2Text: 'Two lights at every station.' });
  // With the change about pictures, the icon is asked for too.
  const pictures = scripted(answer({ change: 'Two lights at every station, with another icon', images: true }), { s1: { ...brighter, b2Icon: 'clock' } });
  const other = await planEdit(template, deck, 'Slide 1: every station gets two lights, and another icon', { llm: pictures.llm, icons });
  assert.deepEqual(Object.keys((schemaOf(pictures.requests[1]).properties.s1 as Schema).properties), ['title', 'b2Icon', 'b2Heading', 'b2Text']);
  assert.deepEqual([other.deck.slides[0]?.slots.b1Icon, other.deck.slides[0]?.slots.b2Icon], [{ asset: 'clock' }, { asset: 'clock' }]);

  // The writer returns the heading and not the text, or nothing at all for the card: what is missing is what the card had.
  // No part leaves a slide because a change left it unsaid. The plan asked for the card to be written, so what did not
  // come is asked for again like any missing slot, and what still has not come is told: the change was not made.
  for (const [written, expected, lines] of [
    [{ title: 'Three ways in', b2Heading: 'Brighter docks' }, { b2Heading: 'Brighter docks', b2Text: 'Every station gets a light.' }, ['b2Text']],
    [{ title: 'Three ways in' }, { b2Heading: 'Safer docks', b2Text: 'Every station gets a light.' }, ['b2Heading', 'b2Text']],
  ] as const) {
    const partly = scripted(answer({ change: 'Two lights at every station' }), { s1: written }, { s1: written });
    const kept = await planEdit(template, deck, 'Slide 1: every station gets two lights', { llm: partly.llm, icons, maxRepairs: 1 });
    assert.equal(partly.requests.length, 3);
    assert.ok(
      promptOf(partly.requests[2]).includes(`\nIt has these problems:\n${lines.map((slot) => `- s1.${slot}: nothing was written for it, and it is to be written again`).join('\n')}\n\nReturn the full content again`),
      promptOf(partly.requests[2]).slice(-300),
    );
    assert.deepEqual(kept.deck.slides[0]?.slots, { ...ways.slots, ...expected });
    assert.deepEqual(kept.deck.slides[0]?.blocks, ['card', 'card', 'card']);
    assert.deepEqual(withoutCodes(kept.warnings), [{ severity: 'warning', message: 'slide 1: nothing was written for 1 part, and it stays as it was' }]);
    assert.doesNotMatch(kept.warnings.map((issue) => issue.message).join(' '), /characters|paragraphs|lines/);
  }
  // Asked again, the writer gives the text: the card is as the change wants it, and nothing is left to tell.
  const late = scripted(answer({ change: 'Two lights at every station' }), { s1: { title: 'Three ways in', b2Heading: 'Brighter docks' } }, { s1: { b2Text: 'Two lights at every station.' } });
  const mendedCard = await planEdit(template, deck, 'Slide 1: every station gets two lights', { llm: late.llm, icons });
  // The second answer holds the text alone: the heading of the first answer is not lost for that.
  assert.deepEqual(mendedCard.deck.slides[0]?.slots, { ...ways.slots, b2Heading: 'Brighter docks', b2Text: 'Two lights at every station.' });
  assert.deepEqual([late.requests.length, mendedCard.warnings], [3, []]);
  // The title too: left unsaid it is the one the slide had, and an empty one is how the slide is given none.
  const unsaid = scripted(answer({ change: 'Two lights at every station' }), { s1: { b2Heading: 'Brighter docks', b2Text: 'Two lights at every station.' } });
  assert.equal((await planEdit(template, deck, 'Slide 1: every station gets two lights', { llm: unsaid.llm, icons })).deck.slides[0]?.slots.title, 'Three ways in');
  const untitled = scripted(answer({ change: 'No title, and two lights at every station' }), { s1: { ...brighter, title: '' } });
  assert.ok(!('title' in ((await planEdit(template, deck, 'Slide 1: no title, and two lights at every station', { llm: untitled.llm, icons })).deck.slides[0]?.slots ?? {})));
});

test('a chart that is written again and comes back unusable, or not at all, stays as it was', async () => {
  const deck = storesDeck();
  const before = deck.slides[1] as SlideSpec;
  const plan = changeOf(stay(1, 'cover'), stay(2, 'free', { change: 'North sold 20 in the last quarter', parts: sixAndNote({ 1: part(1, 'chart', true) }), side_by_side: 3 }), stay(3, 'closing'));

  // The writer leaves the chart out and returns its caption: the chart is the one the slide had, and nothing is taken off.
  // A chart is required of the writer, as every chart of a free slide is. One the plan asks to be written again and
  // that does not come is asked for again, and when it still does not come the deck says that the change was not made.
  const unsaid = { s2: { title: 'Sales by store', b1Caption: 'North, revised' } };
  const { llm, requests } = scripted(plan, unsaid, unsaid);
  const result = await planEdit(demo, deck, 'North sold 20 in the last quarter', { llm, ...adaptive, maxRepairs: 1 });
  assert.equal(requests.length, 3);
  assert.deepEqual((schemaOf(requests[1]).properties.s2 as Schema).required, ['b1Chart']);
  assert.ok(promptOf(requests[1]).includes(`\nCurrent content: ${JSON.stringify({ title: 'Sales by store', b1Caption: 'North', b1Chart: before.slots.b1Chart })}\n`), 'the chart to write again is shown with its numbers');
  // The line on what stays speaks of the six parts that are not written, and not of the chart that is.
  assert.match(promptOf(requests[1]), /\nWhat stays as it is and is not to be written: 2 chart "South" \(column, 4 categories\); .*; 7 note "Every store[^\n]*\nCurrent content: /);
  assert.ok(promptOf(requests[2]).includes('\nIt has these problems:\n- s2.b1Chart: nothing was written for it, and it is to be written again\n\nReturn the full content again'), promptOf(requests[2]).slice(-300));
  assert.deepEqual(result.deck.slides[1]?.slots, { ...before.slots, b1Caption: 'North, revised' });
  assert.deepEqual(result.deck.slides[1]?.blocks, before.blocks);
  assert.deepEqual(withoutCodes(result.warnings.filter((issue) => /stays?/.test(issue.message))), [{ severity: 'warning', message: 'slide 2: nothing was written for 1 part, and it stays as it was' }]);

  // A chart that cannot be drawn is sent back like any; still wrong after the last round, the old one stays and the deck says so.
  const broken = { s2: { title: 'Sales by store', b1Caption: 'North', b1Chart: { type: 'column', categories: quarters, series: [{ name: 'North', values: [10, 14, 20] }] } } };
  const model = scripted(plan, broken, broken);
  const kept = await planEdit(demo, deck, 'North sold 20 in the last quarter', { llm: model.llm, ...adaptive, maxRepairs: 1 });
  assert.match(promptOf(model.requests[2]), /\nIt has these problems:\n- s2\.b1Chart: series "North" has 3 values; the chart has 4 categories\n\nReturn/);
  assert.deepEqual(kept.deck.slides[1]?.slots, before.slots);
  assert.deepEqual(withoutCodes(kept.warnings.filter((issue) => /stays?/.test(issue.message))), [{ severity: 'warning', message: 'slide 2: what was written for 1 part could not be shown, and it stays as it was' }]);
  assert.doesNotMatch(kept.warnings.map((issue) => issue.message).join(' '), /characters|paragraphs|lines/);

  // Written by a call of its own, a part of which no item comes back keeps all it had; a new part of which none comes back is left out.
  const twelve: SlideSpec = { ...freeSlide({ title: 'Every store', blocks: range(1, 12).map((number) => ({ chart: chartOf(`Store ${number}`), caption: `Store ${number}` })) }), base: 'list-with-photo' };
  const parts = [...range(1, 12).map((number) => part(number, 'chart', number <= 2)), part(0, 'chart', true)];
  const apart = scripted(changeOf(stay(1, 'cover'), stay(2, 'free', { change: 'New numbers for stores 1 and 2, and a store 13', parts }), stay(3, 'closing')), { r01: [chartItem(1, { caption: 'Store one' })], title: 'Every store' });
  const some = await planEdit(demo, storesDeck(twelve), 'New numbers for stores 1 and 2, and a store 13', { llm: apart.llm, ...adaptive, maxRepairs: 0 });
  const after = some.deck.slides[1] as SlideSpec;
  assert.equal(after.blocks?.length, 12, 'the twelve it had: the thirteenth was never written');
  assert.deepEqual([after.slots.b1Caption, after.slots.b2Caption, after.slots.b2Chart], ['Store one', 'Store 2', twelve.slots.b2Chart]);
  assert.deepEqual(
    some.warnings.filter((issue) => issue.slide === undefined && /part/.test(issue.message)).map((issue) => issue.message).sort(),
    ['slide 2: 1 part had no content and was left out', 'slide 2: nothing was written for 1 part, and it stays as it was'],
  );
});

test('charts that hold several series keep them when a slide of many charts is written again', async () => {
  // At this size a chart is drawn as its line or bars alone, and a new one is asked for as one list of numbers, named by
  // its caption. A chart the deck holds may have more all the same: it was written when the slide had fewer charts, or by code.
  const pair = (): ChartValue => ({ type: 'column', categories: quarters, series: [{ name: '2024', values: [9, 12, 11, 15] }, { name: '2025', values: [10, 14, 12, 18] }] });
  // A writer that returns every item exactly as it was shown it.
  const echo = (request: JsonRequest): unknown =>
    's1' in schemaOf(request).properties ? { s1: { title: 'Magasins' }, s3: { title: 'Merci' } } : JSON.parse(/\nCurrent content: (.*)/.exec(request.prompt)?.[1] ?? '{}');
  // Sixty charts, in French: every chart is written again, by calls of their own. `requests` counts every call of the change.
  const translated = async (chart: (number: number) => ChartValue) => {
    const slide: SlideSpec = { ...freeSlide({ title: 'Every store', blocks: range(1, 60).map((number) => ({ chart: chart(number), caption: `Store ${number}` })) }), base: 'list-with-photo' };
    const { llm, requests } = scripted({ ...changeOf(stay(1, 'cover'), stay(2, 'free'), stay(3, 'closing')), language: 'fr' }, echo, echo, echo, echo);
    const result = await planEdit(demo, storesDeck(slide), 'In French', { llm, ...adaptive });
    const calls = requests.filter((request) => /Write parts \d+ to \d+ of slide 2/.test(request.prompt));
    return { calls, ranges: calls.map((request) => /Write (parts \d+ to \d+) of slide 2/.exec(request.prompt)?.[1]), requests: requests.length, after: result.deck.slides[1] as SlideSpec, warnings: result.warnings };
  };
  const itemsOf = (request: JsonRequest | undefined) => Object.keys((schemaOf(request).properties.r01?.items as Schema).properties);

  // Two series in every chart: the run is asked for, and shown, in the shape that has them. Nothing is lost.
  const two = await translated(pair);
  // A chart of two series is a longer answer than a chart of one: three calls where sixty charts of one series take two.
  // No call is made twice: the plan of the change, the three, and the two other slides.
  assert.deepEqual([two.ranges, two.requests], [['parts 1 to 20', 'parts 21 to 40', 'parts 41 to 60'], 5]);
  // A new chart of this size would be told one series. These hold two and are asked for as they are: the limit said is
  // no lower than what the charts of the array hold, so that no writer drops a series to meet it.
  assert.equal(
    schemaOf(two.calls[1]).properties.r01?.description,
    'Exactly 20 charts: parts 21 to 40. A small chart. When the brief gives no numbers for it, leave it empty. at most 12 categories, at most 2 series. ' +
      'A chart of numbers the brief states. Use only those numbers: never estimate, interpolate or invent a value to complete a series.',
  );
  for (const call of two.calls) assert.deepEqual(itemsOf(call), ['block', 'caption', 'categories', 'format', 'series', 'type']);
  assert.ok(promptOf(two.calls[0]).includes(`\nCurrent content: {"r01":[{"block":1,"caption":"Store 1","categories":["Q1","Q2","Q3","Q4"],"format":"number","series":${JSON.stringify(pair().series)},"type":"column"},`));
  for (const number of range(1, 60)) assert.deepEqual(two.after.slots[`b${number}Chart`], pair(), `chart ${number}`);
  assert.deepEqual(two.warnings.filter((issue) => issue.slide === undefined && /part/.test(issue.message)), []);
  // The deck is told what any check of it tells, before the change as after: a chart of this size reads well with one
  // series. That is a warning on each chart, and no problem to send back: the charts are as large as the slide had them.
  const large = two.warnings.filter((issue) => issue.message === 'the chart has 2 series; the slot fits 1');
  assert.deepEqual([large.length, large[0]?.slide, large[0]?.slot, large[59]?.slot], [60, 2, 'b1Chart', 'b60Chart']);
  const asItWas = storesDeck({ ...freeSlide({ title: 'Every store', blocks: range(1, 60).map((number) => ({ chart: pair(), caption: `Store ${number}` })) }), base: 'list-with-photo' });
  assert.equal(validateDeck(demo, asItWas, adaptive).filter((issue) => issue.message === 'the chart has 2 series; the slot fits 1').length, 60);

  // One series whose name is not the chart's caption: one list of numbers would give it the caption for a name.
  const named = await translated(() => chartOf('2025'));
  assert.deepEqual([named.ranges, named.requests], [['parts 1 to 30', 'parts 31 to 60'], 4]);
  assert.ok(itemsOf(named.calls[0]).includes('series'));
  assert.match(schemaOf(named.calls[0]).properties.r01?.description ?? '', /\. at most 12 categories, at most 1 series\. A chart of numbers/, 'an item that has series is told how many');
  assert.deepEqual(named.after.slots.b7Chart, chartOf('2025'));
  // One series named by the caption is what one list of numbers gives back: such charts are asked for as new ones are.
  const plain = await translated((number) => chartOf(`Store ${number}`));
  assert.deepEqual(plain.ranges, ['parts 1 to 30', 'parts 31 to 60']);
  for (const call of plain.calls) assert.deepEqual(itemsOf(call), ['block', 'caption', 'categories', 'format', 'type', 'values']);
  assert.deepEqual([plain.after.blocks?.length, plain.after.slots.b7Chart, plain.after.slots.b60Chart], [60, chartOf('Store 7'), chartOf('Store 60')]);

  // Each of the two calls is told that the other parts are written apart. In a change the parts are told by what they
  // hold now, so nothing is said of the order of a brief.
  for (const call of plain.calls) assert.ok(promptOf(call).includes('give "" or [].\nThe other parts of the slide are written apart in the same way.\nCurrent content: {"r01":[{"block":'), promptOf(call).slice(-200));
});

test('a change plans a free slide with the optional parts it has, so that a caption or a heading is written again with its part', async () => {
  const photo = (name: string) => ({ asset: `gen-${name}`, prompt: `A dog called ${name}`, alt: name });
  const answer = (...slides: unknown[]) => ({ summary: 'Done.', language: 'en', title: 'Dogs', text_size: 'normal', fields, slides });
  const deckOf = (slide: SlideSpec): DeckSpec => ({ version: 1, title: 'Dogs', language: 'en', direction: 'ltr', brief: 'A deck about dogs', fields, slides: [{ ...slide, base: 'content' }, { pattern: 'closing', slots: {} }] });
  const s1 = (request: JsonRequest | undefined) => schemaOf(request).properties.s1 as Schema;
  const context: CatalogContext = { layout: 'strict', fields, icons: [], assets: [], imageDensity: 'medium', figures: false, textScale: 1 };
  const names = (blocks: FreeBlockKind[], withParts?: string[]) => {
    const made = freeWritePattern(template, { direction: 'ltr', textScale: 1 }, { blocks }, context, undefined, withParts);
    assert.ok('pattern' in made);
    return made.pattern.slots.map((planned) => planned.slot.name);
  };

  // One photo and a note. A caption is planned for the photos of a slide that has several, and a heading over the
  // texts of a slide of several parts. A slide may have either all the same: a person wrote it, or the slide once had more parts.
  assert.deepEqual(names(['photo', 'note']), ['title', 'b1Photo', 'b2Text']);
  assert.deepEqual(names(['photo', 'note'], ['b1Caption']), ['title', 'b1Photo', 'b1Caption', 'b2Text']);
  assert.deepEqual(names(['text']), ['title', 'b1Text']);
  assert.deepEqual(names(['text'], ['b1Heading', 'b1Text', 'b7Caption']), ['title', 'b1Heading', 'b1Text'], 'a part the slide does not have is passed over');
  // What a change passes: every slot that holds a value on the slide after it, kept or to be written again, and not the title.
  assert.deepEqual(heldParts({ carried: { title: 'Rex', b1Photo: photo('Rex'), b2Heading: '' }, current: { b1Caption: 'Rex in the park', b2Text: ['One.', 'Two.'] } }), ['b1Photo', 'b1Caption', 'b2Text']);

  const rex = deckOf(freeSlide({ title: 'Rex', blocks: [{ photo: photo('Rex'), caption: 'Rex in the park' }, { note: 'Rex is looking for a home.' }] }));
  const before = rex.slides[0] as SlideSpec;
  // The photo is replaced: the picture is asked for, and its caption with it, shown as it is. A new picture under the old caption would be a slip nobody asked for.
  const cat = scripted(answer(stay(1, 'free', { change: 'A cat called Tom in place of Rex', images: true, parts: [part(1, 'photo', true), part(2, 'note')] }), stay(2, 'closing')), {
    s1: { title: 'Tom', b1Photo: { prompt: 'A cat on a sofa', alt: 'A cat' }, b1Caption: 'Tom on the sofa' },
  });
  const swapped = await planEdit(template, rex, 'A cat called Tom in place of Rex on slide 1', { llm: cat.llm });
  assert.deepEqual([Object.keys(s1(cat.requests[1]).properties), s1(cat.requests[1]).required], [['title', 'b1Photo', 'b1Caption'], ['b1Photo']]);
  assert.ok(promptOf(cat.requests[1]).includes('\nCurrent content: {"title":"Rex","b1Photo":{"prompt":"A dog called Rex","alt":"Rex"},"b1Caption":"Rex in the park"}\n'), promptOf(cat.requests[1]).slice(-400));
  const after = swapped.deck.slides[0] as SlideSpec;
  assert.deepEqual([(after.slots.b1Photo as { prompt?: string }).prompt, after.slots.b1Caption, after.slots.b2Text, after.slots.title], ['A cat on a sofa', 'Tom on the sofa', before.slots.b2Text, 'Tom']);
  assert.deepEqual(swapped.warnings, []);

  // Where the change is not about pictures the photo stays, and its caption alone is written again.
  const named = scripted(answer(stay(1, 'free', { change: 'The caption becomes "Rex at home"', parts: [part(1, 'photo', true), part(2, 'note')] }), stay(2, 'closing')), { s1: { b1Caption: 'Rex at home' } });
  const renamed = await planEdit(template, rex, 'The caption of the photo on slide 1 becomes "Rex at home"', { llm: named.llm });
  assert.deepEqual(Object.keys(s1(named.requests[1]).properties), ['title', 'b1Caption']);
  assert.ok(promptOf(named.requests[1]).includes('\nOf the parts that are written, this stays as it is: 1 photo of "A dog called Rex"\nCurrent content: {"title":"Rex","b1Caption":"Rex in the park"}\n'), promptOf(named.requests[1]).slice(-400));
  assert.deepEqual(renamed.deck.slides[0]?.slots, { ...before.slots, b1Caption: 'Rex at home' });
  // The caption of a part that stays is not asked for. It stays, and the slide is set out with it as before.
  const note = scripted(answer(stay(1, 'free', { change: 'Say that Rex has a home now', parts: [part(1, 'photo'), part(2, 'note', true)] }), stay(2, 'closing')), { s1: { b2Text: 'Rex has a home now.' } });
  const noted = await planEdit(template, rex, 'Say in the note on slide 1 that Rex has a home now', { llm: note.llm });
  assert.deepEqual(Object.keys(s1(note.requests[1]).properties), ['title', 'b2Text']);
  assert.deepEqual(noted.deck.slides[0]?.slots, { ...before.slots, b2Text: 'Rex has a home now.' });
  assert.deepEqual(validateDeck(template, noted.deck), []);

  // One text under a heading, in another language: the heading is written again with the text. Planned by the rules
  // alone, the slide would be offered no heading, and its heading would stay in the language it was written in.
  const hours = deckOf(freeSlide({ title: 'Hours', blocks: [{ text: 'Bikes unlock from five in the morning at every dock in town.', heading: 'From five' }] }));
  const french = scripted({ ...answer(stay(1, 'free'), stay(2, 'closing')), language: 'fr', fields: { closing: 'Merci' } }, { s1: { title: 'Horaires', b1Heading: 'Dès cinq heures', b1Text: ['Les vélos se déverrouillent dès cinq heures du matin.'] }, s2: {} });
  const translated = await planEdit(template, hours, 'In French', { llm: french.llm });
  assert.deepEqual(Object.keys(s1(french.requests[1]).properties), ['title', 'b1Heading', 'b1Text']);
  assert.ok(promptOf(french.requests[1]).includes('\nCurrent content: {"title":"Hours","b1Heading":"From five","b1Text":"Bikes unlock from five in the morning at every dock in town."}\n'), promptOf(french.requests[1]).slice(-400));
  assert.deepEqual(translated.deck.slides[0]?.slots, { title: 'Horaires', b1Heading: 'Dès cinq heures', b1Text: ['Les vélos se déverrouillent dès cinq heures du matin.'] });
});

test('in a change, a chart that is larger than its place reads well at is asked for as it is, and is not sent back for its size', async () => {
  // Six charts of three series each. At this size a new chart is told two series; these were written when the slide had
  // fewer charts, or by code. A change that is not about their numbers must not cost a chart a series.
  const three = (name: string): ChartValue => ({ type: 'column', categories: quarters, series: ['2023', '2024', '2025'].map((year) => ({ name: `${name} ${year}`, values: [10, 14, 12, 18] })) });
  const slide: SlideSpec = { ...freeSlide({ title: 'Sales by store', blocks: [...storeNames.map((name) => ({ chart: three(name), caption: name })), { note: NOTE }], columns: 3 }), base: 'list-with-photo' };
  const deck = storesDeck(slide);
  const large = (warnings: Issue[]) => warnings.filter((issue) => /^the chart has \d+ series; the slot fits \d+$/.test(issue.message));
  // The deck says of itself what it says after the change: each chart holds more than reads well in its place.
  assert.deepEqual(large(validateDeck(demo, deck, adaptive)).map((issue) => [issue.slot, issue.message]), storeNames.map((_, index) => [`b${index + 1}Chart`, 'the chart has 3 series; the slot fits 2']));
  const s2 = (request: JsonRequest | undefined) => schemaOf(request).properties.s2 as Schema;
  const translate = { ...changeOf(stay(1, 'cover'), stay(2, 'free', { side_by_side: 3 }), stay(3, 'closing')), language: 'fr' };
  const french = { ...slide.slots, title: 'Ventes par magasin', b7Text: 'Chaque magasin a vendu le plus au dernier trimestre.' };

  // Every chart is offered with what it holds, and its schema allows as many series as it has: a schema that allowed two would take the third away.
  const kept = scripted(translate, { s1: { title: 'Magasins' }, s2: french, s3: { title: 'Merci' } });
  const result = await planEdit(demo, deck, 'In French', { llm: kept.llm, ...adaptive });
  const chart = s2(kept.requests[1]).properties.b1Chart as Schema;
  assert.match(chart.description ?? '', /^A small chart\. When the brief gives no numbers for it, leave it empty\. at most \d+ categories, at most 3 series\. A chart of numbers/);
  assert.equal(chart.properties.series?.maxItems, 3);
  assert.equal(kept.requests.length, 2, 'the plan of the change and one call: no chart is asked for again');
  for (const number of range(1, 6)) assert.deepEqual(result.deck.slides[1]?.slots[`b${number}Chart`], slide.slots[`b${number}Chart`], `chart ${number}`);
  // The deck is told as before, in a warning on each chart.
  assert.equal(large(result.warnings).length, 6);

  // A chart that comes back larger than it was is sent back like any chart over its limits.
  const grown = { ...french, b1Chart: { ...three('North'), series: [...three('North').series, { name: 'North 2026', values: [10, 14, 12, 18] }] } };
  const more = scripted(translate, { s1: { title: 'Magasins' }, s2: grown, s3: { title: 'Merci' } }, { s1: { title: 'Magasins' }, s2: french, s3: { title: 'Merci' } });
  await planEdit(demo, deck, 'In French', { llm: more.llm, ...adaptive });
  assert.equal(more.requests.length, 3);
  assert.ok(promptOf(more.requests[2]).includes('\nIt has these problems:\n- s2.b1Chart: the chart has 4 series; the slot fits 2\n\nReturn the full content again'), promptOf(more.requests[2]).slice(-300));

  // A chart that is new on the same slide is told what its place reads well at, and so is the chart of a new slide.
  const parts = [...sixAndNote({ 1: part(1, 'chart', true) }).slice(0, 6), part(0, 'chart', true), part(7, 'note')];
  const added = scripted(changeOf(stay(1, 'cover'), stay(2, 'free', { change: 'New numbers for North, and a chart for Airport', parts, side_by_side: 4 }), stay(3, 'closing')), { s2: {} });
  await planEdit(demo, deck, 'New numbers for North, and a chart for Airport', { llm: added.llm, ...adaptive, maxRepairs: 0 });
  assert.deepEqual([s2(added.requests[1]).properties.b1Chart?.properties.series?.maxItems, s2(added.requests[1]).properties.b7Chart?.properties.series?.maxItems], [3, 2]);

  // The limits are those of the slot wherever the part holds no more: nothing is raised for a chart of one series.
  const context: CatalogContext = { layout: 'adaptive', fields: stores.fields as Record<string, string>, icons: [], assets: [], imageDensity: 'medium', figures: true, textScale: 1 };
  const made = freeWritePattern(demo, { direction: 'ltr', textScale: 1 }, { blocks: slide.blocks as FreeBlockKind[], arrange: { columns: 3 } }, context);
  assert.ok('pattern' in made);
  const limitsOf = (pattern: typeof made.pattern, name: string) => {
    const slot = pattern.slots.find((planned) => planned.slot.name === name)?.slot;
    return slot?.kind === 'chart' ? [slot.maxCategories, slot.maxSeries] : slot?.kind === 'table' ? [slot.maxColumns, slot.maxRows] : [];
  };
  const [categories] = limitsOf(made.pattern, 'b1Chart');
  const wide: ChartValue = { type: 'line', categories: range(1, 20).map(String), series: [{ name: 'Sales', values: range(1, 20) }] };
  const raised = roomAsHeld(made.pattern, { b1Chart: three('North'), b2Chart: chartOf('South'), b3Chart: wide, b7Text: NOTE });
  assert.deepEqual([limitsOf(raised, 'b1Chart'), limitsOf(raised, 'b2Chart'), limitsOf(raised, 'b3Chart'), limitsOf(raised, 'b4Chart')], [[categories, 3], [categories, 2], [20, 2], [categories, 2]]);
  assert.deepEqual(limitsOf(made.pattern, 'b1Chart'), [categories, 2], 'the pattern it was given is not changed');
});

test('a plan that names the parts of a free slide as they are, and has words for the slide, gets its title written and no part', async () => {
  const deck = storesDeck();
  const before = deck.slides[1] as SlideSpec;
  // Every part named, none to be written again, and a new title: the parts are kept by code, as the plan says of each.
  const { llm, requests } = scripted(changeOf(stay(1, 'cover'), stay(2, 'free', { change: 'The title becomes "Sales 2025"', parts: sixAndNote(), side_by_side: 3 }), stay(3, 'closing')), { s2: { title: 'Sales 2025' } });
  const result = await planEdit(demo, deck, 'Call slide 2 "Sales 2025"', { llm, ...adaptive });
  assert.equal(requests.length, 2);
  assert.deepEqual(Object.keys((schemaOf(requests[1]).properties.s2 as Schema).properties), ['title']);
  assert.ok(!JSON.stringify(requests[1]).includes('[10,14,12,18]'), 'no number of a chart is shown to the writer');
  assert.deepEqual(result.deck.slides[1], { ...before, slots: { ...before.slots, title: 'Sales 2025' } });
  assert.deepEqual([result.rewritten, result.warnings.filter((issue) => issue.slot !== undefined)], [[2], []]);
  const unchanged = (warnings: Issue[]) => warnings.filter((issue) => /nothing on it/.test(issue.message));
  assert.deepEqual(unchanged(result.warnings), [], 'the title changed');

  // The plan has words that are about a part, and marks no part to be written again. The title is all a writer can be
  // offered, and it comes back as it was, or not at all: nothing on the slide changed. The deck is told so in a note
  // with no slide number, which starts no rewrite, and no call is added to find out more.
  const NOTHING: Issue = { severity: 'warning', message: 'slide 2: nothing on it was changed: the plan of the change named no part to write again' };
  for (const answer of [{ s2: { title: 'Sales by store' } }, { s2: {} }]) {
    const idle = scripted(changeOf(stay(1, 'cover'), stay(2, 'free', { change: 'Call the chart of the west "West (new)"', parts: sixAndNote(), side_by_side: 3 }), stay(3, 'closing')), answer);
    const same = await planEdit(demo, deck, 'Call the chart of the west "West (new)" on slide 2', { llm: idle.llm, ...adaptive });
    assert.equal(idle.requests.length, 2, 'the plan of the change and the one small call, as before');
    assert.deepEqual(same.deck.slides[1], before);
    assert.deepEqual(withoutCodes(unchanged(same.warnings)), [NOTHING]);
    assert.ok(!('slide' in (unchanged(same.warnings)[0] as Issue)));
    assert.deepEqual(same.rewritten, [], 'the slide is not said to be rewritten');
  }
  // The same title on a slide the plan sets out anew is a change all the same, and nothing is told.
  const wider = scripted(changeOf(stay(1, 'cover'), stay(2, 'free', { change: 'Two side by side', parts: sixAndNote(), side_by_side: 2 }), stay(3, 'closing')), { s2: { title: 'Sales by store' } });
  const set = await planEdit(demo, deck, 'Two side by side on slide 2', { llm: wider.llm, ...adaptive });
  assert.deepEqual([set.deck.slides[1]?.arrange, unchanged(set.warnings)], [{ columns: 2 }, []]);

  // The plan has the same words and names no part: every part is then offered with what it holds. A writer that returns
  // what changes and nothing else has done as it was asked. No part was named to be written, so none is asked for again.
  const loose = scripted(changeOf(stay(1, 'cover'), stay(2, 'free', { change: 'The title becomes "Sales 2025"', side_by_side: 3 }), stay(3, 'closing')), { s2: { title: 'Sales 2025' } });
  const offeredAll = await planEdit(demo, deck, 'Call slide 2 "Sales 2025"', { llm: loose.llm, ...adaptive });
  assert.equal(loose.requests.length, 2);
  assert.equal(Object.keys((schemaOf(loose.requests[1]).properties.s2 as Schema).properties).length, 14, 'the title, six captions, six charts and the note');
  assert.deepEqual(offeredAll.deck.slides[1], { ...before, slots: { ...before.slots, title: 'Sales 2025' } });
  assert.deepEqual(offeredAll.warnings.filter((issue) => issue.slide === undefined && /part/.test(issue.message)), []);

  // Sixty charts: one small call for the title, where writing every chart again would take two long ones.
  const sixty: SlideSpec = { ...freeSlide({ title: 'Every store', blocks: range(1, 60).map((number) => ({ chart: chartOf(`Store ${number}`), caption: `Store ${number}` })) }), base: 'list-with-photo' };
  const named = changeOf(stay(1, 'cover'), stay(2, 'free', { change: 'The title becomes "Sales 2025"', parts: range(1, 60).map((number) => part(number, 'chart')) }), stay(3, 'closing'));
  const many = scripted(named, { title: 'Sales 2025' });
  const renamed = await planEdit(demo, storesDeck(sixty), 'Call slide 2 "Sales 2025"', { llm: many.llm, ...adaptive });
  assert.equal(many.requests.length, 2);
  const call = many.requests[1] as JsonRequest;
  assert.deepEqual([Object.keys(schemaOf(call).properties), schemaOf(call).required], [['title'], ['title']]);
  assert.ok(promptOf(call).includes('\nWrite the title of slide 2 (free: 60 charts): The title becomes "Sales 2025"\nCurrent content: {"title":"Every store"}\nWhat stays as it is and is not to be written: 1 chart "Store 1" (column, 4 categories); '), promptOf(call).slice(-900, -500));
  assert.ok((call.maxOutputTokens ?? Infinity) < 1000, String(call.maxOutputTokens));
  assert.equal(renamed.deck.slides[1]?.slots.title, 'Sales 2025');
  assert.equal(JSON.stringify(renamed.deck.slides[1]?.slots.b60Chart), JSON.stringify(sixty.slots.b60Chart));
  assert.deepEqual(unchanged(renamed.warnings), []);
  // The title of sixty charts comes back as it was: told as for a slide of seven parts, after the one call.
  const still60 = scripted(named, { title: 'Every store' });
  const idle60 = await planEdit(demo, storesDeck(sixty), 'Call store 7 "Harbour"', { llm: still60.llm, ...adaptive });
  assert.deepEqual([still60.requests.length, withoutCodes(unchanged(idle60.warnings)), idle60.rewritten], [2, [NOTHING], []]);
  assert.equal(JSON.stringify(idle60.deck.slides[1]?.slots), JSON.stringify(sixty.slots));

  // The title is asked of every such answer. A slide that has one is told to repeat it unless the request is about it,
  // and an empty answer leaves it as it is: a call that adds two charts must not take the title off by a slip.
  assert.match(schemaOf(call).properties.title?.description ?? '', /^The title of the slide\. at most \d+ characters\. Repeat the current title unless the request is about it$/);
  const adding = changeOf(stay(1, 'cover'), stay(2, 'free', { change: 'Add stores 61 and 62', parts: [...range(1, 60).map((number) => part(number, 'chart')), part(0, 'chart', true), part(0, 'chart', true)] }), stay(3, 'closing'));
  const slip = scripted(adding, { r01: [sparkItem(61), sparkItem(62)], title: '' });
  const grown = await planEdit(demo, storesDeck(sixty), 'Add stores 61 and 62 with the numbers 10, 14, 12 and 18', { llm: slip.llm, ...adaptive });
  assert.deepEqual([grown.deck.slides[1]?.blocks?.length, grown.deck.slides[1]?.slots.title], [62, 'Every store']);
  // A slide without a title is asked as a new one is: "" is how it stays without.
  const { title: _none, ...untitledSlots } = sixty.slots;
  const bare = scripted(adding, { r01: [sparkItem(61), sparkItem(62)], title: '' });
  const still = await planEdit(demo, storesDeck({ ...sixty, slots: untitledSlots }), 'Add stores 61 and 62 with the numbers 10, 14, 12 and 18', { llm: bare.llm, ...adaptive });
  assert.match(schemaOf(bare.requests[1]).properties.title?.description ?? '', /\. "" for a slide without one$/);
  assert.ok(!('title' in (still.deck.slides[1]?.slots ?? {})));
});

test('a change that offers the title of a free slide that has one asks for it to be repeated unless the request is about it', async () => {
  const deck = storesDeck();
  const before = deck.slides[1] as SlideSpec;
  const worded = 'Online grew fastest of all the stores this year.';
  const repeated = /(?:^|\. )at most \d+ characters\. Repeat the current title unless the request is about it$/;
  const titleOf = (request: JsonRequest | undefined, slide = 's2') => (schemaOf(request).properties[slide] as Schema).properties.title?.description ?? '';
  const noteAgain = stay(2, 'free', { change: 'Say in the note that Online grew fastest', parts: sixAndNote({ 7: part(7, 'note', true) }), side_by_side: 3 });

  // Written with the deck, beside a part: the title is offered, may be left out, and is to be said again as it is.
  const withPart = scripted(changeOf(stay(1, 'cover', { change: 'A shorter title' }), noteAgain, stay(3, 'closing')), { s1: { title: 'Shops' }, s2: { b7Text: worded } });
  const kept = await planEdit(demo, deck, 'Say in the note on slide 2 that Online grew fastest, and shorten the title of slide 1', { llm: withPart.llm, ...adaptive });
  assert.match(titleOf(withPart.requests[1]), repeated);
  assert.deepEqual((schemaOf(withPart.requests[1]).properties.s2 as Schema).required, ['b7Text'], 'left out, the title stays');
  assert.deepEqual([kept.deck.slides[1]?.slots.title, kept.deck.slides[1]?.slots.b7Text], ['Sales by store', worded]);
  // A slide of a pattern that is written in the same call is asked for its title as it always was.
  assert.doesNotMatch(titleOf(withPart.requests[1], 's1'), /Repeat/);

  // The title alone, in the one small call of a plan that has words for the slide and names no part to write.
  const alone = scripted(changeOf(stay(1, 'cover'), stay(2, 'free', { change: 'The title becomes "Sales 2025"', parts: sixAndNote(), side_by_side: 3 }), stay(3, 'closing')), { s2: { title: 'Sales 2025' } });
  await planEdit(demo, deck, 'Call slide 2 "Sales 2025"', { llm: alone.llm, ...adaptive });
  assert.deepEqual(Object.keys((schemaOf(alone.requests[1]).properties.s2 as Schema).properties), ['title']);
  assert.match(titleOf(alone.requests[1]), repeated);

  // On a slide written with the deck an explicit "" takes the title off: that is how a request for no title is met.
  const none = scripted(changeOf(stay(1, 'cover'), stay(2, 'free', { change: 'No title, and say in the note that Online grew fastest', parts: sixAndNote({ 7: part(7, 'note', true) }), side_by_side: 3 }), stay(3, 'closing')), { s2: { title: '', b7Text: worded } });
  const untitled = await planEdit(demo, deck, 'No title on slide 2, and say in the note that Online grew fastest', { llm: none.llm, ...adaptive });
  assert.ok(!('title' in (untitled.deck.slides[1]?.slots ?? {})));
  assert.equal(untitled.deck.slides[1]?.slots.b7Text, worded);

  // A slide that has no title has none to repeat, and neither has a slide that is new or was a slide of a pattern.
  const { title: _none, ...bare } = before.slots;
  const without = scripted(changeOf(stay(1, 'cover'), noteAgain, stay(3, 'closing')), { s2: { b7Text: worded } });
  await planEdit(demo, storesDeck({ ...before, slots: bare }), 'Say in the note on slide 2 that Online grew fastest', { llm: without.llm, ...adaptive });
  assert.doesNotMatch(titleOf(without.requests[1]), /Repeat/);
  const cards = { title: 'Why stores matter', b1Heading: 'Near', b1Text: 'Close to where people live.', b2Heading: 'Seen', b2Text: 'People can try things.' };
  const added = scripted(changeOf(stay(1, 'cover'), stay(2, 'free', { side_by_side: 3 }), stay(0, 'free', { change: 'Two cards on why stores matter', parts: [part(0, 'card', true), part(0, 'card', true)] }), stay(3, 'closing')), { s3: cards });
  await planEdit(demo, deck, 'Add a slide with two cards on why stores matter', { llm: added.llm, ...adaptive });
  assert.doesNotMatch(titleOf(added.requests[1], 's3'), /Repeat/);
});

test('in a change, an array may come back without the parts that have content and were not named to be written again', async () => {
  const twelve: SlideSpec = { ...freeSlide({ title: 'Every store', blocks: range(1, 12).map((number) => ({ chart: chartOf(`Store ${number}`), caption: `Store ${number}` })) }), base: 'list-with-photo' };
  const deck = storesDeck(twelve);
  const harbour = chartItem(3, { caption: 'Harbour', series: [{ name: 'Store 3', values: [10, 14, 12, 18] }] });
  const counted = (warnings: Issue[]) => warnings.filter((issue) => issue.slide === undefined && /part/.test(issue.message));

  // The plan has words for the slide and names no part: all twelve charts are offered with what they hold. The writer
  // returns the one that changes. The eleven others keep what they had, and no repair round is spent on their number.
  const loose = scripted(changeOf(stay(1, 'cover'), stay(2, 'free', { change: 'Call store 3 "Harbour"' }), stay(3, 'closing')), { r01: [harbour], title: 'Every store' });
  const renamed = await planEdit(demo, deck, 'Call store 3 "Harbour" on slide 2', { llm: loose.llm, ...adaptive });
  assert.equal(loose.requests.length, 2, 'the plan of the change and one call: nothing is asked again');
  assert.match(schemaOf(loose.requests[1]).properties.r01?.description ?? '', /^Exactly 12 charts: parts 1 to 12\./);
  const after = renamed.deck.slides[1] as SlideSpec;
  assert.deepEqual(after.slots, { ...twelve.slots, b3Caption: 'Harbour' });
  assert.deepEqual([after.blocks?.length, counted(renamed.warnings)], [12, []]);

  // An item that does not say which part it is could be any of the twelve. It is given to none of them: the first chart
  // must not be written over by its place alone. The array is asked for whole, and the answer that then names its part
  // is taken, with nothing left of the item that did not.
  const { block: _which, ...blind } = harbour;
  const unsure = scripted(changeOf(stay(1, 'cover'), stay(2, 'free', { change: 'Call store 3 "Harbour"' }), stay(3, 'closing')), { r01: [blind], title: 'Every store' }, { r01: [harbour], title: 'Every store' });
  const asked = await planEdit(demo, deck, 'Call store 3 "Harbour" on slide 2', { llm: unsure.llm, ...adaptive });
  assert.equal(unsure.requests.length, 3);
  assert.ok(promptOf(unsure.requests[2]).includes('\nIt has these problems:\n- s2.r01: 1 item; it needs exactly 12, parts 1 to 12\n\nReturn the full content again'), promptOf(unsure.requests[2]).slice(-300));
  assert.deepEqual((asked.deck.slides[1] as SlideSpec).slots, { ...twelve.slots, b3Caption: 'Harbour' });
  // Still without a number after the last round, the item is dropped and the slide is as it was.
  const never = scripted(changeOf(stay(1, 'cover'), stay(2, 'free', { change: 'Call store 3 "Harbour"' }), stay(3, 'closing')), { r01: [blind], title: 'Every store' });
  const untouched = await planEdit(demo, deck, 'Call store 3 "Harbour" on slide 2', { llm: never.llm, ...adaptive, maxRepairs: 0 });
  assert.deepEqual((untouched.deck.slides[1] as SlideSpec).slots, twelve.slots);
  // Twelve items for twelve parts are taken in their order, numbered or not.
  const echo = scripted(changeOf(stay(1, 'cover'), stay(2, 'free', { change: 'Call store 3 "Harbour"' }), stay(3, 'closing')), {
    r01: range(1, 12).map((number) => chartItem(undefined, { caption: number === 3 ? 'Harbour' : `Store ${number}`, series: [{ name: `Store ${number}`, values: [10, 14, 12, 18] }] })),
    title: 'Every store',
  });
  const inOrder = await planEdit(demo, deck, 'Call store 3 "Harbour" on slide 2', { llm: echo.llm, ...adaptive });
  assert.deepEqual([echo.requests.length, (inOrder.deck.slides[1] as SlideSpec).slots], [2, { ...twelve.slots, b3Caption: 'Harbour' }]);

  // Parts the plan named outright to be written again are due, and new parts have nothing to keep: for those the
  // number of items is a problem, and is sent back once more.
  const parts = [...range(1, 12).map((number) => part(number, 'chart', number <= 2)), part(0, 'chart', true)];
  const one = { r01: [chartItem(1, { caption: 'Store one', series: [{ name: 'Store 1', values: [10, 14, 12, 18] }] })], title: 'Every store' };
  const due = scripted(changeOf(stay(1, 'cover'), stay(2, 'free', { change: 'New names for stores 1 and 2, and a store 13', parts }), stay(3, 'closing')), one, one);
  const some = await planEdit(demo, deck, 'New names for stores 1 and 2, and a store 13', { llm: due.llm, ...adaptive, maxRepairs: 1 });
  assert.equal(due.requests.length, 3);
  assert.ok(promptOf(due.requests[2]).includes('\nIt has these problems:\n- s2.r01: 1 item; it needs exactly 3, parts 1, 2 and 13\n\nReturn the full content again'), promptOf(due.requests[2]).slice(-300));
  assert.deepEqual(counted(some.warnings).map((issue) => issue.message).sort(), ['slide 2: 1 part had no content and was left out', 'slide 2: nothing was written for 1 part, and it stays as it was']);
});

test('a short array of a change is a problem for the parts that are new or due, and names them', () => {
  // Four cards as one call writes them. Cards 1 to 3 have content; card 2 was named to be written again; card 4 is new.
  const blocks: FreeBlockKind[] = ['card', 'card', 'card', 'card'];
  const context: CatalogContext = { layout: 'adaptive', fields: stores.fields as Record<string, string>, icons: [], assets: [], imageDensity: 'medium', figures: true, textScale: 1 };
  const made = freeWritePattern(demo, { direction: 'ltr', textScale: 1 }, { blocks }, context);
  assert.ok('pattern' in made);
  const current = Object.fromEntries([1, 2, 3].flatMap((number) => [[`b${number}Heading`, `Card ${number}`], [`b${number}Text`, 'What it had.']]));
  const job: FreeJob = { number: 2, pattern: made.pattern, purpose: 'A change', free: { blocks, base: made.base }, current, due: new Set(['b2Heading', 'b2Text']) };
  const [call] = partRanges(job, context);
  assert.ok(call);
  assert.deepEqual(call.runs, [{ name: 'r01', kind: 'card', parts: [1, 2, 3, 4] }]);
  const card = (block?: number) => ({ ...(block === undefined ? {} : { block }), heading: 'New', text: 'What it has now.' });
  const counts = (items: unknown[], of: FreeJob = job) => readRange(of, call, { r01: items, title: 'T' }, context).counts;

  // Cards 1 and 3 may be left out: they keep what they had.
  assert.deepEqual(counts([card(2), card(4)]), []);
  assert.deepEqual(counts([card(4), card(2), card(1)]), []);
  // What is due or new and did not come is named, and nothing else is asked for.
  assert.deepEqual(counts([card(4)]), ['s2.r01: 1 item, and none for part 2; it needs that one too']);
  assert.deepEqual(counts([card(1)]), ['s2.r01: 1 item, and none for parts 2 and 4; it needs those too']);
  assert.deepEqual(counts([]), ['s2.r01: 0 items, and none for parts 2 and 4; it needs those too']);
  // An item without a number of its own, with one that is taken or with one the array does not have takes the next part
  // that has to be written: the new card here. It never takes a part that may stay.
  for (const second of [card(), card(2), card(9)]) {
    const read = readRange(job, call, { r01: [card(2), { ...second, heading: 'Second' }], title: 'T' }, context);
    assert.deepEqual([read.counts, read.values.b4Heading, read.values.b1Heading, read.values.b3Heading], [[], 'Second', undefined, undefined]);
  }
  // With no such part left for it, it could be any of the others: it is given to none, and all four are asked for.
  const lost = readRange(job, call, { r01: [card(2), card(4), card()], title: 'T' }, context);
  assert.deepEqual([lost.counts, lost.absent], [['s2.r01: 3 items; it needs exactly 4, parts 1 to 4'], [1, 3]]);
  // A slide that is new has nothing to keep: every part is needed, as it always was, and items are taken in their order.
  const { current: _has, due: _named, ...fresh } = job;
  assert.deepEqual(counts([card(2), card(4)], fresh), ['s2.r01: 2 items; it needs exactly 4, parts 1 to 4']);
  assert.equal(readRange(fresh, call, { r01: [card(2), { heading: 'Blind', text: 'x' }], title: 'T' }, context).values.b1Heading, 'Blind');
  // An item for every part is no problem, whatever the numbers say, on a slide that is new or one that is changed.
  assert.deepEqual(counts([card(), card(), card(), card()], fresh), []);
  assert.deepEqual(readRange(job, call, { r01: [card(), card(), card(), card()], title: 'T' }, context).absent, []);
});

test('in a change, a chart of a free slide is asked for as in a new deck: required, and empty where there is nothing for it', async () => {
  const deck = storesDeck();
  const before = deck.slides[1] as SlideSpec;
  const counted = (warnings: Issue[]) => warnings.filter((issue) => issue.slide === undefined && /part/.test(issue.message));
  const s2 = (request: JsonRequest | undefined) => schemaOf(request).properties.s2 as Schema;

  // Charts that are merely offered with what they hold: the plan has words for the slide and names no part. Each is
  // asked for, and one that is given empty keeps what it had, with no repair round and nothing to tell.
  const offeredAll = scripted(changeOf(stay(1, 'cover'), stay(2, 'free', { change: 'The title becomes "Sales 2025"', side_by_side: 3 }), stay(3, 'closing')), {
    s2: { title: 'Sales 2025', ...Object.fromEntries(range(1, 6).map((number) => [`b${number}Chart`, emptyChart])), b7Text: NOTE },
  });
  const renamed = await planEdit(demo, deck, 'Call slide 2 "Sales 2025"', { llm: offeredAll.llm, ...adaptive });
  assert.deepEqual(s2(offeredAll.requests[1]).required, [...range(1, 6).map((number) => `b${number}Chart`), 'b7Text']);
  assert.doesNotMatch(JSON.stringify(s2(offeredAll.requests[1]).properties.b1Chart), /minItems/);
  assert.equal(offeredAll.requests.length, 2);
  assert.deepEqual(renamed.deck.slides[1], { ...before, slots: { ...before.slots, title: 'Sales 2025' } });
  assert.deepEqual(counted(renamed.warnings), []);

  // A chart the plan names to be written again is asked for in the same way, and nothing in its schema differs from
  // that of a chart that is new. Given empty it keeps what it had too, and since it was named it is asked for once more.
  const again = changeOf(stay(1, 'cover'), stay(2, 'free', { change: 'North sold 20 in the last quarter, and add a chart for Airport', parts: [...sixAndNote({ 1: part(1, 'chart', true) }).slice(0, 6), part(0, 'chart', true), part(7, 'note')], side_by_side: 3 }), stay(3, 'closing'));
  const empty = { s2: { title: 'Sales by store', b1Caption: 'North', b1Chart: emptyChart, b7Caption: 'Airport', b7Chart: emptyChart } };
  const named = scripted(again, empty, empty);
  const result = await planEdit(demo, deck, 'North sold 20 in the last quarter, and add a chart for Airport', { llm: named.llm, ...adaptive, maxRepairs: 1 });
  assert.deepEqual(s2(named.requests[1]).required, ['b1Chart', 'b7Chart']);
  assert.deepEqual({ ...s2(named.requests[1]).properties.b1Chart, description: '' }, { ...s2(named.requests[1]).properties.b7Chart, description: '' });
  assert.equal(named.requests.length, 3);
  // The chart that was due is the one problem: the new one had no numbers, and asking again brings none.
  assert.ok(promptOf(named.requests[2]).includes('\nIt has these problems:\n- s2.b1Chart: nothing was written for it, and it is to be written again\n\nReturn the full content again'), promptOf(named.requests[2]).slice(-300));
  const after = result.deck.slides[1] as SlideSpec;
  assert.deepEqual([after.blocks, after.slots], [before.blocks, before.slots]);
  assert.deepEqual(counted(result.warnings).map((issue) => issue.message).sort(), ['slide 2: 1 part had no content and was left out', 'slide 2: nothing was written for 1 part, and it stays as it was']);
});

test('a part of another kind takes the place of a part, and text that no longer fits a larger size is written again by itself', async () => {
  const deck = storesDeck();
  const before = deck.slides[1] as SlideSpec;
  // The note becomes a quotation: written new, with what the note said as material.
  const quote = scripted(changeOf(stay(1, 'cover'), stay(2, 'free', { change: 'Make the note a quotation by the head of sales', parts: sixAndNote({ 7: part(7, 'quote', true) }), side_by_side: 3 }), stay(3, 'closing')), {
    s2: { title: 'Sales by store', b7Quote: 'Our best quarter ever.', b7Source: 'Head of sales' },
  });
  const quoted = await planEdit(demo, deck, 'Make the note on slide 2 a quotation by the head of sales', { llm: quote.llm, ...adaptive });
  assert.deepEqual(Object.keys((schemaOf(quote.requests[1]).properties.s2 as Schema).properties), ['title', 'b7Quote', 'b7Source']);
  assert.ok(promptOf(quote.requests[1]).includes(`\nThe new parts with these numbers take the place of parts that held: ${JSON.stringify({ 7: { text: NOTE } })}\nCurrent content: {"title":"Sales by store"}\n`));
  assert.deepEqual(quoted.deck.slides[1]?.blocks, [...Array<string>(6).fill('chart'), 'quote']);
  assert.deepEqual([quoted.deck.slides[1]?.slots.b7Quote, quoted.deck.slides[1]?.slots.b7Source, quoted.deck.slides[1]?.slots.b7Text], ['Our best quarter ever.', 'Head of sales', undefined]);
  assert.equal(JSON.stringify(quoted.deck.slides[1]?.slots.b6Chart), JSON.stringify(before.slots.b6Chart));

  // Larger text: the title of this slide is too long for its line, and nothing else is. The title alone is written again.
  const long: SlideSpec = { ...before, slots: { ...before.slots, title: 'Sales by store and what we learnt from them this year' } };
  const larger = scripted({ ...changeOf(stay(1, 'cover'), stay(2, 'free', { side_by_side: 3 }), stay(3, 'closing')), text_size: 'large' }, { s2: { title: 'Sales by store' } });
  const sized = await planEdit(demo, storesDeck(long), 'Larger text', { llm: larger.llm, ...adaptive });
  assert.equal(sized.deck.textScale, 1.25);
  assert.deepEqual(Object.keys(schemaOf(larger.requests[1]).properties), ['s2']);
  assert.deepEqual(Object.keys((schemaOf(larger.requests[1]).properties.s2 as Schema).properties), ['title']);
  assert.match(promptOf(larger.requests[1]), /\ns2 \(free: b1 to b6: 6 charts; b7: a note\)\. Change: Shorten the text to fit the new text size\. Keep the meaning\.\nWhat stays as it is and is not to be written: 1 chart "North" .*; 7 note "Every store[^\n]*\nCurrent content: \{"title":"Sales by store and what we learnt from them this year"\}\n/);
  assert.deepEqual(sized.deck.slides[1]?.slots, before.slots);
  assert.deepEqual(sized.rewritten, [2]);
});

test('how a free slide is set out is changed by code, with nothing written', async () => {
  const deck = storesDeck();
  const before = deck.slides[1] as SlideSpec;
  const arranged = async (more: Record<string, unknown>) => {
    const { llm, requests } = scripted(changeOf(stay(1, 'cover'), stay(2, 'free', more), stay(3, 'closing')));
    const result = await planEdit(demo, deck, 'Set slide 2 out anew', { llm, ...adaptive });
    assert.equal(requests.length, 1);
    assert.deepEqual(result.deck.slides[1]?.slots, before.slots);
    return result;
  };
  // Two side by side, with the first chart on top: the slide's content is not touched, and the slide counts as changed.
  const two = await arranged({ side_by_side: 2, shows_main: 'top' });
  assert.deepEqual(two.deck.slides[1], { ...before, arrange: { columns: 2, lead: 'top' } });
  assert.deepEqual(two.rewritten, [2]);
  // The current values, repeated, change nothing: the plan is told to repeat them unless the request is about the arrangement.
  assert.deepEqual((await arranged({ side_by_side: 3 })).rewritten, []);
  // auto and 0 on a slide the plan leaves alone (no words for it, its parts as they are) state nothing: they are what an
  // answer gives for a slide it has nothing to say of. The slide is copied as it is, three side by side.
  const auto = await arranged({ side_by_side: 0, shows_main: 'auto' });
  assert.deepEqual([auto.deck.slides[1], auto.rewritten], [before, []]);
  assert.equal(JSON.stringify(auto.deck.slides[1]), JSON.stringify(before), 'byte for byte');
  // A value that is stated always applies, each of the two by itself: all parts alike, with the number side by side as
  // it was, whether the answer repeats it or gives 0.
  assert.deepEqual((await arranged({ side_by_side: 3, shows_main: 'none' })).deck.slides[1]?.arrange, { columns: 3, lead: 'none' });
  const alike = await arranged({ side_by_side: 0, shows_main: 'none' });
  assert.deepEqual([alike.deck.slides[1]?.arrange, alike.rewritten], [{ columns: 3, lead: 'none' }, [2]]);
  assert.deepEqual((await arranged({ side_by_side: 2, shows_main: 'auto' })).deck.slides[1]?.arrange, { columns: 2 });
  // More side by side than the slide has parts is as many as it has.
  assert.deepEqual((await arranged({ side_by_side: 40 })).deck.slides[1]?.arrange, { columns: 7 });
  // An answer that leaves the two properties out states nothing, and the slide stays as it is set out.
  const bare = scripted(changeOf(stay(1, 'cover'), { from: 2, pattern: 'free', change: '', images: false, parts: [] }, stay(3, 'closing')));
  const same = await planEdit(demo, deck, 'Nothing to change', { llm: bare.llm, ...adaptive });
  assert.deepEqual([same.deck.slides[1], same.rewritten], [before, []]);
});

test('auto and 0 take an arrangement away only on a free slide the plan changes', async () => {
  // Six charts and a note, three side by side with all parts alike.
  const set = { columns: 3, lead: 'none' as const };
  const deck = storesDeck({ ...storesSlide(), arrange: set });
  const before = deck.slides[1] as SlideSpec;
  const worded = 'Online grew fastest of all the stores this year.';
  const noteAgain = sixAndNote({ 7: part(7, 'note', true) });
  const written = { s2: { title: 'Sales by store', b7Text: worded } };
  const arrangeAfter = async (more: Record<string, unknown>, answer: unknown = written, options: Record<string, unknown> = {}, plan: Record<string, unknown> = {}) => {
    const { llm } = scripted({ ...changeOf(stay(1, 'cover'), stay(2, 'free', more), stay(3, 'closing')), ...plan }, answer);
    return (await planEdit(demo, deck, 'A change to slide 2', { llm, ...adaptive, ...options })).deck.slides[1] as SlideSpec;
  };

  // The plan changes the slide, by words and by a part to write again. What it states of the arrangement is taken at its
  // word: the values repeated keep it, auto and 0 hand it back to the software, each of the two by itself.
  assert.deepEqual((await arrangeAfter({ change: 'A new note', parts: noteAgain, shows_main: 'none', side_by_side: 3 })).arrange, set);
  assert.ok(!('arrange' in (await arrangeAfter({ change: 'A new note', parts: noteAgain }))), 'auto and 0: the slide states no arrangement');
  assert.deepEqual((await arrangeAfter({ change: 'A new note', parts: noteAgain, shows_main: 'none' })).arrange, { lead: 'none' });
  assert.deepEqual((await arrangeAfter({ change: 'A new note', parts: noteAgain, side_by_side: 3 })).arrange, { columns: 3 });
  // Words alone change the slide, and so do parts alone: a part taken off, with nothing to write.
  assert.ok(!('arrange' in (await arrangeAfter({ change: 'A new note' }))));
  assert.ok(!('arrange' in (await arrangeAfter({ parts: sixAndNote().slice(0, 6) }))));
  // A stated value applies on such a slide as on any, and a property that is left out states nothing.
  assert.deepEqual((await arrangeAfter({ change: 'A new note', parts: noteAgain, shows_main: 'top', side_by_side: 2 })).arrange, { columns: 2, lead: 'top' });
  const { llm } = scripted(changeOf(stay(1, 'cover'), { from: 2, pattern: 'free', change: 'A new note', images: false, parts: noteAgain }, stay(3, 'closing')), written);
  assert.deepEqual((await planEdit(demo, deck, 'A new note on slide 2', { llm, ...adaptive })).deck.slides[1]?.arrange, set);

  // The plan leaves the slide alone and code writes it again all the same: another language, or text that no longer
  // fits a larger size. auto and 0 state nothing there either, and the slide stays set out as it was.
  const french = { s1: { title: 'Magasins' }, s2: { ...before.slots, title: 'Ventes par magasin' }, s3: { title: 'Merci' } };
  assert.deepEqual((await arrangeAfter({}, french, {}, { language: 'fr' })).arrange, set);
  const long = storesDeck({ ...before, slots: { ...before.slots, title: 'Sales by store and what we learnt from them this year' } });
  const larger = scripted({ ...changeOf(stay(1, 'cover'), stay(2, 'free'), stay(3, 'closing')), text_size: 'large' }, { s2: { title: 'Sales by store' } });
  const sized = await planEdit(demo, long, 'Larger text', { llm: larger.llm, ...adaptive });
  assert.deepEqual([sized.rewritten, sized.deck.slides[1]?.arrange], [[2], set]);
});

test('a slide of a pattern becomes a free slide, and a new free slide is written from its parts', async () => {
  // One chart on a slide of the template; the request wants a second beside it.
  const north = chartOf('North');
  const deck: DeckSpec = { ...storesDeck(), slides: [{ pattern: 'cover', slots: { title: 'Stores' } }, { pattern: 'chart-native', slots: { title: 'Sales in the north', chart: north, chartDescription: 'The north sold most in the last quarter.' } }, { pattern: 'closing', slots: { title: 'Thanks' } }] };
  const south = { type: 'column', categories: quarters, series: [{ name: 'South', values: [18, 14, 12, 10] }] };
  const freed = stay(2, 'free', { change: 'Add a chart for the south beside the one for the north', parts: [part(0, 'chart'), part(0, 'chart', true), part(0, 'note', true)] });
  const { llm, requests } = scripted(changeOf(stay(1, 'cover'), freed, stay(3, 'closing')), { s2: { title: 'North and south', b1Caption: 'North', b2Caption: 'South', b2Chart: south, b3Text: 'The two moved in opposite ways.' } });
  const result = await planEdit(demo, deck, 'Add a chart for the south (18, 14, 12, 10) beside the one on slide 2', { llm, ...adaptive });

  // Its chart goes to the first part that is a chart and is not to be written again: no model copies its numbers.
  const s2 = schemaOf(requests[1]).properties.s2 as Schema;
  assert.deepEqual(Object.keys(s2.properties), ['title', 'b1Caption', 'b2Caption', 'b2Chart', 'b3Text']);
  const task = promptOf(requests[1]);
  assert.ok(
    task.includes(
      '\ns2 (free: b1 to b2: 2 charts; b3: a note). Change: Add a chart for the south beside the one for the north\n' +
        // Its caption is written, so the part is not among what "is not to be written": the chart it keeps is said by itself.
        'Of the parts that are written, this stays as it is: 1 chart (column, 4 categories, series "North")\n' +
        'Content of the slide it replaces: {"title":"Sales in the north","chartDescription":"The north sold most in the last quarter."}\n',
    ),
    task.slice(-600),
  );
  assert.ok(!task.includes('[10,14,12,18]'));
  assert.deepEqual(result.deck.slides[1], {
    pattern: 'free',
    blocks: ['chart', 'chart', 'note'],
    base: 'list-with-photo',
    slots: { b1Chart: north, title: 'North and south', b1Caption: 'North', b2Caption: 'South', b2Chart: south, b3Text: 'The two moved in opposite ways.' },
  });
  assert.deepEqual(result.warnings.filter((issue) => issue.slot !== undefined), []);

  // A chart the plan asks to be written again is written from the slide's content, which is then shown whole.
  const again = scripted(changeOf(stay(1, 'cover'), stay(2, 'free', { change: 'Two charts', parts: [part(0, 'chart', true), part(0, 'chart', true)] }), stay(3, 'closing')), {
    s2: { title: 'North and south', b1Caption: 'North', b1Chart: north, b2Caption: 'South', b2Chart: south },
  });
  await planEdit(demo, deck, 'Two charts on slide 2: north as it is, south with 18, 14, 12, 10', { llm: again.llm, ...adaptive });
  assert.ok(promptOf(again.requests[1]).includes(`Content of the slide it replaces: {"title":"Sales in the north","chart":${JSON.stringify(north)},`));
  // The plan is told how to say which it is: a new part is otherwise marked as one to write, and the chart would be copied by a model.
  assert.ok(
    promptOf(again.requests[0]).includes(
      '\n- A slide of a pattern that takes the layout free hands its chart, its table and its photos to the first parts of those kinds that it names. Give such a part rewrite false unless the request changes what it shows.\n',
    ),
  );

  // "Part 1, as it is" on a slide that has no numbered parts: asked about once, and then taken as it was meant.
  // The part keeps the answer it gave on being written again, so the chart is handed to it by code.
  const pointing = changeOf(stay(1, 'cover'), stay(2, 'free', { change: 'Add a chart for the south', parts: [part(1, 'chart'), part(0, 'chart', true)] }), stay(3, 'closing'));
  const meant = scripted(pointing, pointing, { s2: { title: 'North and south', b1Caption: 'North', b2Caption: 'South', b2Chart: south } });
  const handed = await planEdit(demo, deck, 'Add a chart for the south (18, 14, 12, 10) beside the one on slide 2', { llm: meant.llm, ...adaptive });
  assert.ok(promptOf(meant.requests[1]).endsWith('Your previous answer could not be used: slide 2: the slide it is based on has no numbered parts, so "from" of each of its parts must be 0. Answer again.'));
  assert.deepEqual(Object.keys((schemaOf(meant.requests[2]).properties.s2 as Schema).properties), ['title', 'b1Caption', 'b2Caption', 'b2Chart']);
  assert.ok(!promptOf(meant.requests[2]).includes('[10,14,12,18]'), 'no model is shown the numbers of the chart that stays');
  assert.equal(JSON.stringify(handed.deck.slides[1]?.slots.b1Chart), JSON.stringify(north));

  // A new slide: every part is written, and the slide stands where the plan puts it.
  const cards = { title: 'Why stores matter', b1Heading: 'Near', b1Text: 'Close to where people live.', b2Heading: 'Seen', b2Text: 'People can try things.', b3Heading: 'Fast', b3Text: 'Bought and taken home.' };
  const added = scripted(changeOf(stay(1, 'cover'), stay(2, 'chart-native'), stay(0, 'free', { change: 'Three cards on why stores matter', parts: [part(0, 'card', true), part(0, 'card', true), part(0, 'card', true)], side_by_side: 3 }), stay(3, 'closing')), { s3: cards });
  const grown = await planEdit(demo, deck, 'Add a slide with three cards on why stores matter', { llm: added.llm, ...adaptive });
  assert.match(promptOf(added.requests[1]), /\ns3 \(free: b1 to b3: 3 cards\)\. New slide: Three cards on why stores matter\n\nFill every required property\./);
  assert.deepEqual(grown.deck.slides[2], { pattern: 'free', blocks: ['card', 'card', 'card'], arrange: { columns: 3 }, base: 'list-with-photo', slots: cards });
  assert.deepEqual(grown.rewritten, [3]);

  // A new slide of charts for which the request gives no numbers: the writer leaves them out, as it is told to. The slide
  // then says the title it was given, once. It does not show the words of the plan, which say what to do with the slide.
  const charts = stay(0, 'free', { change: 'Three charts of sales, one for each region', parts: [part(0, 'chart', true), part(0, 'chart', true), part(0, 'chart', true)] });
  const empty = scripted(changeOf(stay(1, 'cover'), stay(2, 'chart-native'), charts, stay(3, 'closing')), { s3: { title: 'Sales by region' } });
  const lone = await planEdit(demo, deck, 'Add a slide with three charts, one for each region', { llm: empty.llm, ...adaptive });
  assert.deepEqual(lone.deck.slides[2], { pattern: 'free', blocks: ['text'], base: 'list-with-photo', slots: { b1Text: 'Sales by region' } });
  assert.deepEqual(withoutCodes(lone.warnings.filter((issue) => /left out/.test(issue.message))), [{ severity: 'warning', message: 'slide 3: 3 parts had no content and were left out' }]);
  // Without a title it says the title of the deck, and the words of the plan only when there is nothing else at all.
  const mute = scripted(changeOf(stay(1, 'cover'), stay(2, 'chart-native'), charts, stay(3, 'closing')), { s3: {} });
  assert.deepEqual((await planEdit(demo, deck, 'Add a slide with three charts, one for each region', { llm: mute.llm, ...adaptive })).deck.slides[2]?.slots, { b1Text: 'Stores' });
});

test('what a plan says wrongly of the parts of a free slide is asked for once more, and then put right by code', async () => {
  const deck = storesDeck();
  const again = (problem: string) => `\n\nYour previous answer could not be used: ${problem}. Answer again.`;
  const note = { s2: { title: 'Sales by store', b7Text: 'Online grew fastest of all the stores this year.' } };

  // A part that points at a part the slide does not have.
  const astray = changeOf(stay(1, 'cover'), stay(2, 'free', { change: 'A new note', parts: sixAndNote({ 7: part(9, 'note', true) }), side_by_side: 3 }), stay(3, 'closing'));
  const mended = scripted(astray, changeOf(stay(1, 'cover'), stay(2, 'free', { change: 'A new note', parts: sixAndNote({ 7: part(7, 'note', true) }), side_by_side: 3 }), stay(3, 'closing')), note);
  await planEdit(demo, deck, 'A new note on slide 2', { llm: mended.llm, ...adaptive });
  assert.ok(promptOf(mended.requests[1]).endsWith(again('slide 2: "from" of a part must be 0 or the number of an existing part (1 to 7), not 9')));
  // Said twice, it is a new part: written from the change, with nothing kept for it. No change fails over the parts of a free slide.
  const stubborn = scripted(astray, astray, note);
  const result = await planEdit(demo, deck, 'A new note on slide 2', { llm: stubborn.llm, ...adaptive });
  assert.equal(result.deck.slides[1]?.slots.b7Text, note.s2.b7Text);
  assert.ok(!promptOf(stubborn.requests[2]).includes('Current content: {"b7Text"'), 'the note is written anew');

  // A plan that names the one part that changes, and not the six that stay, would take the charts off the slide. That is asked about once.
  const only = changeOf(stay(1, 'cover'), stay(2, 'free', { change: 'A new note', parts: [part(7, 'note', true)], side_by_side: 3 }), stay(3, 'closing'));
  const whole = scripted(only, changeOf(stay(1, 'cover'), stay(2, 'free', { change: 'A new note', parts: sixAndNote({ 7: part(7, 'note', true) }), side_by_side: 3 }), stay(3, 'closing')), note);
  const complete = await planEdit(demo, deck, 'A new note on slide 2', { llm: whole.llm, ...adaptive });
  assert.ok(promptOf(whole.requests[1]).endsWith(again('slide 2: 6 of the 7 parts of the slide are not named in "parts" and would be taken off it; name every part that stays, also those that do not change. If they are to go, answer the same again')));
  assert.equal(complete.deck.slides[1]?.blocks?.length, 7);
  // The same answer again is taken at its word: the note is all the slide keeps.
  const meant = scripted(only, only, { s2: { title: 'Sales by store', b1Text: note.s2.b7Text } });
  assert.deepEqual((await planEdit(demo, deck, 'Only the note on slide 2, reworded', { llm: meant.llm, ...adaptive })).deck.slides[1]?.blocks, ['note']);
  // One part of three taken off is no such case, and needs no second answer: the photo that was taken off in the test above.

  // A part that is named by its number and whose kind is left out, or is no kind, is the part of that number: it stays
  // what it is. Passed over, it would be a chart taken off the slide that nobody asked to remove.
  const kindless = changeOf(stay(1, 'cover'), stay(2, 'free', { change: 'A new note', parts: sixAndNote({ 3: { from: 3, rewrite: false }, 5: part(5, 'charts'), 7: part(7, 'note', true) }), side_by_side: 3 }), stay(3, 'closing'));
  const named = scripted(kindless, note);
  const all = await planEdit(demo, deck, 'A new note on slide 2', { llm: named.llm, ...adaptive });
  assert.equal(named.requests.length, 2, 'every part is there, so nothing is asked about');
  assert.deepEqual(all.deck.slides[1]?.blocks, deck.slides[1]?.blocks);
  for (const chart of ['b3Chart', 'b5Chart']) assert.equal(JSON.stringify(all.deck.slides[1]?.slots[chart]), JSON.stringify(deck.slides[1]?.slots[chart]));
  // A new part without a usable kind names nothing, and is passed over as in an outline.
  const nameless = changeOf(stay(1, 'cover'), stay(2, 'free', { change: 'A new note', parts: [...sixAndNote({ 7: part(7, 'note', true) }), { from: 0, rewrite: true }], side_by_side: 3 }), stay(3, 'closing'));
  assert.equal((await planEdit(demo, deck, 'A new note on slide 2', { llm: scripted(nameless, note).llm, ...adaptive })).deck.slides[1]?.blocks?.length, 7);

  // A slide that is new, or was a slide of a pattern, has no parts to point at, and needs some.
  const bare = changeOf(stay(1, 'cover'), stay(2, 'free', { side_by_side: 3 }), stay(0, 'free', { change: 'Why stores matter' }), stay(3, 'closing'));
  const text = { s3: { title: 'Why stores matter', b1Text: ['They are near.'] } };
  const empty = scripted(bare, bare, text);
  const one = await planEdit(demo, deck, 'Add a slide on why stores matter', { llm: empty.llm, ...adaptive });
  assert.ok(promptOf(empty.requests[1]).endsWith(again('slide 3: a slide that becomes a free slide needs "parts", one entry for each part it shows')));
  assert.deepEqual(one.deck.slides[2], { pattern: 'free', blocks: ['text'], base: 'list-with-photo', slots: { title: 'Why stores matter', b1Text: ['They are near.'] } });
  const pointing = changeOf(stay(1, 'cover'), stay(2, 'free', { side_by_side: 3 }), stay(0, 'free', { change: 'Why stores matter', parts: [part(2, 'text', true)] }), stay(3, 'closing'));
  const told = scripted(pointing, pointing, text);
  await planEdit(demo, deck, 'Add a slide on why stores matter', { llm: told.llm, ...adaptive });
  assert.ok(promptOf(told.requests[1]).endsWith(again('slide 3: the slide it is based on has no numbered parts, so "from" of each of its parts must be 0')));

  // More photos than can be made for one slide: said once, then the new ones that are too many are left out.
  const photos = changeOf(stay(1, 'cover'), stay(2, 'free', { side_by_side: 3 }), stay(0, 'free', { change: 'Eight dogs', parts: Array.from({ length: 8 }, () => part(0, 'photo', true)) }), stay(3, 'closing'));
  const dogs = Object.fromEntries(range(1, 6).map((number) => [`b${number}Photo`, { prompt: `Dog ${number}`, alt: 'A dog' }]));
  const many = scripted(photos, photos, { s3: { title: 'Dogs', ...dogs } });
  const six = await planEdit(demo, deck, 'Add a slide with photos of eight dogs', { llm: many.llm, ...adaptive });
  assert.ok(promptOf(many.requests[1]).endsWith(again('slide 3: at most 6 photos can be made for one slide; supply pictures for more')));
  assert.equal(six.deck.slides[2]?.blocks?.length, 6);

  // A kind that is not on offer is no part: a chart, in a deck whose brief and request hold no number.
  const wordy: DeckSpec = { version: 1, title: 'Dogs', language: 'en', direction: 'ltr', brief: 'A deck about dogs', fields, slides: [{ pattern: 'cards', items: 2, slots: cardsContent }, { pattern: 'closing', slots: {} }] };
  const charts = { summary: 'Done.', language: 'en', title: 'Dogs', text_size: 'normal', fields, slides: [stay(1, 'cards'), stay(0, 'free', { change: 'Why dogs', parts: [part(0, 'chart', true), part(0, 'note', true)] }), stay(2, 'closing')] };
  const none = scripted(charts, { s2: { b1Text: 'Because.' } });
  const noted = await planEdit(template, wordy, 'Add a slide on why dogs', { llm: none.llm });
  assert.deepEqual(entryOf(none.requests[0]).properties.parts?.items?.properties.kind?.enum, ['text', 'list', 'card', 'step', 'quote', 'note', 'photo']);
  assert.deepEqual(noted.deck.slides[1]?.blocks, ['note']);
});

test('a deck keeps and changes its free slides where none is on offer for a new slide, and without one an edit is asked as before', async () => {
  const { deck } = await writeDeck(template, cardsPlan({ arrange: { columns: 3 } }), { llm: scripted({ s1: cardsContent, s2: threeCards, s3: {} }).llm });
  const free = JSON.stringify(deck.slides[1]);
  const answer = (...slides: unknown[]) => ({ summary: 'Done.', language: 'en', title: 'Dogs', text_size: 'normal', fields, slides });

  // Turned off by the app, or outside the list the deck is limited to: the slide is listed, may be named and changed, and stays a free slide.
  for (const options of [{ free: false }, { deckPatterns: ['cards', 'closing'] }, { patterns: ['cards', 'closing'] }]) {
    const { llm, requests } = scripted(
      answer(stay(1, 'cards'), stay(2, 'free', { change: 'A new second card', parts: [part(1, 'card'), part(2, 'card', true), part(3, 'card')], side_by_side: 3 }), stay(3, 'closing')),
      { s2: { title: 'Three ways in', b2Heading: 'Brighter docks', b2Text: 'Two lights at every station.' } },
    );
    const result = await planEdit(template, deck, 'A new second card on slide 2', { llm, ...options });
    assert.ok(entryOf(requests[0]).properties.pattern?.enum?.includes('free'));
    assert.deepEqual(entryOf(requests[0]).properties.parts?.items?.properties.kind?.enum, ['card'], 'the kinds the deck uses, and no other');
    assert.equal(
      freeLineOf(requests[0]),
      '- free: A slide the software lays out itself from the parts you name in "parts", any number of each: card (a heading and a short text in a box). Only for the slides of the deck that have it already.',
    );
    assert.deepEqual(result.deck.slides[1]?.slots, { ...threeCards, b2Heading: 'Brighter docks', b2Text: 'Two lights at every station.' });
    assert.equal(result.deck.slides[1]?.base, 'content');
    // The rules speak of the parts of the slide the deck has, and not of a slide that becomes a free slide: none can here.
    assert.ok(promptOf(requests[0]).includes('\n- parts: for a slide with the layout free whose parts change. Leave it [] otherwise.\n- A slide with the layout free whose parts are only taken off or moved needs nothing written: '));
    assert.ok(!promptOf(requests[0]).includes('takes the layout free'));

    // A new free slide is not to be had there: said once, and then the change fails like any plan that cannot be used.
    const extra = answer(stay(1, 'cards'), stay(2, 'free', { side_by_side: 3 }), stay(0, 'free', { change: 'More', parts: [part(0, 'card', true)] }), stay(3, 'closing'));
    const refused = scripted(extra, extra);
    await assert.rejects(planEdit(template, deck, 'Add a slide of cards', { llm: refused.llm, ...options }), /^PlanError: The model did not produce a usable edit plan: slide 3: pattern "free" cannot be used for a new or re-laid-out slide here$/);
  }

  // A deck without a free slide, with free slides turned off: the plan of a change is asked exactly as it was before there were any.
  const plain: DeckSpec = { ...deck, slides: [deck.slides[0] as SlideSpec, deck.slides[2] as SlideSpec] };
  const off = scripted(answer(stay(1, 'cards', { change: 'A shorter title' }), stay(2, 'closing')), { s1: { ...cardsContent, title: 'Two!' } });
  await planEdit(template, plain, 'Shorten the title of slide 1', { llm: off.llm, free: false });
  assert.equal(off.requests[0]?.prompt, EDIT_PROMPT_BEFORE);
  assert.equal(JSON.stringify(off.requests[0]?.schema), JSON.stringify(EDIT_SCHEMA_BEFORE));
  assert.equal(off.requests[0]?.maxOutputTokens, (400 + 12 * 160) * 1.5);
  assert.equal(off.requests[1]?.prompt, WRITE_PROMPT_BEFORE);
  // Under a list that does not name free slides it is the same request, but for the rule about the list.
  const listed = scripted(answer(stay(1, 'cards', { change: 'A shorter title' }), stay(2, 'closing')), { s1: { ...cardsContent, title: 'Two!' } });
  await planEdit(template, plain, 'Shorten the title of slide 1', { llm: listed.llm, patterns: ['content', 'closing', 'cards'] });
  assert.deepEqual(listed.requests, off.requests);
  assert.equal(JSON.stringify(deck.slides[1]), free, 'the deck given is not changed');
});

test('an edit copies a free slide it does not touch, and writes the slides around it', async () => {
  const { deck } = await writeDeck(template, cardsPlan({ arrange: { columns: 3 } }), { llm: scripted({ s1: cardsContent, s2: threeCards, s3: {} }).llm });
  const free = JSON.stringify(deck.slides[1]);
  const keep = (from: number, pattern: string, change = '') => ({ from, pattern, change, images: false });
  const edit = (...slides: unknown[]) => ({ summary: 'Done.', language: 'en', title: 'Dogs', text_size: 'normal', fields, slides });

  // Moved: the plan of the edit may name the slide, and code copies it, with its parts, its arrangement and its frame.
  const move = scripted(edit(keep(3, 'closing'), keep(2, 'free'), keep(1, 'cards')));
  const moved = await planEdit(template, deck, 'Open with the closing slide', { llm: move.llm });
  assert.deepEqual(entryOf(move.requests[0]).properties.pattern?.enum, ['content', 'closing', 'cards', 'free']);
  assert.equal(move.requests.length, 1, 'nothing is written');
  assert.deepEqual(moved.deck.slides.map((slide) => slide.pattern), ['closing', 'free', 'cards']);
  assert.equal(JSON.stringify(moved.deck.slides[1]), free, 'byte for byte');

  // Another slide is rewritten: the free slide beside it stays as it is, and the deck is checked with it in place.
  const rewrite = scripted(edit(keep(1, 'cards', 'A shorter title'), keep(2, 'free'), keep(3, 'closing')), { s1: { ...cardsContent, title: 'Two!' } });
  const rewritten = await planEdit(template, deck, 'Shorten the title of slide 1', { llm: rewrite.llm });
  assert.deepEqual(rewritten.rewritten, [1]);
  assert.equal(rewritten.deck.slides[0]?.slots.title, 'Two!');
  assert.equal(JSON.stringify(rewritten.deck.slides[1]), free);
  assert.deepEqual(rewritten.warnings, []);
  assert.deepEqual(Object.keys(schemaOf(rewrite.requests[1]).properties), ['s1'], 'only the slide that changes is asked for');
});

// --- the chat tools ---------------------------------------------------------------------------------

type Definition = { name: string; description: string; parameters: { properties: Record<string, Schema>; required: string[] } };
const toolNamed = (tools: { definitions: unknown[] }, name: string) => (tools.definitions as Definition[]).find((tool) => tool.name === name) as Definition;

test('the tools show a free slide in a plan and in a presentation, and name the part a warning is about', async () => {
  // The text of the second card stays far too long, through both repair rounds.
  const long = { s1: cardsContent, s2: { ...threeCards, b2Text: 'Bikes unlock from five. '.repeat(40) } };
  const { llm, requests } = scripted(outlineOf(entry('cards', 'Two things', { items: 2 }), freeEntry('Three ways in', [['card', 3]], 'auto', 0)), long, long, long);
  let next = 0;
  const store = memoryStore();
  const tools = presentationTools({ template, llm, store, fields, thumbnail: (pattern) => `/thumbs/${pattern}.png`, newId: (kind) => `${kind}_${++next}` });

  // The declarations: a slide can hold what no layout has, and `free` is one more layout to name.
  const plan = toolNamed(tools, TOOL_NAMES.plan);
  // These tools have no image model and no pictures of their own, so no slide of theirs gets a
  // picture, and the declarations promise none: the planner of the same tools is told "Do not plan any photos".
  const holds = (several: string) => `A slide can also hold what no layout has (${several} on one slide, or any number of cards or figures): pass such a wish on in instructions, in the user's words. `;
  assert.ok(plan.description.includes(holds('several charts or tables')), plan.description);
  assert.doesNotMatch(JSON.stringify(tools.definitions), /pictures on one slide/);
  // With an image model, or with a photo or a graphic of the app's own, pictures are named too.
  const drawn = placeholderPhoto(20, 10, [0, 0, 0], [255, 255, 255]);
  const painter = { generateImage: async () => drawn };
  const graphic = { id: 'map', kind: 'graphic' as const, description: 'A map of the town', data: drawn };
  for (const supplied of [{ images: painter }, { assets: [graphic] }, { assets: [{ ...graphic, kind: 'photo' as const }] }]) {
    const seeing = presentationTools({ template, llm, store: memoryStore(), fields, ...supplied });
    assert.ok(toolNamed(seeing, TOOL_NAMES.plan).description.includes(holds('several charts, tables or pictures')));
    assert.ok(toolNamed(seeing, TOOL_NAMES.edit).description.includes('the number of cards on it, put several charts, tables or pictures on one slide, make the text larger or smaller, '));
  }
  assert.deepEqual(plan.parameters.properties.layouts?.items?.enum, ['closing', 'cards', 'free']);
  assert.ok((plan.parameters.properties.layouts?.description ?? '').endsWith('\n- cards: A row of cards.\n- free: A slide laid out by the software from the parts asked for.'));

  // A plan: the slide is called free, and described by what it holds. There is no picture of a layout to show.
  const planned = await tools.call(TOOL_NAMES.plan, { topic: 'A deck about dogs', language: 'en', slide_count: 2 });
  assert.deepEqual(planned.output.slides, [
    { number: 1, layout: 'cards', layout_description: 'A row of cards.', purpose: 'Two things', thumbnail: '/thumbs/cards.png' },
    { number: 2, layout: 'free', layout_description: 'A slide laid out for this presentation: 3 cards.', purpose: 'Three ways in', free_layout: true },
  ]);
  assert.match(promptOf(requests[0]), /\n- free: A slide the software lays out itself/);

  // A presentation: the same, with its parts in words. A warning names the part, not the slot.
  const created = await tools.call(TOOL_NAMES.create, { plan_id: 'plan_1' });
  assert.equal(created.output.error, undefined, String(created.output.error));
  assert.deepEqual(created.output.slides, [
    { number: 1, layout: 'cards', title: 'Two' },
    { number: 2, layout: 'free', title: 'Three ways in', free_layout: true, parts: '3 cards' },
  ]);
  const warnings = created.output.warnings as string[];
  assert.equal(warnings.length, 1);
  assert.match(warnings[0] ?? '', /^slide 2, part 2 \(card\) text: text is \d+ characters; the slot fits about \d+$/);
  const saved = JSON.parse(new TextDecoder().decode(await store.get(storeKeys.deck('deck_2')))) as DeckSpec;
  assert.deepEqual(saved.slides[1]?.blocks, ['card', 'card', 'card']);

  // A person may limit a presentation to free slides, like to any layout.
  const only = scripted(outlineOf(freeEntry('Why dogs', [['text', 1]])));
  const limited = presentationTools({ template, llm: only.llm, store: memoryStore(), fields });
  const listed = await limited.call(TOOL_NAMES.plan, { topic: 'Dogs', layouts: ['free'], slide_count: 1 });
  assert.deepEqual(listed.plan?.brief.patterns, ['free']);
  assert.deepEqual(listed.output.layouts, ['free']);
  assert.deepEqual(offered(only.requests[0]), ['free']);
});

test('tools with free slides turned off, or under a list that leaves them out, say nothing of them', async () => {
  for (const options of [{ free: false }, { patterns: ['cards', 'closing'] }]) {
    const { llm, requests } = scripted(outlineOf({ pattern: 'cards', purpose: 'Two things', items: 2 }));
    const tools = presentationTools({ template, llm, store: memoryStore(), fields, ...options });
    assert.doesNotMatch(JSON.stringify(tools.definitions), /free|no layout has|several charts/);
    assert.deepEqual(toolNamed(tools, TOOL_NAMES.plan).parameters.properties.layouts?.items?.enum, ['closing', 'cards']);
    await tools.call(TOOL_NAMES.plan, { topic: 'A deck about dogs', language: 'en', slide_count: 1 });
    assert.doesNotMatch(JSON.stringify(requests[0]), /free|shows/);
    // Named all the same, it is refused with the reason.
    const refused = await tools.call(TOOL_NAMES.plan, { topic: 'Dogs', layouts: ['free'] });
    assert.equal(
      refused.output.error,
      `"layouts" names a layout that cannot be used here: free ${'free' in options ? '(free slides are turned off)' : 'is not among the layouts allowed here'}. The layouts to choose from: closing, cards`,
    );
  }
  // A plan that was stored with a free slide is shown and written all the same.
  const store = memoryStore();
  await store.set(storeKeys.plan('plan_0'), new TextEncoder().encode(JSON.stringify(cardsPlan())));
  const tools = presentationTools({ template, llm: scripted({ s1: cardsContent, s2: threeCards, s3: {} }).llm, store, fields, free: false, newId: () => 'deck_1' });
  const created = await tools.call(TOOL_NAMES.create, { plan_id: 'plan_0' });
  assert.equal(created.output.error, undefined, String(created.output.error));
  assert.deepEqual((created.output.slides as Array<Record<string, unknown>>)[1], { number: 2, layout: 'free', title: 'Three ways in', free_layout: true, parts: '3 cards' });
});

test('the tools change a presentation that has a free slide: a part is added, and the slide says what it holds now', async () => {
  const { deck } = await writeDeck(template, cardsPlan({ arrange: { columns: 3 } }), { llm: scripted({ s1: cardsContent, s2: threeCards, s3: {} }).llm });
  const store = memoryStore();
  await store.set(storeKeys.deck('deck_0'), new TextEncoder().encode(JSON.stringify(deck)));
  const fourth = { title: 'Four ways in', b4Heading: 'Free repairs', b4Text: 'A mechanic at every hub.' };
  const { llm, requests } = scripted(
    {
      summary: 'Added a fourth card to slide 2.',
      language: 'en',
      title: 'Dogs',
      text_size: 'normal',
      fields,
      slides: [stay(1, 'cards'), stay(2, 'free', { change: 'A fourth card: free repairs', parts: [part(1, 'card'), part(2, 'card'), part(3, 'card'), part(0, 'card', true)], side_by_side: 4 }), stay(3, 'closing')],
    },
    { s2: fourth },
  );
  const tools = presentationTools({ template, llm, store, fields, newId: () => 'deck_1' });
  // The tool says that a slide can be given several charts or tables: what a description leaves out, a chat model takes for
  // impossible. Pictures are not named here, where the tools have no image model and no pictures of their own.
  assert.ok(toolNamed(tools, TOOL_NAMES.edit).description.includes('change the layout of a slide or the number of cards on it, put several charts or tables on one slide, make the text larger or smaller, '));

  const edited = await tools.call(TOOL_NAMES.edit, { presentation_id: 'deck_0', instruction: 'A fourth card on slide 2: free repairs' });
  assert.equal(edited.output.error, undefined, String(edited.output.error));
  assert.equal(requests.length, 2);
  assert.deepEqual(edited.output.slides, [
    { number: 1, layout: 'cards', title: 'Two' },
    { number: 2, layout: 'free', title: 'Four ways in', free_layout: true, parts: '4 cards' },
    { number: 3, layout: 'closing' },
  ]);
  assert.deepEqual([edited.output.summary, edited.output.rewritten_slides, edited.output.warnings], ['Added a fourth card to slide 2.', [2], undefined]);
  assert.deepEqual(edited.deck?.slides[1], { pattern: 'free', blocks: ['card', 'card', 'card', 'card'], arrange: { columns: 4 }, base: 'content', slots: { ...threeCards, ...fourth } });
  assert.ok((edited.pptx?.length ?? 0) > 1000);

  // A plan that has words for the slide, marks no part to be written and gets the title back as it was has changed
  // nothing. The chat model is told so in a warning that names the slide in its words, so that it does not report
  // the summary of the plan as done.
  const idle = scripted(
    { summary: 'Reworded the second card.', language: 'en', title: 'Dogs', text_size: 'normal', fields, slides: [stay(1, 'cards'), stay(2, 'free', { change: 'Reword the second card', parts: [part(1, 'card'), part(2, 'card'), part(3, 'card')], side_by_side: 3 }), stay(3, 'closing')] },
    { s2: { title: 'Three ways in' } },
  );
  const same = await presentationTools({ template, llm: idle.llm, store, fields, newId: () => 'deck_2' }).call(TOOL_NAMES.edit, { presentation_id: 'deck_0', instruction: 'Reword the second card on slide 2' });
  assert.equal(same.output.error, undefined, String(same.output.error));
  assert.deepEqual(same.output.warnings, ['slide 2: nothing on it was changed: the plan of the change named no part to write again']);
  assert.deepEqual(same.output.rewritten_slides, []);
  assert.deepEqual(same.deck?.slides[1], deck.slides[1]);
});

// --- what the documentation says of all this ---------------------------------------------------------

test('the pages on planning show the answer a model gives and quote the messages as the code gives them', async () => {
  const page = (name: string) => readFileSync(join(root, 'docs', name), 'utf8').replace(/\s+/g, ' ');
  const pages = ['free-slides.md', 'generating.md', 'chat-tools.md', 'api-reference.md'].map((name) => [name, page(name)] as const);

  // The entry of an answer that free-slides.md shows is read into the slide of a plan it shows beside it.
  const lines = readFileSync(join(root, 'docs', 'free-slides.md'), 'utf8').split('\n');
  const shown = JSON.parse(lines.find((line) => line.startsWith('{ "pattern": "free"') && line.includes('"shows"')) ?? 'null') as Record<string, unknown>;
  const planned = JSON.parse(lines.find((line) => line.startsWith('{ "pattern": "free"') && line.includes('"blocks"')) ?? 'null') as OutlineEntry;
  assert.deepEqual(Object.keys(shown), ['pattern', 'purpose', 'shows', 'shows_main', 'side_by_side']);
  const { llm, requests } = scripted({ title: 'Stores', slides: [shown] });
  const plan = await planOutline(demo, { ...stores, slideCount: 1 }, { llm, layout: 'adaptive' });
  assert.deepEqual(plan.slides, [planned]);
  // The properties of the entry are those the schema asks for, the count of a run under the name it has there.
  const entry = entryOf(requests[0]);
  assert.deepEqual(entry.required, Object.keys(shown));
  assert.deepEqual(entry.properties.shows?.items?.required, Object.keys((shown.shows as unknown[])[0] as object));
  for (const [name, text] of pages) assert.doesNotMatch(text, /"shows": \[\{[^\]]*"number"/, `${name} shows the count of a run under its old name`);
  assert.match(page('generating.md'), /"shows": \[\{ "count": 6, "kind": "chart" \}, \{ "count": 1, "kind": "note" \}\]/);

  // What the check of chart numbers does and does not tell is said wherever the pages say that numbers must be in the brief.
  for (const [name, text] of pages) assert.match(text, /does not tell whether (?:it|a number|the number) stands in the right chart/, name);

  // The note on a slide of which nothing changed, in the words of the code. The reference gives its start.
  const deck = storesDeck();
  const idle = scripted(changeOf(stay(1, 'cover'), stay(2, 'free', { change: 'Rename a chart', parts: sixAndNote(), side_by_side: 3 }), stay(3, 'closing')), { s2: {} });
  const told = (await planEdit(demo, deck, 'Rename a chart on slide 2', { llm: idle.llm, ...adaptive })).warnings.find((issue) => /nothing on it/.test(issue.message))?.message ?? '';
  for (const name of ['free-slides.md', 'generating.md', 'chat-tools.md']) assert.ok(page(name).includes(`\`${told.replace('slide 2', 'slide 3')}\``), name);
  assert.ok(page('api-reference.md').includes('`slide 3: nothing on it was changed: ...`'));

  // The line of a later call that generating.md quotes is a line of such a call.
  const charts = (request: JsonRequest): unknown =>
    's1' in schemaOf(request).properties ? cover : { r01: range(Number(/Write parts (\d+) to/.exec(request.prompt)?.[1]), Number(/Write parts \d+ to (\d+)/.exec(request.prompt)?.[1])).map(sparkItem), title: 'Every store' };
  const sixty = scripted(charts, charts, charts);
  await writeDeck(demo, { ...chartsDeck(0), slides: [{ pattern: 'cover', purpose: 'Open' }, { pattern: 'free', purpose: 'Every store', blocks: Array<'chart'>(60).fill('chart') }] }, { llm: sixty.llm, layout: 'adaptive' });
  const quoted = /`(- r01: parts 31 to 60, [^`]*)`/.exec(page('generating.md'))?.[1] ?? 'no such line in generating.md';
  assert.ok(sixty.requests.some((request) => request.prompt.includes(`\n${quoted}\n`)), quoted);
  assert.ok(page('free-slides.md').includes('(`the 31st to the 60th chart of the slide`)'));
});
