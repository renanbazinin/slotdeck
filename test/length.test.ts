// An amount of text a person asks for on a slide ("200 words"): the plan reads it, gives the text a
// slide that holds it, and the writer is told the length before what a slot says of length. No test
// calls a live model.
import assert from 'node:assert/strict';
import { readFileSync } from 'node:fs';
import { dirname, join } from 'node:path';
import { test } from 'node:test';
import { fileURLToPath } from 'node:url';
import { TOOL_NAMES, TextCutError, editPlan, freeSlide, loadTemplate, memoryStore, planEdit, planOutline, planTextFit, presentationTools, writeDeck } from '../src/index.ts';
import type { Brief, DeckPlan, DeckSpec, JsonRequest, LlmProvider, SlideTextFit, TemplateManifest } from '../src/index.ts';
import { slotNote } from '../src/plan/catalog.ts';
import type { CatalogContext } from '../src/plan/catalog.ts';
import { freeWritePattern } from '../src/plan/free.ts';
import { WORD_CHARS, charsOfWords, readWords, withoutLength } from '../src/plan/length.ts';
import type { TextSlotDef } from '../src/template/manifest.ts';
import { buildFixturePptx, fixtureManifest } from './fixture.ts';

const root = join(dirname(fileURLToPath(import.meta.url)), '..');
const demoManifest = JSON.parse(readFileSync(join(root, 'templates', 'demo.manifest.json'), 'utf8')) as TemplateManifest;
const demoPptx = new Uint8Array(readFileSync(join(root, 'templates', 'demo.pptx')));
const demo = loadTemplate(demoPptx, demoManifest);
const fields = { organizationDepartment: 'Retail' };
const brief: Brief = { prompt: 'Why the city should keep its bike share', language: 'en', fields, imageDensity: 'none', slideCount: 1 };
const options = { layout: 'adaptive' as const, fields };

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

type Schema = { properties: Record<string, Schema>; description?: string; maxItems?: number };
const slotOf = (request: JsonRequest | undefined, slide: string, slot: string): Schema | undefined => (request?.schema as unknown as Schema | undefined)?.properties[slide]?.properties[slot];
/** An outline entry as a model answers it where free slides are on offer. */
const entry = (pattern: string, more: Record<string, unknown> = {}) => ({ pattern, purpose: 'Keep the bike share', shows: [], shows_main: 'auto', side_by_side: 0, ...more });
const outlineOf = (...slides: unknown[]) => ({ title: 'Bikes', slides });
/** Running text of so many characters, in paragraphs of about 150. */
const textOf = (chars: number): string[] => {
  const count = Math.max(1, Math.round(chars / 150));
  return Array.from({ length: count }, () => 'Riders rely on it daily. '.repeat(10).slice(0, Math.floor(chars / count) - 1).trim() + '.');
};
/** The limit a slot states after its description: "at most 1281 characters". */
const limitOf = (description: string | undefined): number => Number(/at most (\d+) characters/.exec(description ?? '')?.[1] ?? 0);

const textOnly = demoManifest.patterns.find((pattern) => pattern.id === 'text-only');
const textOnlyBody = textOnly?.slots.find((slot) => slot.name === 'body') as TextSlotDef | undefined;

test('the words of a slide are a whole number from 1 to 3000, and a description loses what it says of length', () => {
  assert.equal(readWords(200), 200);
  assert.equal(readWords('80'), 80);
  assert.equal(readWords(79.6), 80);
  assert.equal(readWords(' 120 '), 120);
  // More than a slide ever takes is a slip, and so is what no person writes for a number of words.
  for (const none of [0, 0.6, -3, 5000, Number.NaN, Number.POSITIVE_INFINITY, '', 'many', '1e9', '0x20', '12.5', null, undefined, {}, [200]]) assert.equal(readWords(none), undefined, String(none));
  assert.equal(charsOfWords(200), 200 * WORD_CHARS);
  // Hebrew words, with their prefixes, take more: 200 of them are about 1300 characters.
  assert.equal(charsOfWords(200, 'he'), 1300);
  assert.equal(charsOfWords(200, 'he-IL'), 1300);
  assert.equal(charsOfWords(200, 'hr'), 1200, 'only Hebrew');

  assert.equal(withoutLength('About 420 characters in all: four or five paragraphs of full sentences. In Hebrew, open each paragraph with a Hebrew word'), 'Four or five paragraphs of full sentences. In Hebrew, open each paragraph with a Hebrew word');
  assert.equal(withoutLength('180 to 260 characters: two to four full sentences'), 'Two to four full sentences');
  assert.equal(withoutLength('One short point in each item. Write 260 to 500 characters in all'), 'One short point in each item');
  assert.equal(withoutLength('Write 600 to 1130 characters'), '');
  assert.equal(withoutLength('Between 100 and 200 characters: the case'), 'The case');
  // A length inside a sentence is part of what it says, and stays.
  for (const kept of ['A paragraph, about 300 characters, in plain words', 'Two to four sentences of 15 to 25 words each', '50 to 70 characters, or more when there are fewer of them']) assert.equal(withoutLength(kept), kept);
  assert.equal(withoutLength('One line'), 'One line', 'a description that says nothing of length stays');
  assert.equal(withoutLength(undefined), '');
});

test('a slide whose layout holds the text asked keeps it, and the writer is told the length first', async () => {
  // 45 words are 270 characters: within nine tenths of the 331 the layout holds, which is what an asked length may fill.
  const { llm, requests } = scripted(outlineOf(entry('text-only', { words: 45 })));
  const plan = await planOutline(demo, brief, { llm, ...options });
  assert.deepEqual(plan.slides, [{ pattern: 'text-only', purpose: 'Keep the bike share', words: 45 }]);
  // Every outline is told when a slide has words, and the entry may say so.
  assert.match(requests[0]?.prompt ?? '', /\n- When the brief asks for an amount of text on a slide \("200 words", "a paragraph of about 80 words"\), set "words" on that slide/);

  const writer = scripted({ s1: { title: 'Keep it', body: textOf(280) } });
  const written = await writeDeck(demo, plan, { llm: writer.llm, ...options });
  const body = slotOf(writer.requests[0], 's1', 'body');
  const rest = withoutLength(textOnlyBody?.description);
  assert.ok(body?.description?.startsWith(`Write 45 words, as was asked: 270 to 290 characters in all${rest ? `. ${rest}` : ''}`), String(body?.description));
  // The limits are the slot's own. 50 words would not fit in nine tenths of them, and go to a free slide.
  assert.equal(limitOf(body?.description), textOnlyBody?.maxChars);
  assert.ok(!written.warnings.some((issue) => /words were asked/.test(issue.message)));
});

test('a text that no layout holds goes to a free slide set out for that length, whose budget holds it', async () => {
  const { llm } = scripted(outlineOf(entry('text-only', { words: 200 })));
  const plan = await planOutline(demo, brief, { llm, ...options });
  // A free text alone with asked words is set dense: where it fits at four fifths of the template's
  // size it is set as without the flag, and the writer is told the same range.
  assert.deepEqual(plan.slides, [{ pattern: 'free', purpose: 'Keep the bike share', words: 200, blocks: ['text'], arrange: { dense: true } }]);

  const writer = scripted({ s1: { title: 'Keep it', b1Text: textOf(1150) } });
  const written = await writeDeck(demo, plan, { llm: writer.llm, ...options });
  const text = slotOf(writer.requests[0], 's1', 'b1Text');
  // From the length to 15% over it: the box holds more than that.
  assert.match(String(text?.description), /^Write 200 words, as was asked: 1200 to 1370 characters in all\. In paragraphs of two to four sentences, one string each\. at most \d+ characters/);
  assert.ok(limitOf(text?.description) >= 1200, String(text?.description));
  assert.ok(!written.warnings.some((issue) => /words were asked/.test(issue.message)), JSON.stringify(written.warnings));
  assert.equal(written.deck.slides[0]?.pattern, 'free');

  // The same slide planned for a usual text states much less: a text set large on a short line.
  const context: CatalogContext = { layout: 'adaptive', fields, icons: [], assets: [], imageDensity: 'none', figures: false, textScale: 1 };
  const usual = freeWritePattern(demo, { direction: 'ltr' }, { blocks: ['text'] }, context);
  const asked = freeWritePattern(demo, { direction: 'ltr' }, { blocks: ['text'], lengths: { b1Text: 1200 } }, context);
  assert.ok('pattern' in usual && 'pattern' in asked);
  const most = (written: typeof usual) => ('pattern' in written ? limitOf(slotNote(written.pattern.slots.find((planned) => planned.slot.name === 'b1Text')?.slot as never, context)) : 0);
  assert.ok(most(usual) < 1200 && most(asked) >= 1200, `${most(usual)} and ${most(asked)}`);
});

test('a slide of plain text moves to the smallest layout of plain text that holds the text, before a free slide', async () => {
  // A second layout of plain text that holds more: the text-only slide with a body of 900 characters.
  const long = structuredClone(textOnly) as NonNullable<typeof textOnly>;
  long.id = 'text-long';
  long.description = 'A title over a long text.';
  for (const slot of long.slots) if (slot.name === 'body' && slot.kind === 'text') slot.maxChars = 900;
  const roomy = loadTemplate(demoPptx, { ...demoManifest, patterns: [...demoManifest.patterns, long] });
  const { llm } = scripted(outlineOf(entry('text-only', { words: 100 })));
  const plan = await planOutline(roomy, brief, { llm, ...options });
  assert.deepEqual(plan.slides, [{ pattern: 'text-long', purpose: 'Keep the bike share', words: 100 }]);
});

test('without free slides a slide holds what its layout holds, and the plan says so before writing', async () => {
  // Up to six slides, and still one: a text asked for a slide stays on it. 1200 characters, and
  // the layout holds 297 of them, nine tenths of its 331.
  const spread = scripted(outlineOf({ pattern: 'text-only', purpose: 'Keep the bike share', words: 200 }));
  const plan = await planOutline(demo, { ...brief, slideCount: { min: 1, max: 6 } }, { llm: spread.llm, ...options, free: false });
  assert.deepEqual(plan.slides, [{ pattern: 'text-only', purpose: 'Keep the bike share', words: 200 }]);
  assert.deepEqual(planTextFit(demo, plan, { ...options, free: false }), [{ slide: 1, words: 200, holds: 49 }]);

  // The slide holds what it can, the writer is told so, and the deck too.
  const one = scripted(outlineOf({ pattern: 'text-only', purpose: 'Keep the bike share', words: 200 }));
  const single = await planOutline(demo, brief, { llm: one.llm, ...options, free: false });
  assert.deepEqual(single.slides, [{ pattern: 'text-only', purpose: 'Keep the bike share', words: 200 }]);
  const writer = scripted({ s1: { title: 'Keep it', body: textOf(320) } });
  const written = await writeDeck(demo, single, { llm: writer.llm, ...options, free: false });
  const body = slotOf(writer.requests[0], 's1', 'body');
  assert.ok(body?.description?.startsWith('Write 49 words, 260 to 290 characters: of the 200 words asked for this slide, the most it holds'), String(body?.description));
  assert.deepEqual(
    written.warnings.filter((issue) => /words were asked/.test(issue.message)),
    [{ severity: 'warning', slide: 1, message: '200 words were asked for this slide, and it holds about 49: ask for the text over more slides, or for a shorter one' }],
  );
});

test('a change to a plan lists the words of a slide and keeps them', async () => {
  const plan: DeckPlan = { version: 1, title: 'Bikes', brief, slides: [{ pattern: 'free', purpose: 'Keep the bike share', blocks: ['text'], words: 200 }] };
  const { llm, requests } = scripted({ summary: 'Done.', language: 'en', title: 'Bikes!', slides: [entry('free', { shows: [{ kind: 'text', count: 1 }], words: 200 })] });
  const changed = await editPlan(demo, plan, 'A better title', { llm, ...options });
  assert.match(requests[0]?.prompt ?? '', /\n1\. free \(a text; 200 words\): Keep the bike share\n/);
  assert.match(requests[0]?.prompt ?? '', /Leave "words" out of every other slide, and keep the words a slide has unless the request changes them\./);
  assert.deepEqual(changed.plan.slides, [{ pattern: 'free', purpose: 'Keep the bike share', blocks: ['text'], words: 200 }]);
});

test('a change to a deck that asks for an amount of text makes the room for it on that slide, or says what it holds', async () => {
  const deck: DeckSpec = { version: 1, language: 'en', fields, title: 'Bikes', slides: [{ pattern: 'text-only', slots: { title: 'Keep it', body: ['Riders rely on it.'] } }] };
  const change = { summary: 'Longer.', language: 'en', title: 'Bikes', slides: [{ from: 1, pattern: 'text-only', change: 'Make the text 200 words', images: false, words: 200, parts: [], shows_main: 'auto', side_by_side: 0 }], text_size: 'normal', fields };

  // Free slides on: the slide becomes a free slide set out for the text, written from what it held.
  const free = scripted(change, { s1: { title: 'Keep it', b1Text: textOf(1150) } });
  const edited = await planEdit(demo, deck, 'Make slide 1 200 words', { llm: free.llm, ...options });
  assert.match(free.requests[0]?.prompt ?? '', /\n- words: When the request asks for an amount of text on a slide/);
  assert.match(String(slotOf(free.requests[1], 's1', 'b1Text')?.description), /^Write 200 words, as was asked: 1200 to 1370 characters in all\./);
  assert.equal(edited.deck.slides.length, 1);
  assert.equal(edited.deck.slides[0]?.pattern, 'free');
  assert.deepEqual(edited.deck.slides[0]?.arrange, { dense: true });
  assert.ok(!edited.warnings.some((issue) => /words/.test(issue.message)), JSON.stringify(edited.warnings));

  // Free slides off: the text stays on its slide, which holds what its layout holds, and the deck is told.
  const one = scripted(change, { s1: { title: 'Keep it', body: textOf(290) } });
  const longer = await planEdit(demo, deck, 'Make slide 1 200 words', { llm: one.llm, ...options, free: false });
  assert.deepEqual(longer.deck.slides.map((slide) => slide.pattern), ['text-only']);
  assert.deepEqual(longer.rewritten, [1]);
  assert.deepEqual(
    longer.warnings.filter((issue) => /words/.test(issue.message)),
    [{ severity: 'warning', slide: 1, message: '200 words were asked for this slide, and it holds about 49: ask for the text over more slides, or for a shorter one' }],
  );
});

test('the plan tool shows the words asked for a slide', async () => {
  const { llm } = scripted(outlineOf(entry('text-only', { words: 200 })));
  let next = 0;
  const tools = presentationTools({ template: demo, llm, store: memoryStore(), fields, layout: 'adaptive', newId: (kind) => `${kind}_${++next}` });
  const planned = await tools.call(TOOL_NAMES.plan, { topic: 'Why the city should keep its bike share', language: 'en', slide_count: 1 });
  const slides = planned.output.slides as Array<Record<string, unknown>>;
  assert.equal(slides[0]?.layout, 'free');
  assert.equal(slides[0]?.words, 200);
});

test('a slide asked for words keeps them on one slide', async () => {
  // 99 words on a layout that takes 49, in a deck that may have six slides: one slide, cut to what it holds.
  const spread = scripted(outlineOf({ pattern: 'text-only', purpose: 'Keep the bike share', words: 99 }));
  const plan = await planOutline(demo, { ...brief, slideCount: { min: 1, max: 6 } }, { llm: spread.llm, ...options, free: false });
  assert.deepEqual(plan.slides, [{ pattern: 'text-only', purpose: 'Keep the bike share', words: 99 }]);
  assert.deepEqual(planTextFit(demo, plan, { ...options, free: false }), [{ slide: 1, words: 99, holds: 49 }]);

  // A slide with a chart keeps its chart, and the text it holds.
  const charted = scripted(outlineOf({ pattern: 'chart-native', purpose: 'Rides by quarter', words: 200 }));
  const one = await planOutline(demo, { ...brief, prompt: 'Rides by quarter in 2025: 12, 18, 23 and 14 thousand.', slideCount: { min: 1, max: 6 } }, { llm: charted.llm, ...options, free: false });
  assert.deepEqual(one.slides, [{ pattern: 'chart-native', purpose: 'Rides by quarter', words: 200 }]);
});

test('a change to a plan does not fit again a slide whose words it leaves as they were', async () => {
  // One slide, as the deck asked, that holds less than its 200 words.
  const plan: DeckPlan = { version: 1, title: 'Bikes', brief, slides: [{ pattern: 'text-only', purpose: 'Keep the bike share', words: 200 }] };
  const title = scripted({ summary: 'Done.', language: 'en', title: 'Bikes!', slides: [{ pattern: 'text-only', purpose: 'Keep the bike share', words: 200 }] });
  const same = await editPlan(demo, plan, 'A better title', { llm: title.llm, ...options, free: false });
  assert.deepEqual(same.plan.slides, plan.slides);
  // Words that change are fitted, on the slide: 300 words stay on one slide, which holds what it can.
  const more = scripted({ summary: 'Done.', language: 'en', title: 'Bikes', slides: [{ pattern: 'text-only', purpose: 'Keep the bike share', words: 300 }] });
  const longer = await editPlan(demo, plan, 'Make it 300 words', { llm: more.llm, ...options, free: false });
  assert.deepEqual(longer.plan.slides, [{ pattern: 'text-only', purpose: 'Keep the bike share', words: 300 }]);
});

test('a change to a deck that asks for a length keeps what the slide shows besides its text', async () => {
  const photo = { asset: 'bikes', alt: 'Bikes at a station' };
  const assets = [{ id: 'bikes', kind: 'photo' as const, description: 'Bikes at a station', data: new Uint8Array([1]) }];
  const deck: DeckSpec = { version: 1, language: 'en', fields, title: 'Bikes', slides: [{ pattern: 'text-and-photo', slots: { title: 'Keep it', body: 'Riders rely on it.', photo } }] };
  const change = { summary: 'Longer.', language: 'en', title: 'Bikes', slides: [{ from: 1, pattern: 'text-and-photo', change: 'Make the text 120 words', images: false, words: 120, parts: [], shows_main: 'auto', side_by_side: 0 }], text_size: 'normal', fields };
  // The slide becomes a free slide of the text and its photo: the photo is handed on, and not asked of the writer.
  const { llm, requests } = scripted(change, { s1: { title: 'Keep it', b1Text: textOf(700) } });
  const edited = await planEdit(demo, deck, 'Make slide 1 120 words', { llm, ...options, assets });
  assert.deepEqual(edited.deck.slides[0]?.blocks, ['text', 'photo']);
  assert.deepEqual(edited.deck.slides[0]?.slots.b2Photo, photo);
  const asked = Object.keys((requests[1]?.schema as unknown as Schema).properties.s1?.properties ?? {});
  assert.ok(asked.includes('b1Text') && !asked.includes('b2Photo'), asked.join(', '));
  assert.ok(!edited.warnings.some((issue) => /words/.test(issue.message)), JSON.stringify(edited.warnings));
});

test('a change to a free slide that names no part writes them all again, and a text far over what it holds is cut on it', async () => {
  // Two cards and no text: the slide gets a text first, and the cards are written again with it, as the change asked.
  const cards = freeSlide({ title: 'Why', blocks: [{ card: { heading: 'Cheap', text: 'One euro a ride.' } }, { card: { heading: 'Fast', text: 'No traffic.' } }] });
  const deck: DeckSpec = { version: 1, language: 'en', fields, title: 'Bikes', slides: [cards] };
  const change = (words: number) => ({ summary: 'Longer.', language: 'en', title: 'Bikes', slides: [{ from: 1, pattern: 'free', change: `Add ${words} words on why`, images: false, words, parts: [], shows_main: 'auto', side_by_side: 0 }], text_size: 'normal', fields });
  const card = { heading: 'Cheap', text: 'One euro a ride.' };
  const written = scripted(change(60), { s1: { title: 'Why', b1Text: textOf(360), b2Heading: card.heading, b2Text: card.text, b3Heading: 'Fast', b3Text: 'No traffic.' } });
  const edited = await planEdit(demo, deck, 'Add 60 words on why', { llm: written.llm, ...options });
  assert.deepEqual(edited.deck.slides[0]?.blocks, ['text', 'card', 'card']);
  const asked = Object.keys((written.requests[1]?.schema as unknown as Schema).properties.s1?.properties ?? {});
  for (const slot of ['b1Text', 'b2Heading', 'b3Heading']) assert.ok(asked.includes(slot), asked.join(', '));

  // A free slide of one text asked for far more than it holds: it stays one slide, set dense, and holds what it can.
  const text = freeSlide({ title: 'Why', blocks: [{ text: ['Riders rely on it.'] }] });
  const answer = { s1: { title: 'Why', b1Text: textOf(5000) } };
  // The same answer again for each round of repair: what it is asked is not what this test is about.
  const long = scripted({ ...change(1500), slides: [{ ...change(1500).slides[0], change: 'Make it 1500 words' }] }, answer, answer, answer);
  const longer = await planEdit(demo, { ...deck, slides: [text] }, 'Make it 1500 words', { llm: long.llm, ...options });
  assert.equal(longer.deck.slides.length, 1);
  assert.deepEqual(longer.deck.slides[0]?.blocks, ['text']);
  assert.deepEqual(longer.deck.slides[0]?.arrange, { dense: true });
  assert.ok(longer.warnings.some((issue) => /^1500 words were asked for this slide, and it holds about \d+ even with its text set small, at 9 pt: the rest is cut\./.test(issue.message)), JSON.stringify(longer.warnings));
});

// --- 0.6.2: a text asked for a slide stays on it, set smaller where it needs it ---

// A right-to-left template on a slide 10 inches wide, whose frame leaves a text a narrow room.
const rtl = loadTemplate(buildFixturePptx(), fixtureManifest);
const rtlBrief: Brief = { prompt: 'השירות בעיר', language: 'he', imageDensity: 'none', slideCount: 1, fields: { closing: 'Thank you' } };
const SENTENCES = { en: 'Riders rely on the service every day, and the city gains from it in ways that are easy to see. ', he: 'הצוות עובד יחד כל שבוע כדי לשפר את השירות לכל מי שגר בעיר ובסביבתה. ' };
/** A text of exactly so many characters in so many paragraphs, as alike as they can be. */
const paragraphs = (chars: number, count: number, sentence = SENTENCES.en): string[] =>
  Array.from({ length: count }, (_, at) => {
    const length = Math.floor(chars / count) + (at < chars % count ? 1 : 0);
    return sentence.repeat(Math.ceil(length / sentence.length) + 1).slice(0, length - 1) + '.';
  });
const textEntry = (more: Record<string, unknown> = {}) => entry('free', { shows: [{ kind: 'text', count: 1 }], ...more });
const asked = (issue: { message: string }) => /words were asked/.test(issue.message);
const wordsOf = (value: unknown): number => (Array.isArray(value) ? value.join(' ') : String(value ?? '')).split(/\s+/).filter(Boolean).length;
const { slideCount: _count, ...openBrief } = brief;

test('two slides of 500 words each are two dense slides', async () => {
  const two = outlineOf(entry('text-only', { words: 500 }), entry('text-only', { purpose: 'Grow the bike share', words: 500 }));
  const { llm } = scripted(two);
  const plan = await planOutline(demo, { ...brief, slideCount: 2 }, { llm, ...options });
  assert.deepEqual(plan.slides, [
    { pattern: 'free', purpose: 'Keep the bike share', words: 500, blocks: ['text'], arrange: { dense: true } },
    { pattern: 'free', purpose: 'Grow the bike share', words: 500, blocks: ['text'], arrange: { dense: true } },
  ]);
  assert.deepEqual(planTextFit(demo, plan, options), [
    { slide: 1, words: 500, holds: 500, size: 10.5, templateSize: 18, dense: true, most: 808 },
    { slide: 2, words: 500, holds: 500, size: 10.5, templateSize: 18, dense: true, most: 808 },
  ]);

  const writer = scripted({ s1: { title: 'Keep it', b1Text: paragraphs(3000, 7) }, s2: { title: 'Grow it', b1Text: paragraphs(3000, 7) } });
  const written = await writeDeck(demo, plan, { llm: writer.llm, ...options });
  for (const slide of ['s1', 's2']) {
    const text = slotOf(writer.requests[0], slide, 'b1Text');
    // The range of 0.6.1 (1.15 times 3000 is 3449.99... in floating point, so 3440), and the budget of twelve paragraphs.
    assert.match(String(text?.description), /^Write 500 words, as was asked: 3000 to 3440 characters in all\. In paragraphs of two to four sentences, one string each\. at most \d+ characters in total, at most 12 items/);
    assert.ok(limitOf(text?.description) >= 3440, String(text?.description));
    assert.equal(text?.maxItems, 12);
  }
  assert.equal(writer.requests.length, 1);
  assert.deepEqual(written.warnings, []);
  assert.deepEqual(written.deck.slides.map((slide) => slide.arrange), [{ dense: true }, { dense: true }]);

  // Without a count the outline has its six slides, and the two texts keep their words, each on its slide.
  const six = scripted(outlineOf(entry('cover'), entry('text-only', { words: 500 }), entry('text-only', { purpose: 'Grow the bike share', words: 500 }), entry('text-only'), entry('text-only'), entry('closing')));
  const open = await planOutline(demo, openBrief, { llm: six.llm, ...options });
  assert.deepEqual(open.slides.map((slide) => [slide.pattern, slide.words ?? null, slide.arrange ?? null]), [
    ['cover', null, null],
    ['free', 500, { dense: true }],
    ['free', 500, { dense: true }],
    ['text-only', null, null],
    ['text-only', null, null],
    ['closing', null, null],
  ]);
});

test('a text below the floor is cut on its slide and said before writing', async () => {
  const { llm } = scripted(outlineOf(entry('text-only', { words: 1000 })));
  const plan = await planOutline(demo, brief, { llm, ...options });
  assert.deepEqual(plan.slides, [{ pattern: 'free', purpose: 'Keep the bike share', words: 1000, blocks: ['text'], arrange: { dense: true } }]);
  // 6000 characters, and the slide holds 4890 of them at 9 pt, the least that reads on a slide 13.33
  // inches wide, in any number of paragraphs up to twelve. `most` is that room found in steps of 50
  // characters: 4850.
  assert.deepEqual(planTextFit(demo, plan, options), [{ slide: 1, words: 1000, holds: 815, size: 9, templateSize: 18, dense: true, most: 808 }]);

  const writer = scripted({ s1: { title: 'Keep it', b1Text: paragraphs(5000, 12) } });
  const written = await writeDeck(demo, plan, { llm: writer.llm, ...options });
  assert.match(String(slotOf(writer.requests[0], 's1', 'b1Text')?.description), /^Write 815 words, 4400 to 4890 characters: of the 1000 words asked for this slide, the most it holds\./);
  assert.deepEqual(written.warnings.filter(asked), [
    { severity: 'warning', slide: 1, message: '1000 words were asked for this slide, and it holds about 815 even with its text set small, at 9 pt: the rest is cut. Ask for the text over more slides, or for a shorter one' },
  ]);
});

test('a change to a plan keeps the slides it has', async () => {
  // Four parts of 250 words, as 0.6.1 spread 1000 words; "keep it to 2 slides" is answered with two of 500.
  const parts: DeckPlan = { version: 1, title: 'Bikes', brief, slides: [1, 2, 3, 4].map((part) => ({ pattern: 'free', purpose: `Keep the bike share (${part}/4)`, blocks: ['text'], words: 250 })) };
  const two = scripted({ summary: 'Done.', language: 'en', title: 'Bikes', slides: [textEntry({ words: 500 }), textEntry({ purpose: 'Grow the bike share', words: 500 })] });
  const kept = await editPlan(demo, parts, 'Keep it to 2 slides', { llm: two.llm, ...options });
  assert.deepEqual(kept.plan.slides, [
    { pattern: 'free', purpose: 'Keep the bike share', blocks: ['text'], words: 500, arrange: { dense: true } },
    { pattern: 'free', purpose: 'Grow the bike share', blocks: ['text'], words: 500, arrange: { dense: true } },
  ]);
  // A plan with words is told that code adds no slide for a text. 250 words are set at 14.5 pt, four fifths of 18: no mark.
  assert.match(two.requests[0]?.prompt ?? '', /\n1\. free \(a text; 250 words\): Keep the bike share \(1\/4\)\n/);
  assert.match(two.requests[0]?.prompt ?? '', /\n- A slide given words holds them all: the software sets its text smaller where they need it \("text at 7 pt"\), so do not add slides to make room for a text\. /);

  // "A better title": the answer names the slides' parts again, which sets out their arrangement anew, and they stay dense.
  const title = scripted({ summary: 'Done.', language: 'en', title: 'Bikes!', slides: kept.plan.slides.map((slide) => textEntry({ purpose: slide.purpose, words: 500 })) });
  const renamed = await editPlan(demo, kept.plan, 'A better title', { llm: title.llm, ...options });
  assert.deepEqual(renamed.plan.slides.map((slide) => slide.arrange), [{ dense: true }, { dense: true }]);
  // Each line says the size its text is set at in the plan and the most the slide holds.
  assert.match(title.requests[0]?.prompt ?? '', /\n1\. free \(a text; 500 words; text at 10\.5 pt; at most about 800 words\): Keep the bike share\n/);

  // On the right-to-left fixture, whose frame leaves a text a narrow room, 500 Hebrew words are more than the slide holds at 7 pt.
  const narrow: DeckPlan = { version: 1, title: 'x', brief: rtlBrief, slides: [{ pattern: 'free', purpose: 'השירות', blocks: ['text'], arrange: { dense: true }, words: 500 }] };
  const same = scripted({ summary: 'בוצע', language: 'he', title: 'y', slides: [textEntry({ purpose: 'השירות', words: 500 })] });
  const kept2 = await editPlan(rtl, narrow, 'כותרת טובה יותר', { llm: same.llm, layout: 'adaptive' });
  assert.match(same.requests[0]?.prompt ?? '', /\n1\. free \(a text; 500 words; cut to fit; text at 7 pt; at most about 270 words\): השירות\n/);
  assert.deepEqual(kept2.plan.slides, narrow.slides);

  // A plan without words is shown and told as before.
  const plain: DeckPlan = { version: 1, title: 'Bikes', brief, slides: [{ pattern: 'free', purpose: 'Keep the bike share', blocks: ['text'] }] };
  const none = scripted({ summary: 'Done.', language: 'en', title: 'Bikes!', slides: [textEntry()] });
  await editPlan(demo, plain, 'A better title', { llm: none.llm, ...options });
  assert.match(none.requests[0]?.prompt ?? '', /\n1\. free \(a text\): Keep the bike share\n/);
  assert.ok(!(none.requests[0]?.prompt ?? '').includes('A slide given words holds them all'));
});

test('a change to a deck sets a longer text dense', async () => {
  const deck: DeckSpec = { version: 1, language: 'en', fields, title: 'Bikes', slides: [{ pattern: 'text-only', slots: { title: 'Keep it', body: ['Riders rely on it.'] } }] };
  const plan = (slide: Record<string, unknown>) => ({ summary: 'Done.', language: 'en', title: 'Bikes', slides: [{ from: 1, images: false, shows_main: 'auto', side_by_side: 0, ...slide }], text_size: 'normal', fields });
  const longer = scripted(plan({ pattern: 'text-only', change: 'Make the text 500 words', words: 500, parts: [] }), { s1: { title: 'Keep it', b1Text: paragraphs(3100, 8) } });
  const edited = await planEdit(demo, deck, 'Make slide 1 500 words', { llm: longer.llm, ...options });
  assert.equal(edited.deck.slides.length, 1);
  assert.deepEqual(edited.deck.slides[0]?.arrange, { dense: true });
  assert.deepEqual(edited.rewritten, [1]);
  assert.deepEqual(edited.warnings, []);

  // "Simpler words": the plan gives no words and writes the text again. The slide keeps its length.
  const dense = edited.deck;
  // Its length is its characters, at six a word: counted between spaces, this text of short words
  // has 652 words, which would ask for more characters than it has.
  const had = Math.floor((dense.slides[0]?.slots.b1Text as string[]).join('').length / WORD_CHARS);
  assert.deepEqual([had, wordsOf(dense.slides[0]?.slots.b1Text)], [516, 652]);
  const simpler = scripted(plan({ pattern: 'free', change: 'Simpler words', parts: [{ from: 1, kind: 'text', rewrite: true }] }), { s1: { title: 'Keep it', b1Text: paragraphs(charsOfWords(had), 10) } });
  const plain = await planEdit(demo, dense, 'Simpler words on slide 1', { llm: simpler.llm, ...options });
  // The plan is shown the words of the slide, the size its text is set at and the most it holds, and how to keep them.
  assert.match(simpler.requests[0]?.prompt ?? '', new RegExp(`\\n1\\. \\[free: a text; ${had} words; text at [\\d.]+ pt; at most about 800 words\\] title: `));
  assert.match(simpler.requests[0]?.prompt ?? '', /\n- A slide marked "at most about N words" carries a text that was asked for, set smaller where it needs it\. Keep that text whole when the slide is written again: leave "words" out to keep its length\. /);
  assert.match(String(slotOf(simpler.requests[1], 's1', 'b1Text')?.description), new RegExp(`^Write ${had} words, as was asked: `));
  assert.deepEqual(plain.deck.slides[0]?.arrange, { dense: true });
  assert.deepEqual(plain.warnings, []);

  // "A better title": the text is named as kept. The writer is asked for the title alone, and the text stays as it was.
  const title = scripted(plan({ pattern: 'free', change: 'A better title', parts: [{ from: 1, kind: 'text', rewrite: false }] }), { s1: { title: 'Keep the bikes' } });
  const renamed = await planEdit(demo, dense, 'A better title on slide 1', { llm: title.llm, ...options });
  assert.deepEqual(Object.keys((title.requests[1]?.schema as unknown as Schema).properties.s1?.properties ?? {}), ['title']);
  assert.deepEqual(renamed.deck.slides[0]?.slots.b1Text, dense.slides[0]?.slots.b1Text);
  assert.deepEqual(renamed.deck.slides[0]?.arrange, { dense: true });
});

test("a photo gives way to an asked text, and a plan's texts become one; a list, a deck's second text and a chart do not", async () => {
  const photo = { asset: 'bikes', alt: 'Bikes at a station' };
  const assets = [{ id: 'bikes', kind: 'photo' as const, description: 'Bikes at a station', data: new Uint8Array([1]) }];
  const change = (pattern: string) => ({ summary: 'Longer.', language: 'en', title: 'Bikes', slides: [{ from: 1, pattern, change: 'Make the text 500 words', images: false, words: 500, parts: [], shows_main: 'auto', side_by_side: 0 }], text_size: 'normal', fields });
  const left = { severity: 'warning', slide: 1, message: 'the photo of this slide is left out to make room for the 500 words asked for it' };
  // A free slide of a text and a photo: the text alone, set dense, and the deck is told of the photo.
  const free = freeSlide({ title: 'Keep it', blocks: [{ text: ['Riders rely on it.'] }, { photo }] });
  const one = scripted(change('free'), { s1: { title: 'Keep it', b1Text: paragraphs(3100, 8) } });
  const edited = await planEdit(demo, { version: 1, language: 'en', fields, title: 'Bikes', slides: [free] }, 'Make slide 1 500 words', { llm: one.llm, ...options, assets });
  assert.deepEqual([edited.deck.slides[0]?.blocks, edited.deck.slides[0]?.arrange], [['text'], { dense: true }]);
  assert.deepEqual(edited.warnings, [left]);
  // A slide of a layout with a photo: the same.
  const layout: DeckSpec = { version: 1, language: 'en', fields, title: 'Bikes', slides: [{ pattern: 'text-and-photo', slots: { title: 'Keep it', body: 'Riders rely on it.', photo } }] };
  const two = scripted(change('text-and-photo'), { s1: { title: 'Keep it', b1Text: paragraphs(3100, 8) } });
  const relaid = await planEdit(demo, layout, 'Make slide 1 500 words', { llm: two.llm, ...options, assets });
  assert.deepEqual([relaid.deck.slides[0]?.blocks, relaid.deck.slides[0]?.arrange], [['text'], { dense: true }]);
  assert.deepEqual(relaid.warnings, [left]);

  // A list: its part stays, and the slide holds what it can of the words, which the plan says.
  const listed = scripted(outlineOf(entry('free', { shows: [{ kind: 'list', count: 1 }], words: 700 })));
  const plan = await planOutline(demo, brief, { llm: listed.llm, ...options });
  assert.deepEqual([plan.slides[0]?.blocks, plan.slides[0]?.arrange], [['list'], undefined]);
  const [listFit] = planTextFit(demo, plan, options);
  assert.ok(listFit && listFit.holds < 700 && listFit.dense === undefined, JSON.stringify(listFit));
  const items = { s1: { title: 'Keep it', b1Items: ['Riders rely on it.', 'The city gains.'] } };
  const writer = scripted(items, items, items);
  const written = await writeDeck(demo, plan, { llm: writer.llm, ...options });
  assert.deepEqual(written.warnings.filter(asked), [{ severity: 'warning', slide: 1, message: `700 words were asked for this slide, and it holds about ${listFit?.holds}: ask for the text over more slides, or for a shorter one` }]);
  // Two texts of a plan, which nothing has written yet: where they would be cut, they are one text, set dense, which holds the words.
  const texts = scripted(outlineOf(entry('free', { shows: [{ kind: 'text', count: 2 }], words: 700 })));
  const joined = await planOutline(demo, brief, { llm: texts.llm, ...options });
  assert.deepEqual([joined.slides[0]?.blocks, joined.slides[0]?.arrange], [['text'], { dense: true }]);
  assert.deepEqual(
    planTextFit(demo, joined, options).map((fit) => [fit.words, fit.holds]),
    [[700, 700]],
  );
  // Where they hold the words as they are, they stay.
  const fewer = scripted(outlineOf(entry('free', { shows: [{ kind: 'text', count: 2 }], words: 150 })));
  assert.deepEqual((await planOutline(demo, brief, { llm: fewer.llm, ...options })).slides[0]?.blocks, ['text', 'text']);
  // A change to a deck keeps a second text the slide has: it holds what it can.
  const pair = freeSlide({ title: 'Keep it', blocks: [{ text: ['Riders rely on it.'] }, { text: ['The city gains.'] }] });
  const both = { s1: { title: 'Keep it', b1Text: paragraphs(1500, 4), b2Text: paragraphs(1500, 4) } };
  const keeps = scripted(change('free'), both, both, both);
  const separate = await planEdit(demo, { version: 1, language: 'en', fields, title: 'Bikes', slides: [pair] }, 'Make slide 1 500 words', { llm: keeps.llm, ...options });
  assert.deepEqual(separate.deck.slides[0]?.blocks, ['text', 'text']);
  // A chart keeps its chart.
  const charted = scripted(outlineOf({ pattern: 'chart-native', purpose: 'Rides by quarter', words: 500 }));
  const chart = await planOutline(demo, { ...brief, prompt: 'Rides by quarter in 2025: 12, 18, 23 and 14 thousand.' }, { llm: charted.llm, ...options });
  const kept = chart.slides[0];
  assert.ok(kept?.pattern === 'chart-native' || kept?.blocks?.includes('chart'), JSON.stringify(kept));
  assert.equal(kept?.arrange?.dense, undefined);
  const [fit] = planTextFit(demo, chart, options);
  assert.ok(fit && fit.holds < 500, JSON.stringify(fit));
});

test('a free text with asked words validates at every number of paragraphs', async () => {
  // The writer is told up to twelve paragraphs and a little over the length. Whatever number it
  // writes, a text of 1.04 times the length it was asked for stays inside its room.
  const cases = [
    { template: rtl, base: rtlBrief, language: 'he', sentence: SENTENCES.he, all: [200, 230, 250, 270] },
    { template: demo, base: brief, language: 'en', sentence: SENTENCES.en, all: [150, 200, 225] },
  ];
  for (const { template, base, language, sentence, all } of cases) {
    for (const words of all) {
      const { llm } = scripted(outlineOf(textEntry({ words })));
      const plan = await planOutline(template, base, { llm, ...options });
      assert.equal(plan.slides[0]?.arrange?.dense, true);
      for (const count of [6, 8, 10, 12]) {
        const writer = scripted({ s1: { title: 'A title', b1Text: paragraphs(Math.round(1.04 * charsOfWords(words, language)), count, sentence) } });
        const written = await writeDeck(template, plan, { llm: writer.llm, ...options, maxRepairs: 0 });
        assert.deepEqual(written.warnings.filter((issue) => /characters|lines|paragraphs/.test(issue.message)), [], `${language} ${words} words in ${count} paragraphs`);
      }
    }
  }
});

test('a text written much shorter than asked is sent back once', async () => {
  const { llm } = scripted(outlineOf(entry('text-only', { words: 500 })));
  const plan = await planOutline(demo, brief, { llm, ...options });
  const short = { s1: { title: 'Keep it', b1Text: paragraphs(1800, 6) } };
  const whole = { s1: { title: 'Keep it', b1Text: paragraphs(3100, 8) } };
  const once = scripted(short, whole);
  const written = await writeDeck(demo, plan, { llm: once.llm, ...options });
  assert.equal(once.requests.length, 2);
  assert.match(once.requests[1]?.prompt ?? '', /\nIt has these problems:\n- s1\.b1Text: is 1800 characters, and 500 words were asked for this slide: write 3000 to 3440 characters, as was asked\n\n/);
  assert.deepEqual(written.warnings, []);

  // Short again: it is not sent back a second time, and the deck is told.
  const twice = scripted(short, short, short);
  const still = await writeDeck(demo, plan, { llm: twice.llm, ...options });
  assert.equal(twice.requests.length, 2);
  assert.deepEqual(still.warnings, [{ severity: 'warning', slide: 1, message: '500 words were asked for this slide, and it has about 300: ask for it again to have them all' }]);
});

test('a change to a deck that would cut a text stops before writing', async () => {
  const deck: DeckSpec = { version: 1, language: 'en', fields, title: 'Bikes', slides: [{ pattern: 'text-only', slots: { title: 'Keep it', body: ['Riders rely on it.'] } }] };
  const change = { summary: 'Longer.', language: 'en', title: 'Bikes', slides: [{ from: 1, pattern: 'text-only', change: 'Make the text 1500 words', images: false, words: 1500, parts: [], shows_main: 'auto', side_by_side: 0 }], text_size: 'normal', fields };
  const stop = scripted(change);
  await assert.rejects(planEdit(demo, deck, 'Make slide 1 1500 words', { llm: stop.llm, ...options, stopBeforeCut: true }), (error: unknown) => {
    assert.ok(error instanceof TextCutError);
    assert.equal(error.message, 'Nothing was changed: slide 1 would hold about 815 of its 1500 words (text at 9 pt). Ask for more slides or a shorter text, or accept that the rest is cut.');
    assert.deepEqual(error.fits, [{ slide: 1, words: 1500, holds: 815, size: 9, dense: true }]);
    return true;
  });
  assert.equal(stop.requests.length, 1);

  // Without the option the slide is written, cut to what it holds, and the deck is told.
  const written = scripted(change, { s1: { title: 'Keep it', b1Text: paragraphs(5000, 12) } });
  const edited = await planEdit(demo, deck, 'Make slide 1 1500 words', { llm: written.llm, ...options });
  assert.deepEqual(edited.warnings.filter(asked), [
    { severity: 'warning', slide: 1, message: '1500 words were asked for this slide, and it holds about 815 even with its text set small, at 9 pt: the rest is cut. Ask for the text over more slides, or for a shorter one' },
  ]);
});

/** What a change that stops before writing says of its slides, or `undefined` where it did not stop. */
const stopped = (edit: Promise<unknown>): Promise<SlideTextFit[] | undefined> =>
  edit.then(
    () => undefined,
    (error: unknown) => {
      if (error instanceof TextCutError) return error.fits;
      throw error;
    },
  );

test('a change stops before writing for the room its slide is written in, a heading it keeps counted', async () => {
  const deck = (heading: boolean): DeckSpec => ({
    version: 1,
    language: 'en',
    fields,
    title: 'Bikes',
    slides: [{ pattern: 'free', blocks: ['text'], arrange: { dense: true }, slots: { title: 'Keep it', ...(heading ? { b1Heading: 'An overview' } : {}), b1Text: paragraphs(1800, 6) } }],
  });
  const change = { summary: 'Longer.', language: 'en', title: 'Bikes', slides: [{ from: 1, pattern: 'free', change: 'Make the text 780 words', images: false, words: 780, parts: [], shows_main: 'auto', side_by_side: 0 }], text_size: 'normal', fields };
  const write = (heading: boolean) => ({ s1: { title: 'Keep it', ...(heading ? { b1Heading: 'An overview' } : {}), b1Text: paragraphs(4600, 12) } });
  // A text alone holds 780 words, set small: nothing stops.
  assert.equal(await stopped(planEdit(demo, deck(false), 'Make slide 1 780 words', { llm: scripted(change, write(false), write(false)).llm, ...options, stopBeforeCut: true })), undefined);
  // The heading the slide keeps takes some of that room. The change stops after its plan, with what
  // the slide holds as it would be written.
  const stop = scripted(change);
  const fits = await stopped(planEdit(demo, deck(true), 'Make slide 1 780 words', { llm: stop.llm, ...options, stopBeforeCut: true }));
  assert.equal(stop.requests.length, 1);
  assert.equal(fits?.length, 1);
  const [fit] = fits ?? [];
  assert.deepEqual({ ...fit, holds: 0 }, { slide: 1, words: 780, holds: 0, size: 9, dense: true });
  assert.ok((fit?.holds ?? 780) < 780, JSON.stringify(fit));
  // Written all the same, it holds as many as the stop said.
  const cut = await planEdit(demo, deck(true), 'Make slide 1 780 words', { llm: scripted(change, write(true), write(true), write(true)).llm, ...options });
  assert.deepEqual(cut.warnings.filter(asked), [
    { severity: 'warning', slide: 1, message: `780 words were asked for this slide, and it holds about ${fit?.holds} even with its text set small, at 9 pt: the rest is cut. Ask for the text over more slides, or for a shorter one` },
  ]);
});

test('a part added beside a dense text leaves it less room: the text is written for that room, and the cut is said before writing', async () => {
  // 1800 characters: the length of 300 words, set dense on a slide of its own.
  const deck: DeckSpec = { version: 1, language: 'en', fields, title: 'Bikes', slides: [{ pattern: 'free', blocks: ['text'], arrange: { dense: true }, slots: { title: 'Keep it', b1Text: paragraphs(1800, 8) } }] };
  const parts = [
    { from: 1, kind: 'text', rewrite: false },
    { from: 0, kind: 'list', rewrite: true },
  ];
  const change = { summary: 'Added.', language: 'en', title: 'Bikes', slides: [{ from: 1, pattern: 'free', change: 'Add three key points beside the text', images: false, parts, shows_main: 'auto', side_by_side: 0 }], text_size: 'normal', fields };
  const stop = scripted(change);
  const fits = await stopped(planEdit(demo, deck, 'Add three key points beside the text of slide 1', { llm: stop.llm, ...options, stopBeforeCut: true }));
  assert.equal(stop.requests.length, 1);
  const [fit] = fits ?? [];
  assert.deepEqual({ ...fit, holds: 0 }, { slide: 1, words: 300, holds: 0 });
  assert.ok((fit?.holds ?? 300) < 300, JSON.stringify(fit));
  // Without the stop, the text is written again for the room it has beside the list, at the
  // template's size, and the deck is told. Kept as it was, it ran out of its box.
  const written = { s1: { title: 'Keep it', b1Text: paragraphs(500, 3), b2Items: ['Riders rely on it.', 'The city gains.', 'It costs little.'] } };
  const model = scripted(change, written, written, written);
  const edited = await planEdit(demo, deck, 'Add three key points beside the text of slide 1', { llm: model.llm, ...options });
  assert.match(String(slotOf(model.requests[1], 's1', 'b1Text')?.description), /^Write \d+ words, \d+ to \d+ characters: of the 300 words asked for this slide, the most it holds/);
  assert.deepEqual([edited.deck.slides[0]?.blocks, edited.deck.slides[0]?.arrange], [['text', 'list'], undefined]);
  assert.deepEqual(edited.warnings, [{ severity: 'warning', slide: 1, message: `300 words were asked for this slide, and it holds about ${fit?.holds}: ask for the text over more slides, or for a shorter one` }]);
});

test('a dense text that still holds beside a part added to it stays as it is, and a photo it leaves no room for gives way', async () => {
  const deck = (chars: number): DeckSpec => ({ version: 1, language: 'en', fields, title: 'Bikes', slides: [{ pattern: 'free', blocks: ['text'], arrange: { dense: true }, slots: { title: 'Keep it', b1Text: paragraphs(chars, 3) } }] });
  const what = { list: 'three key points', photo: 'a photo of bikes', text: 'a short text on costs' };
  const change = (kind: keyof typeof what) => ({
    summary: 'Added.',
    language: 'en',
    title: 'Bikes',
    slides: [{ from: 1, pattern: 'free', change: `Add ${what[kind]}`, images: kind === 'photo', parts: [{ from: 1, kind: 'text', rewrite: false }, { from: 0, kind, rewrite: true }], shows_main: 'auto', side_by_side: 0 }],
    text_size: 'normal',
    fields,
  });
  const assets = [{ id: 'bikes', kind: 'photo' as const, description: 'Bikes at a station', data: new Uint8Array([1]) }];
  const writer = (request: JsonRequest) => {
    const asked = Object.keys((request.schema as unknown as Schema).properties.s1?.properties ?? {});
    return { s1: Object.fromEntries(asked.map((name) => [name, name === 'title' ? 'Keep it' : name.endsWith('Items') ? ['Riders rely on it.', 'The city gains.', 'It costs little.'] : name.endsWith('Photo') ? { asset: 'bikes', alt: 'Bikes' } : ['Written again.']])) };
  };
  // A text of 60 words, set dense when it was asked for: beside a list, a photo or a second text it
  // still holds. It stays as it is, and nothing of it is asked of the writer. A new text beside it
  // is written as the request says, not given half of its length.
  for (const kind of ['list', 'photo', 'text'] as const) {
    const model = scripted(change(kind), writer, writer);
    const edited = await planEdit(demo, deck(330), `Add ${what[kind]} to slide 1`, { llm: model.llm, ...options, assets, stopBeforeCut: true });
    const slide = edited.deck.slides[0];
    assert.deepEqual([slide?.blocks, slide?.arrange, slide?.slots.b1Text], [['text', kind], undefined, deck(330).slides[0]?.slots.b1Text], kind);
    assert.equal(slotOf(model.requests[1], 's1', 'b1Text'), undefined, kind);
    if (kind === 'text') assert.doesNotMatch(String(slotOf(model.requests[1], 's1', 'b2Text')?.description), /words/);
    assert.deepEqual(edited.warnings, [], kind);
  }
  // 300 words leave a photo no room: it gives way to the text, which stays as it is. Nothing was
  // asked of the slide's length now, and the deck is told what it was left out for.
  const model = scripted(change('photo'), writer, writer);
  const edited = await planEdit(demo, deck(1800), 'Add a photo of bikes to slide 1', { llm: model.llm, ...options, assets, stopBeforeCut: true });
  const slide = edited.deck.slides[0];
  assert.deepEqual([slide?.blocks, slide?.arrange, slide?.slots.b1Text], [['text'], { dense: true }, deck(1800).slides[0]?.slots.b1Text]);
  assert.equal(slotOf(model.requests[1], 's1', 'b1Text'), undefined);
  assert.deepEqual(edited.warnings, [{ severity: 'warning', slide: 1, message: 'the photo of this slide is left out to make room for its text of about 300 words' }]);
});

test('a dense text written again with no length asked keeps the length it has, and is never stopped as a cut', async () => {
  const plan = (slide: Record<string, unknown>) => ({ summary: 'Done.', language: 'en', title: 'Bikes', slides: [{ from: 1, images: false, shows_main: 'auto', side_by_side: 0, pattern: 'free', ...slide }], text_size: 'normal', fields });
  const changes = [
    ['Simpler words on slide 1', { change: 'Simpler words', parts: [{ from: 1, kind: 'text', rewrite: true }] }],
    ['A better title on slide 1', { change: 'A better title', parts: [] }],
  ] as const;
  // Texts of short words that the slide holds: counted between spaces, their words at six characters
  // each would be more than the slide holds. The longer one is over what a text written to a length
  // may fill, and is written to that.
  for (const chars of [4800, 5100]) {
    const text = paragraphs(chars, 9);
    assert.ok(wordsOf(text) * WORD_CHARS > 5300, `${chars}`);
    const deck: DeckSpec = { version: 1, language: 'en', fields, title: 'Bikes', slides: [{ pattern: 'free', blocks: ['text'], arrange: { dense: true }, slots: { title: 'Keep it', b1Text: text } }] };
    for (const [instruction, slide] of changes) {
      const model = scripted(plan(slide), { s1: { title: 'Keep the bikes', b1Text: text } }, { s1: { title: 'Keep the bikes', b1Text: text } });
      const edited = await planEdit(demo, deck, instruction, { llm: model.llm, ...options, stopBeforeCut: true });
      const words = Number(/^Write (\d+) words, as was asked: /.exec(String(slotOf(model.requests[1], 's1', 'b1Text')?.description))?.[1]);
      assert.ok(words > 0 && words <= Math.floor(chars / WORD_CHARS), `${chars} ${instruction}: ${words}`);
      if (chars === 4800) assert.equal(words, 800);
      assert.deepEqual(edited.deck.slides[0]?.arrange, { dense: true });
      assert.deepEqual(edited.warnings.filter(asked), [], `${chars} ${instruction}`);
    }
  }
});

test('a writer inside the range it is told is not sent back, however many paragraphs it writes', async () => {
  // 800 words: 4800 characters on one slide, held at 9 pt, the least that reads on a slide 13.33 inches wide.
  const { llm } = scripted(outlineOf(entry('text-only', { words: 800 })));
  const plan = await planOutline(demo, brief, { llm, ...options });
  assert.deepEqual(planTextFit(demo, plan, options), [{ slide: 1, words: 800, holds: 800, size: 9, templateSize: 18, dense: true, most: 808 }]);
  const first = scripted({ s1: { title: 'Keep it', b1Text: paragraphs(4800, 12) } });
  await writeDeck(demo, plan, { llm: first.llm, ...options });
  const description = String(slotOf(first.requests[0], 's1', 'b1Text')?.description);
  // A budget of what 12 paragraphs hold told it up to 5220 characters, and 5220 in 10 paragraphs
  // were sent back to be shortened. The range now ends where every number of paragraphs fits.
  assert.match(description, /^Write 800 words, as was asked: 4640 to 4890 characters in all\./);
  assert.equal(limitOf(description), 4890);
  for (let count = 1; count <= 12; count++) {
    const writer = scripted({ s1: { title: 'Keep it', b1Text: paragraphs(4890, count) } });
    const written = await writeDeck(demo, plan, { llm: writer.llm, ...options });
    assert.equal(writer.requests.length, 1, `${count} paragraphs`);
    assert.deepEqual(written.warnings, [], `${count} paragraphs`);
  }
});
