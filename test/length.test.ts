// An amount of text a person asks for on a slide ("200 words"): the plan reads it, gives the text a
// slide that holds it, and the writer is told the length before what a slot says of length. No test
// calls a live model.
import assert from 'node:assert/strict';
import { readFileSync } from 'node:fs';
import { dirname, join } from 'node:path';
import { test } from 'node:test';
import { fileURLToPath } from 'node:url';
import { TOOL_NAMES, editPlan, freeSlide, loadTemplate, memoryStore, planEdit, planOutline, presentationTools, writeDeck } from '../src/index.ts';
import type { Brief, DeckPlan, DeckSpec, JsonRequest, LlmProvider, TemplateManifest } from '../src/index.ts';
import { slotNote } from '../src/plan/catalog.ts';
import type { CatalogContext } from '../src/plan/catalog.ts';
import { freeWritePattern } from '../src/plan/free.ts';
import { WORD_CHARS, charsOfWords, readWords, withoutLength } from '../src/plan/length.ts';
import type { TextSlotDef } from '../src/template/manifest.ts';

const root = join(dirname(fileURLToPath(import.meta.url)), '..');
const demoManifest = JSON.parse(readFileSync(join(root, 'templates', 'demo.manifest.json'), 'utf8')) as TemplateManifest;
const demoPptx = new Uint8Array(readFileSync(join(root, 'templates', 'demo.pptx')));
const demo = loadTemplate(demoPptx, demoManifest);
const fields = { organizationDepartment: 'Retail' };
const brief: Brief = { prompt: 'Why the city should keep its bike share', language: 'en', fields, imageDensity: 'none', slideCount: 1 };
const options = { layout: 'adaptive' as const, fields };

/** A model that replays prepared answers and records what it was asked. */
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
  assert.deepEqual(plan.slides, [{ pattern: 'free', purpose: 'Keep the bike share', words: 200, blocks: ['text'] }]);

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

test('without free slides the text goes over the slides the deck may have, and a slide that holds less is told', async () => {
  // Up to six slides: 1200 characters take five of the 297 that nine tenths of the layout's 331 hold.
  const spread = scripted(outlineOf({ pattern: 'text-only', purpose: 'Keep the bike share', words: 200 }));
  const plan = await planOutline(demo, { ...brief, slideCount: { min: 1, max: 6 } }, { llm: spread.llm, ...options, free: false });
  assert.deepEqual(
    plan.slides,
    [1, 2, 3, 4, 5].map((part) => ({ pattern: 'text-only', purpose: `Keep the bike share (${part}/5)`, words: 40 })),
  );

  // One slide: the slide holds what it can, the writer is told so, and the deck too.
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

test('a change to a deck that asks for an amount of text makes the room for it, or says where the text went', async () => {
  const deck: DeckSpec = { version: 1, language: 'en', fields, title: 'Bikes', slides: [{ pattern: 'text-only', slots: { title: 'Keep it', body: ['Riders rely on it.'] } }] };
  const change = { summary: 'Longer.', language: 'en', title: 'Bikes', slides: [{ from: 1, pattern: 'text-only', change: 'Make the text 200 words', images: false, words: 200, parts: [], shows_main: 'auto', side_by_side: 0 }], text_size: 'normal', fields };

  // Free slides on: the slide becomes a free slide set out for the text, written from what it held.
  const free = scripted(change, { s1: { title: 'Keep it', b1Text: textOf(1150) } });
  const edited = await planEdit(demo, deck, 'Make slide 1 200 words', { llm: free.llm, ...options });
  assert.match(free.requests[0]?.prompt ?? '', /\n- words: When the request asks for an amount of text on a slide/);
  assert.match(String(slotOf(free.requests[1], 's1', 'b1Text')?.description), /^Write 200 words, as was asked: 1200 to 1370 characters in all\./);
  assert.equal(edited.deck.slides.length, 1);
  assert.equal(edited.deck.slides[0]?.pattern, 'free');
  assert.ok(!edited.warnings.some((issue) => /words/.test(issue.message)), JSON.stringify(edited.warnings));

  // Free slides off: the text goes over five slides of its layout, four of them new, and the deck is told.
  const five = Array.from({ length: 5 }, (_, at) => [`s${at + 1}`, { title: 'Keep it', body: textOf(250) }]);
  const spread = scripted(change, Object.fromEntries(five));
  const longer = await planEdit(demo, deck, 'Make slide 1 200 words', { llm: spread.llm, ...options, free: false });
  assert.deepEqual(longer.deck.slides.map((slide) => slide.pattern), ['text-only', 'text-only', 'text-only', 'text-only', 'text-only']);
  assert.deepEqual(longer.rewritten, [1, 2, 3, 4, 5]);
  assert.match(promptOfWriter(spread.requests[1]), /s2 \(text-only\)\. New slide: Make the text 200 words \(2\/5\)/);
  assert.deepEqual(
    longer.warnings.filter((issue) => /words/.test(issue.message)),
    [{ severity: 'warning', message: 'the 200 words asked for slide 1 go over slides 1 to 5: one slide holds about 49' }],
  );
});
const promptOfWriter = (request: JsonRequest | undefined): string => request?.prompt ?? '';

test('the plan tool shows the words asked for a slide', async () => {
  const { llm } = scripted(outlineOf(entry('text-only', { words: 200 })));
  let next = 0;
  const tools = presentationTools({ template: demo, llm, store: memoryStore(), fields, layout: 'adaptive', newId: (kind) => `${kind}_${++next}` });
  const planned = await tools.call(TOOL_NAMES.plan, { topic: 'Why the city should keep its bike share', language: 'en', slide_count: 1 });
  const slides = planned.output.slides as Array<Record<string, unknown>>;
  assert.equal(slides[0]?.layout, 'free');
  assert.equal(slides[0]?.words, 200);
});

test('a text goes over more slides in whole words, and only from a slide that shows text alone', async () => {
  // 99 words on a layout that takes 49: three slides of 33, none of which holds less than its share.
  const spread = scripted(outlineOf({ pattern: 'text-only', purpose: 'Keep the bike share', words: 99 }));
  const plan = await planOutline(demo, { ...brief, slideCount: { min: 1, max: 6 } }, { llm: spread.llm, ...options, free: false });
  assert.deepEqual(plan.slides.map((slide) => slide.words), [33, 33, 33]);
  const writer = scripted(Object.fromEntries([1, 2, 3].map((n) => [`s${n}`, { title: 'Keep it', body: textOf(190) }])));
  const written = await writeDeck(demo, plan, { llm: writer.llm, ...options, free: false });
  assert.ok(!written.warnings.some((issue) => /words were asked/.test(issue.message)), JSON.stringify(written.warnings));

  // A slide with a chart is not repeated for its text: every copy would ask for a chart the brief has once.
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
  // Words that change are fitted: 300 words go over the slides they need.
  const more = scripted({ summary: 'Done.', language: 'en', title: 'Bikes', slides: [{ pattern: 'text-only', purpose: 'Keep the bike share', words: 300 }] });
  const longer = await editPlan(demo, plan, 'Make it 300 words', { llm: more.llm, ...options, free: false });
  assert.ok(longer.plan.slides.length > 1, JSON.stringify(longer.plan.slides));
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

test('a change to a free slide that names no part writes them all again, and a text it spreads goes on in free slides of text', async () => {
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

  // A free slide of one text asked for far more than it holds: it goes on in free slides of one text, written new.
  const text = freeSlide({ title: 'Why', blocks: [{ text: ['Riders rely on it.'] }] });
  const answers = Object.fromEntries(Array.from({ length: 11 }, (_, at) => [`s${at + 1}`, { title: 'Why', b1Text: textOf(1500) }]));
  // The same answer again for each round of repair: what it is asked is not what this test is about.
  const long = scripted({ ...change(1500), slides: [{ ...change(1500).slides[0], change: 'Make it 1500 words' }] }, answers, answers, answers);
  const longer = await planEdit(demo, { ...deck, slides: [text] }, 'Make it 1500 words', { llm: long.llm, ...options });
  assert.ok(longer.deck.slides.length > 1);
  for (const slide of longer.deck.slides) assert.deepEqual(slide.blocks, ['text']);
  assert.ok(longer.warnings.some((issue) => new RegExp(`^the 1500 words asked for slide 1 go over slides 1 to ${longer.deck.slides.length}: `).test(issue.message)), JSON.stringify(longer.warnings));
});
