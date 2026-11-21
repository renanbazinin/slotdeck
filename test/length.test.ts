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
import { withoutCodes } from './issues.ts';

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
  // 45 words are 270 characters. The count of the layout's spacing holds 263 in three paragraphs
  // and all 331 in two: the writer chooses its paragraphs, and is told two.
  const { llm, requests } = scripted(outlineOf(entry('text-only', { words: 45 })));
  const plan = await planOutline(demo, brief, { llm, ...options });
  assert.deepEqual(plan.slides, [{ pattern: 'text-only', purpose: 'Keep the bike share', words: 45 }]);
  // Every outline is told when a slide has words, and the entry may say so.
  assert.match(requests[0]?.prompt ?? '', /\n- When the brief asks for an amount of text on a slide \("200 words", "a paragraph of about 80 words"\), set "words" on that slide/);

  const writer = scripted({ s1: { title: 'Keep it', body: textOf(280) } });
  const written = await writeDeck(demo, plan, { llm: writer.llm, ...options });
  const body = slotOf(writer.requests[0], 's1', 'body');
  const rest = withoutLength(textOnlyBody?.description);
  assert.ok(body?.description?.startsWith(`Write 45 words, as was asked: 270 to 310 characters in all${rest ? `. ${rest}` : ''}`), String(body?.description));
  assert.match(String(body?.description), /at most 331 characters in total, at most 2 items$/);
  // The limits are the slot's own: 56 words would not fit in two paragraphs, and go to a free slide.
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
  assert.match(String(text?.description), /^Write 200 words, as was asked: 1200 to 1380 characters in all\. In paragraphs of two to four sentences, one string each\. at most \d+ characters/);
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
  // A second layout of plain text that holds more: the text-only slide with a body of 900 characters,
  // in a box as wide as the slide and set smaller, where its spacing counted holds them all.
  const long = structuredClone(textOnly) as NonNullable<typeof textOnly>;
  long.id = 'text-long';
  long.description = 'A title over a long text.';
  for (const slot of long.slots) {
    if (slot.name !== 'body' || slot.kind !== 'text') continue;
    slot.maxChars = 900;
    slot.fontSize = 14;
    slot.box = { x: 0.0375, y: 0.4933, w: 0.925, h: 0.4667 };
  }
  const roomy = loadTemplate(demoPptx, { ...demoManifest, patterns: [...demoManifest.patterns, long] });
  const { llm } = scripted(outlineOf(entry('text-only', { words: 100 })));
  const plan = await planOutline(roomy, brief, { llm, ...options });
  assert.deepEqual(plan.slides, [{ pattern: 'text-long', purpose: 'Keep the bike share', words: 100 }]);
});

test('without free slides a slide holds what its layout holds, and the plan says so before writing', async () => {
  // Up to six slides, and still one: a text asked for a slide stays on it. 1200 characters, and
  // the layout holds all its 331 of them in two paragraphs, its spacing counted.
  const spread = scripted(outlineOf({ pattern: 'text-only', purpose: 'Keep the bike share', words: 200 }));
  const plan = await planOutline(demo, { ...brief, slideCount: { min: 1, max: 6 } }, { llm: spread.llm, ...options, free: false });
  assert.deepEqual(plan.slides, [{ pattern: 'text-only', purpose: 'Keep the bike share', words: 200 }]);
  assert.deepEqual(planTextFit(demo, plan, { ...options, free: false }), [{ slide: 1, words: 200, holds: 55 }]);

  // The slide holds what it can, the writer is told so, and the deck too, with why its text is not set smaller.
  const one = scripted(outlineOf({ pattern: 'text-only', purpose: 'Keep the bike share', words: 200 }));
  const single = await planOutline(demo, brief, { llm: one.llm, ...options, free: false });
  assert.deepEqual(single.slides, [{ pattern: 'text-only', purpose: 'Keep the bike share', words: 200 }]);
  const writer = scripted({ s1: { title: 'Keep it', body: textOf(320) } });
  const written = await writeDeck(demo, single, { llm: writer.llm, ...options, free: false });
  const body = slotOf(writer.requests[0], 's1', 'body');
  assert.ok(body?.description?.startsWith('Write 55 words, 290 to 330 characters: of the 200 words asked for this slide, the most it holds'), String(body?.description));
  assert.match(String(body?.description), /at most 331 characters in total, at most 2 items$/);
  assert.deepEqual(
    withoutCodes(written.warnings.filter((issue) => /words were asked/.test(issue.message))),
    [{ severity: 'warning', slide: 1, message: '200 words were asked for this slide, and it holds about 55: ask for the text over more slides, or for a shorter one; free slides are off, so its text is not set smaller' }],
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
  assert.match(String(slotOf(free.requests[1], 's1', 'b1Text')?.description), /^Write 200 words, as was asked: 1200 to 1380 characters in all\./);
  assert.equal(edited.deck.slides.length, 1);
  assert.equal(edited.deck.slides[0]?.pattern, 'free');
  assert.deepEqual(edited.deck.slides[0]?.arrange, { dense: true });
  assert.ok(!edited.warnings.some((issue) => /words/.test(issue.message)), JSON.stringify(edited.warnings));

  // Free slides off: the text stays on its slide, which holds what its layout holds, and the deck is told, with why.
  const one = scripted(change, { s1: { title: 'Keep it', body: textOf(290) } });
  const longer = await planEdit(demo, deck, 'Make slide 1 200 words', { llm: one.llm, ...options, free: false });
  assert.deepEqual(longer.deck.slides.map((slide) => slide.pattern), ['text-only']);
  assert.deepEqual(longer.rewritten, [1]);
  assert.deepEqual(
    withoutCodes(longer.warnings.filter((issue) => /words/.test(issue.message))),
    [{ severity: 'warning', slide: 1, message: '200 words were asked for this slide, and it holds about 55: ask for the text over more slides, or for a shorter one; free slides are off, so its text is not set smaller' }],
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

test("with free slides off a cut says why its text is not set smaller: in the plan tool's note, and where a change stops before writing", async () => {
  const toolsOf = (more: { free?: boolean; patterns?: string[] }) => {
    let next = 0;
    const { llm } = scripted(outlineOf({ pattern: 'text-only', purpose: 'Keep the bike share', words: 200 }));
    return presentationTools({ template: demo, llm, store: memoryStore(), fields, layout: 'adaptive', newId: (kind) => `${kind}_${++next}`, ...more });
  };
  const said = 'Not all the text fits: slide 1 holds about 55 of its 200 words. The rest will be cut. Tell the user before creating the presentation, and offer more slides or a shorter text.';
  // The option, or an app's list of layouts that does not name free.
  for (const more of [{ free: false }, { patterns: ['cover', 'text-only', 'closing'] }]) {
    const planned = await toolsOf(more).call(TOOL_NAMES.plan, { topic: 'Why the city should keep its bike share', language: 'en', slide_count: 1 });
    assert.equal(planned.output.note, `${said} Free slides are off, so the text is not set smaller.`, JSON.stringify(more));
  }
  const deck: DeckSpec = { version: 1, language: 'en', fields, title: 'Bikes', slides: [{ pattern: 'text-only', slots: { title: 'Keep it', body: ['Riders rely on it.'] } }] };
  const change = { summary: 'Longer.', language: 'en', title: 'Bikes', slides: [{ from: 1, pattern: 'text-only', change: 'Make the text 200 words', images: false, words: 200, shows_main: 'auto', side_by_side: 0 }], text_size: 'normal', fields };
  await assert.rejects(planEdit(demo, deck, 'Make slide 1 200 words', { llm: scripted(change).llm, ...options, free: false, stopBeforeCut: true }), (error: unknown) => {
    assert.ok(error instanceof TextCutError);
    assert.equal(error.message, 'Nothing was changed: slide 1 would hold about 55 of its 200 words. Ask for more slides or a shorter text, or accept that the rest is cut. Free slides are off, so the text is not set smaller.');
    return true;
  });
});

test('a slide asked for words keeps them on one slide', async () => {
  // 99 words on a layout that takes 55, in a deck that may have six slides: one slide, cut to what it holds.
  const spread = scripted(outlineOf({ pattern: 'text-only', purpose: 'Keep the bike share', words: 99 }));
  const plan = await planOutline(demo, { ...brief, slideCount: { min: 1, max: 6 } }, { llm: spread.llm, ...options, free: false });
  assert.deepEqual(plan.slides, [{ pattern: 'text-only', purpose: 'Keep the bike share', words: 99 }]);
  assert.deepEqual(planTextFit(demo, plan, { ...options, free: false }), [{ slide: 1, words: 99, holds: 55 }]);

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
    { slide: 1, words: 500, holds: 500, size: 10, templateSize: 18, dense: true, most: 725 },
    { slide: 2, words: 500, holds: 500, size: 10, templateSize: 18, dense: true, most: 725 },
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
  // 6000 characters, and the slide holds 4350 of them at 9 pt, the least that reads on a slide 13.33
  // inches wide, in any twelve paragraphs or fewer of any length: each column keeps a line free and
  // holds its lines with the space after each paragraph, and each paragraph may end a character
  // into a line. `most` is that room found in steps of 50 characters: 4350.
  assert.deepEqual(planTextFit(demo, plan, options), [{ slide: 1, words: 1000, holds: 725, size: 9, templateSize: 18, dense: true, most: 725 }]);

  const writer = scripted({ s1: { title: 'Keep it', b1Text: paragraphs(4350, 12) } });
  const written = await writeDeck(demo, plan, { llm: writer.llm, ...options });
  assert.match(String(slotOf(writer.requests[0], 's1', 'b1Text')?.description), /^Write 725 words, 3910 to 4350 characters: of the 1000 words asked for this slide, the most it holds\./);
  assert.deepEqual(withoutCodes(written.warnings.filter(asked)), [
    { severity: 'warning', slide: 1, message: '1000 words were asked for this slide, and it holds about 725 even with its text set small, at 9 pt: the rest is cut. Ask for the text over more slides, or for a shorter one' },
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
  assert.match(title.requests[0]?.prompt ?? '', /\n1\. free \(a text; 500 words; text at 10 pt; at most about 720 words\): Keep the bike share\n/);

  // On the right-to-left fixture, whose frame leaves a text a narrow room, 500 Hebrew words are more than the slide holds at 7 pt.
  const narrow: DeckPlan = { version: 1, title: 'x', brief: rtlBrief, slides: [{ pattern: 'free', purpose: 'השירות', blocks: ['text'], arrange: { dense: true }, words: 500 }] };
  const same = scripted({ summary: 'בוצע', language: 'he', title: 'y', slides: [textEntry({ purpose: 'השירות', words: 500 })] });
  const kept2 = await editPlan(rtl, narrow, 'כותרת טובה יותר', { llm: same.llm, layout: 'adaptive' });
  assert.match(same.requests[0]?.prompt ?? '', /\n1\. free \(a text; 500 words; cut to fit; text at 7 pt; at most about 230 words\): השירות\n/);
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
  assert.match(simpler.requests[0]?.prompt ?? '', new RegExp(`\\n1\\. \\[free: a text; ${had} words; text at [\\d.]+ pt; at most about 720 words\\] title: `));
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

test("a photo gives way to an asked text, and a plan's texts become one; a list or a chart that holds most of it does not, nor a deck's second text", async () => {
  const photo = { asset: 'bikes', alt: 'Bikes at a station' };
  const assets = [{ id: 'bikes', kind: 'photo' as const, description: 'Bikes at a station', data: new Uint8Array([1]) }];
  const change = (pattern: string) => ({ summary: 'Longer.', language: 'en', title: 'Bikes', slides: [{ from: 1, pattern, change: 'Make the text 500 words', images: false, words: 500, parts: [], shows_main: 'auto', side_by_side: 0 }], text_size: 'normal', fields });
  const left = { severity: 'warning', slide: 1, message: 'the photo of this slide is left out to make room for the 500 words asked for it' };
  // A free slide of a text and a photo: the text alone, set dense, and the deck is told of the photo.
  const free = freeSlide({ title: 'Keep it', blocks: [{ text: ['Riders rely on it.'] }, { photo }] });
  const one = scripted(change('free'), { s1: { title: 'Keep it', b1Text: paragraphs(3100, 8) } });
  const edited = await planEdit(demo, { version: 1, language: 'en', fields, title: 'Bikes', slides: [free] }, 'Make slide 1 500 words', { llm: one.llm, ...options, assets });
  assert.deepEqual([edited.deck.slides[0]?.blocks, edited.deck.slides[0]?.arrange], [['text'], { dense: true }]);
  assert.deepEqual(withoutCodes(edited.warnings), [left]);
  // A slide of a layout with a photo: the same.
  const layout: DeckSpec = { version: 1, language: 'en', fields, title: 'Bikes', slides: [{ pattern: 'text-and-photo', slots: { title: 'Keep it', body: 'Riders rely on it.', photo } }] };
  const two = scripted(change('text-and-photo'), { s1: { title: 'Keep it', b1Text: paragraphs(3100, 8) } });
  const relaid = await planEdit(demo, layout, 'Make slide 1 500 words', { llm: two.llm, ...options, assets });
  assert.deepEqual([relaid.deck.slides[0]?.blocks, relaid.deck.slides[0]?.arrange], [['text'], { dense: true }]);
  assert.deepEqual(withoutCodes(relaid.warnings), [left]);

  // A list that holds more than half of the words: its part stays, and the slide holds what it can of them, which the plan says.
  const listed = scripted(outlineOf(entry('free', { shows: [{ kind: 'list', count: 1 }], words: 300 })));
  const plan = await planOutline(demo, brief, { llm: listed.llm, ...options });
  assert.deepEqual([plan.slides[0]?.blocks, plan.slides[0]?.arrange], [['list'], undefined]);
  const [listFit] = planTextFit(demo, plan, options);
  assert.ok(listFit && listFit.holds < 300 && listFit.holds >= 150 && listFit.dense === undefined, JSON.stringify(listFit));
  const items = { s1: { title: 'Keep it', b1Items: ['Riders rely on it.', 'The city gains.'] } };
  const writer = scripted(items, items, items);
  const written = await writeDeck(demo, plan, { llm: writer.llm, ...options });
  assert.deepEqual(withoutCodes(written.warnings.filter(asked)), [{ severity: 'warning', slide: 1, message: `300 words were asked for this slide, and it holds about ${listFit?.holds}: ask for the text over more slides, or for a shorter one` }]);
  // One that would hold less than half of them gives way: the text alone, set dense, holds them.
  const longer = await planOutline(demo, brief, { llm: scripted(outlineOf(entry('free', { shows: [{ kind: 'list', count: 1 }], words: 700 }))).llm, ...options });
  assert.deepEqual([longer.slides[0]?.blocks, longer.slides[0]?.arrange], [['text'], { dense: true }]);
  assert.deepEqual(
    planTextFit(demo, longer, options).map((fit) => [fit.words, fit.holds]),
    [[700, 700]],
  );
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
  // A chart that holds most of the words keeps its chart, and the slide holds what it can.
  const rides = { ...brief, prompt: 'Rides by quarter in 2025: 12, 18, 23 and 14 thousand.' };
  const charted = scripted(outlineOf({ pattern: 'chart-native', purpose: 'Rides by quarter', words: 150 }));
  const chart = await planOutline(demo, rides, { llm: charted.llm, ...options });
  const kept = chart.slides[0];
  assert.ok(kept?.pattern === 'chart-native' || kept?.blocks?.includes('chart'), JSON.stringify(kept));
  assert.equal(kept?.arrange?.dense, undefined);
  const [fit] = planTextFit(demo, chart, options);
  assert.ok(fit && fit.holds < 150 && fit.holds >= 75, JSON.stringify(fit));
  // One that would hold less than half of them is the text alone, set dense, as cards are.
  for (const shows of [undefined, [{ kind: 'text', count: 1 }, { kind: 'card', count: 3 }]]) {
    const answer = shows ? entry('free', { shows, words: 500 }) : { pattern: 'chart-native', purpose: 'Rides by quarter', words: 500 };
    const alone = await planOutline(demo, rides, { llm: scripted(outlineOf(answer)).llm, ...options });
    assert.deepEqual([alone.slides[0]?.pattern, alone.slides[0]?.blocks, alone.slides[0]?.arrange], ['free', ['text'], { dense: true }], JSON.stringify(shows));
    assert.deepEqual(
      planTextFit(demo, alone, options).map((one) => [one.words, one.holds]),
      [[500, 500]],
    );
  }
});

test('a change to a plan keeps a chart or cards the request adds beside a text it asks words for', async () => {
  const text = { shows: [{ kind: 'text', count: 1 }], shows_main: 'auto', side_by_side: 0 };
  const two = { ...brief, prompt: 'Two slides on the bike share. Trips grew from 12 to 18 thousand.', slideCount: 2 };
  const plan = await planOutline(demo, two, { llm: scripted(outlineOf({ pattern: 'free', purpose: 'Growth', ...text }, { pattern: 'free', purpose: 'Asks', ...text })).llm, ...options });
  for (const [kind, count] of [
    ['chart', 1],
    ['card', 3],
  ] as const) {
    const shows = [{ kind: 'text', count: 1 }, { kind, count }];
    const parts = ['text', ...Array.from({ length: count }, () => kind)];
    const change = { summary: 'Added them and made the text 300 words.', language: 'en', title: 'Bikes', slides: [{ pattern: 'free', purpose: 'Growth', shows, shows_main: 'auto', side_by_side: 0, words: 300 }, { pattern: 'free', purpose: 'Asks', ...text }] };
    const { plan: next } = await editPlan(demo, plan, `Add ${kind === 'chart' ? 'a chart of the growth' : 'three cards'} to slide 1 and make its text 300 words`, { llm: scripted(change).llm, ...options });
    // The part asked for now stays, and the plan says what of the text it leaves out.
    assert.deepEqual(next.slides[0]?.blocks, parts);
    const [fit] = planTextFit(demo, next, options);
    assert.ok(fit !== undefined && fit.words === 300 && fit.holds < 300, JSON.stringify(fit));
    // It stays at the next change too: 500 words asked for that slide now, whatever the answer
    // calls it or in whatever order it lists its parts.
    for (const [said, slide] of [
      ['as it was', { ...change.slides[0], words: 500 }],
      ['renamed', { ...change.slides[0], purpose: 'Growth of the bike share in detail', words: 500 }],
      ['parts first', { ...change.slides[0], shows: [...shows].reverse(), words: 500 }],
    ] as const) {
      const asked = { ...change, slides: [slide, change.slides[1]] };
      const { plan: longer } = await editPlan(demo, next, 'Make the text of slide 1 500 words', { llm: scripted(asked).llm, ...options });
      assert.deepEqual([...(longer.slides[0]?.blocks ?? [])].sort(), [...parts].sort(), `${kind} ${said}`);
      assert.equal(longer.slides[0]?.arrange?.dense, undefined, `${kind} ${said}`);
    }
    // A new slide of the same parts put where the plan had one is new: its parts stay too, before
    // the slide the plan had there or after it.
    for (const at of [0, 1]) {
      const added = { pattern: 'free', purpose: 'Riders', shows, shows_main: 'auto', side_by_side: 0, words: 300 };
      const slides = [...next.slides.map((slide) => ({ pattern: slide.pattern, purpose: slide.purpose, shows: slide.blocks?.includes(kind) ? shows : text.shows, shows_main: 'auto', side_by_side: 0 }))];
      slides.splice(at, 0, added);
      const inserted = { ...change, slides };
      const { plan: more } = await editPlan(demo, next, `Add a slide on the riders with ${kind === 'chart' ? 'a chart' : 'three cards'} and 300 words`, { llm: scripted(inserted).llm, ...options });
      assert.deepEqual(more.slides[at]?.blocks, parts, `${kind} at ${at}`);
    }
  }
});

test('a part added to a plan with no document and no words stays when words are asked for its slide later', async () => {
  // A part a change adds is marked for the changes after it, in a plan of any kind. The deck is as
  // 0.6.2 wrote it; the mark is internal. Asked for words in a later change, the slide keeps it.
  const two = { ...brief, prompt: 'Two slides on the bike share.', slideCount: 2 };
  const text = { shows: [{ kind: 'text', count: 1 }], shows_main: 'auto', side_by_side: 0 };
  const plan = await planOutline(demo, two, { llm: scripted(outlineOf({ pattern: 'free', purpose: 'Growth', ...text }, { pattern: 'free', purpose: 'Asks', ...text })).llm, ...options });
  const change = {
    summary: 'A chart and cards.',
    language: 'en',
    title: 'Bikes',
    slides: [
      { pattern: 'free', purpose: 'Growth', shows: [{ kind: 'text', count: 1 }, { kind: 'chart', count: 1 }], shows_main: 'auto', side_by_side: 0 },
      { pattern: 'free', purpose: 'Asks', ...text },
      { pattern: 'free', purpose: 'Cards', shows: [{ kind: 'text', count: 1 }, { kind: 'card', count: 3 }], shows_main: 'auto', side_by_side: 0 },
    ],
  };
  const { plan: next } = await editPlan(demo, plan, 'Add a chart to slide 1 and a slide of cards', { llm: scripted(change).llm, ...options });
  assert.deepEqual(
    next.slides.map((slide) => slide.blocks),
    [['text', 'chart'], ['text'], ['text', 'card', 'card', 'card']],
  );
  assert.deepEqual(
    next.slides.map((slide) => slide.partsAsked),
    [true, undefined, true],
  );
  const longer = {
    summary: 'Longer.',
    language: 'en',
    title: 'Bikes',
    slides: [{ ...change.slides[0], words: 500 }, change.slides[1], change.slides[2]],
  };
  const { plan: last } = await editPlan(demo, next, 'Make slide 1 500 words', { llm: scripted(longer).llm, ...options });
  assert.deepEqual(last.slides[0]?.blocks, ['text', 'chart']);
  assert.equal(last.slides[0]?.partsAsked, true);
  // A layout with a chart asked for in a change keeps its chart too, as in 0.6.2: on a free slide
  // beside the text, which holds what it can.
  const layout = { pattern: 'chart-native', purpose: 'Asks', shows: [], shows_main: 'auto', side_by_side: 0 };
  const charted = { summary: 'A chart.', language: 'en', title: 'Bikes', slides: [{ pattern: 'free', purpose: 'Growth', ...text }, layout] };
  const { plan: withChart } = await editPlan(demo, plan, 'Make slide 2 a chart', { llm: scripted(charted).llm, ...options });
  assert.equal(withChart.slides[1]?.partsAsked, true);
  const worded = { ...charted, slides: [charted.slides[0], { ...layout, words: 500 }] };
  const { plan: chartWords } = await editPlan(demo, withChart, 'Make slide 2 500 words', { llm: scripted(worded).llm, ...options });
  assert.deepEqual([chartWords.slides[1]?.pattern, chartWords.slides[1]?.blocks], ['free', ['text', 'chart']]);
});

test('in a change to a plan, parts the plan had give way to the words asked for their slide, however the answer names it', async () => {
  // A plan whose slides have cards and a chart: planned so, not asked for in a change. Asked for
  // more words than they leave room for, a slide is its text alone, set dense. The answer may
  // reword the slide's purpose or list its parts in another order: it is the same slide.
  const shown = (kind: 'card' | 'chart', count: number) => ({ shows: [{ kind: 'text', count: 1 }, { kind, count }], shows_main: 'auto', side_by_side: 0 });
  const two = { ...brief, prompt: 'Two slides on the bike share. Trips grew from 12 to 18 thousand.', slideCount: 2 };
  for (const [kind, count] of [
    ['card', 3],
    ['chart', 1],
  ] as const) {
    const plan = await planOutline(demo, two, { llm: scripted(outlineOf({ pattern: 'free', purpose: 'Growth', ...shown(kind, count) }, { pattern: 'free', purpose: 'Asks', ...shown(kind, count) })).llm, ...options });
    assert.deepEqual(plan.slides[0]?.blocks, ['text', ...Array.from({ length: count }, () => kind)]);
    assert.equal(plan.slides[0]?.partsAsked, undefined);
    for (const [said, slide] of [
      ['as it was', { pattern: 'free', purpose: 'Growth', ...shown(kind, count), words: 500 }],
      ['renamed', { pattern: 'free', purpose: 'Growth of the bike share in detail', ...shown(kind, count), words: 500 }],
      ['parts first', { pattern: 'free', purpose: 'Growth', ...shown(kind, count), shows: [...shown(kind, count).shows].reverse(), words: 500 }],
    ] as const) {
      const change = { summary: 'Longer.', language: 'en', title: 'Bikes', slides: [slide, { pattern: 'free', purpose: 'Asks', ...shown(kind, count) }] };
      const { plan: longer } = await editPlan(demo, plan, 'Make the text of slide 1 500 words', { llm: scripted(change).llm, ...options });
      assert.deepEqual([longer.slides[0]?.blocks, longer.slides[0]?.arrange], [['text'], { dense: true }], `${kind} ${said}`);
      assert.deepEqual(
        planTextFit(demo, longer, options).map((fit) => [fit.words, fit.holds]),
        [[500, 500]],
        `${kind} ${said}`,
      );
      // The slide the answer inserts before it is new, and keeps its parts; the plan's slide after it
      // still gives way. Inserted before a slide whose purpose the answer also rewords, the new slide
      // stands where the plan's did and is taken for it: nothing else tells them apart.
      if (said === 'renamed') continue;
      const inserted = { ...change, slides: [{ pattern: 'free', purpose: 'Riders', ...shown(kind, count), words: 500 }, ...change.slides] };
      const { plan: more } = await editPlan(demo, plan, 'Add a slide on the riders before slide 1, and make both 500 words', { llm: scripted(inserted).llm, ...options });
      assert.deepEqual(
        more.slides.map((one) => one.blocks?.length === 1),
        [false, true, false],
        `${kind} ${said} inserted`,
      );
      assert.equal(more.slides[0]?.partsAsked, true, `${kind} ${said} inserted`);
    }
  }
});

test('a free text with asked words validates at the top of what its writer is told, however its paragraphs fall', async () => {
  // The writer is told a range and a number of paragraphs (`maxItems`). Written at the top of the
  // range in any number of paragraphs up to it, alike, or one long and the others short, the text
  // stays inside its room: it is never sent back, and the deck says nothing of it. A slide that
  // holds less than was asked tells its writer what it holds, and that is so too.
  const cases = [
    { template: rtl, base: rtlBrief, language: 'he', sentence: SENTENCES.he, all: [200, 230, 250, 270] },
    { template: demo, base: brief, language: 'en', sentence: SENTENCES.en, all: [150, 200, 225, 700] },
  ];
  /** A text of `chars` characters in `count` paragraphs: all alike, or all but one of `short` characters and the rest in that one. */
  const shaped = (chars: number, count: number, sentence: string, short?: number): string[] => {
    if (short === undefined || count === 1 || chars - (count - 1) * short < short) return paragraphs(chars, count, sentence);
    return [...paragraphs(chars - (count - 1) * short, 1, sentence), ...Array.from({ length: count - 1 }, (_, at) => paragraphs(short, 1, sentence.slice(at % 7) + sentence)[0] as string)];
  };
  for (const { template, base, language, sentence, all } of cases) {
    for (const words of all) {
      const { llm } = scripted(outlineOf(textEntry({ words })));
      const plan = await planOutline(template, base, { llm, ...options });
      assert.equal(plan.slides[0]?.arrange?.dense, true);
      const probe = scripted({ s1: { title: 'A title', b1Text: paragraphs(charsOfWords(words, language), 3, sentence) } });
      await writeDeck(template, plan, { llm: probe.llm, ...options, maxRepairs: 0 });
      const told = slotOf(probe.requests[0], 's1', 'b1Text');
      const top = Number(/ to (\d+) characters/.exec(String(told?.description))?.[1]);
      const most = told?.maxItems ?? 0;
      assert.ok(top > 0 && most > 0, String(told?.description));
      for (const count of [1, 2, Math.ceil(most / 2), most]) {
        for (const short of [undefined, 30, 60, 90, 130]) {
          const writer = scripted({ s1: { title: 'A title', b1Text: shaped(top, count, sentence, short) } });
          const written = await writeDeck(template, plan, { llm: writer.llm, ...options, maxRepairs: 0 });
          assert.deepEqual(
            written.warnings.filter((issue) => /characters|lines|paragraphs/.test(issue.message)),
            [],
            `${language} ${words} words: ${top} in ${count} paragraphs (${short ?? 'alike'})`,
          );
        }
      }
    }
  }
});

test('a text set dense that is sent back is told what it holds in the paragraphs it was told, not in twelve', async () => {
  // A document's text that its slide holds only with some of its paragraphs joined: its writer is
  // told the most paragraphs that hold all of it, more than twelve. Written over what it was told in
  // those paragraphs, it is sent back for its characters alone; written in more, it is told the
  // number it was told. A text the slide holds only cut is told twelve, and that is what it hears.
  const docOf = (count: number, chars: number) => Array.from({ length: count }, (_, at) => `Point ${at + 1}: ${SENTENCES.en.repeat(3).slice(at % 9, (at % 9) + chars - 11).trim()}.`).join('\n\n');
  for (const [chars, cut] of [
    [120, false],
    [210, true],
  ] as const) {
    const outline = { title: 'Bikes', left_out: [], slides: [{ pattern: 'free', purpose: 'All of it', shows: [{ kind: 'text', count: 1 }], shows_main: 'auto', side_by_side: 0, passages: Array.from({ length: 32 }, (_, at) => at + 1) }] };
    const plan = await planOutline(demo, { ...brief, document: docOf(32, chars) }, { llm: scripted(outline).llm, ...options });
    const [fit] = planTextFit(demo, plan, options);
    assert.ok(fit && fit.dense === true && (cut ? fit.holds < fit.words : fit.holds === fit.words), JSON.stringify(fit));
    const probe = scripted({ s1: { title: 'All of it', b1Text: ['x'] } });
    await writeDeck(demo, plan, { llm: probe.llm, ...options, maxRepairs: 0 });
    const told = slotOf(probe.requests[0], 's1', 'b1Text');
    const items = told?.maxItems ?? 0;
    const top = Number(/ to (\d+) characters/.exec(String(told?.description))?.[1]);
    assert.ok(cut ? items === 12 : items > 12 && items < 32, `${chars}: told ${items}`);
    const over = (count: number) => ({ s1: { title: 'All of it', b1Text: paragraphs(Math.round(1.4 * top), count) } });
    for (const count of [items, items + 4]) {
      const writer = scripted(over(count), over(count));
      await writeDeck(demo, plan, { llm: writer.llm, ...options, maxRepairs: 1 });
      assert.equal(writer.requests.length, 2, `${chars}: ${count}`);
      const problems = /It has these problems:\n([\s\S]*?)\n\n/.exec(writer.requests[1]?.prompt ?? '')?.[1] ?? '';
      assert.match(problems, /text is \d+ characters; the slot fits about \d+/, `${chars}: ${count}`);
      if (!cut) assert.doesNotMatch(problems, /the slot fits 12\b/, `${chars}: ${count}`);
      if (count > items) assert.match(problems, new RegExp(`text has ${count} paragraphs; the slot fits ${items}\\b`), `${chars}: ${count}`);
      else assert.doesNotMatch(problems, /paragraphs; the slot fits/, `${chars}: ${count}`);
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
  assert.deepEqual(withoutCodes(still.warnings), [{ severity: 'warning', slide: 1, message: '500 words were asked for this slide, and it has about 300: ask for it again to have them all' }]);
});

test('a change to a deck that would cut a text stops before writing', async () => {
  const deck: DeckSpec = { version: 1, language: 'en', fields, title: 'Bikes', slides: [{ pattern: 'text-only', slots: { title: 'Keep it', body: ['Riders rely on it.'] } }] };
  const change = { summary: 'Longer.', language: 'en', title: 'Bikes', slides: [{ from: 1, pattern: 'text-only', change: 'Make the text 1500 words', images: false, words: 1500, parts: [], shows_main: 'auto', side_by_side: 0 }], text_size: 'normal', fields };
  const stop = scripted(change);
  await assert.rejects(planEdit(demo, deck, 'Make slide 1 1500 words', { llm: stop.llm, ...options, stopBeforeCut: true }), (error: unknown) => {
    assert.ok(error instanceof TextCutError);
    assert.equal(error.message, 'Nothing was changed: slide 1 would hold about 725 of its 1500 words (text at 9 pt). Ask for more slides or a shorter text, or accept that the rest is cut.');
    assert.deepEqual(error.fits, [{ slide: 1, words: 1500, holds: 725, size: 9, dense: true }]);
    return true;
  });
  assert.equal(stop.requests.length, 1);

  // Without the option the slide is written, cut to what it holds, and the deck is told.
  const written = scripted(change, { s1: { title: 'Keep it', b1Text: paragraphs(4350, 12) } });
  const edited = await planEdit(demo, deck, 'Make slide 1 1500 words', { llm: written.llm, ...options });
  assert.deepEqual(withoutCodes(edited.warnings.filter(asked)), [
    { severity: 'warning', slide: 1, message: '1500 words were asked for this slide, and it holds about 725 even with its text set small, at 9 pt: the rest is cut. Ask for the text over more slides, or for a shorter one' },
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
  const change = { summary: 'Longer.', language: 'en', title: 'Bikes', slides: [{ from: 1, pattern: 'free', change: 'Make the text 720 words', images: false, words: 720, parts: [], shows_main: 'auto', side_by_side: 0 }], text_size: 'normal', fields };
  const write = (heading: boolean) => ({ s1: { title: 'Keep it', ...(heading ? { b1Heading: 'An overview' } : {}), b1Text: paragraphs(4300, 12) } });
  // A text alone holds 720 words, set small: nothing stops.
  assert.equal(await stopped(planEdit(demo, deck(false), 'Make slide 1 720 words', { llm: scripted(change, write(false), write(false)).llm, ...options, stopBeforeCut: true })), undefined);
  // The heading the slide keeps takes some of that room. The change stops after its plan, with what
  // the slide holds as it would be written.
  const stop = scripted(change);
  const fits = await stopped(planEdit(demo, deck(true), 'Make slide 1 720 words', { llm: stop.llm, ...options, stopBeforeCut: true }));
  assert.equal(stop.requests.length, 1);
  assert.equal(fits?.length, 1);
  const [fit] = fits ?? [];
  assert.deepEqual({ ...fit, holds: 0 }, { slide: 1, words: 720, holds: 0, size: 9, dense: true });
  assert.ok((fit?.holds ?? 720) < 720, JSON.stringify(fit));
  // Written all the same, it holds as many as the stop said.
  const cut = await planEdit(demo, deck(true), 'Make slide 1 720 words', { llm: scripted(change, write(true), write(true), write(true)).llm, ...options });
  assert.deepEqual(withoutCodes(cut.warnings.filter(asked)), [
    { severity: 'warning', slide: 1, message: `720 words were asked for this slide, and it holds about ${fit?.holds} even with its text set small, at 9 pt: the rest is cut. Ask for the text over more slides, or for a shorter one` },
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
  assert.deepEqual(withoutCodes(edited.warnings), [{ severity: 'warning', slide: 1, message: `300 words were asked for this slide, and it holds about ${fit?.holds}: ask for the text over more slides, or for a shorter one` }]);
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
  assert.deepEqual(withoutCodes(edited.warnings), [{ severity: 'warning', slide: 1, message: 'the photo of this slide is left out to make room for its text of about 300 words' }]);
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
  for (const chars of [4300, 5100]) {
    const text = paragraphs(chars, 9);
    assert.ok(wordsOf(text) * WORD_CHARS > 5300, `${chars}`);
    const deck: DeckSpec = { version: 1, language: 'en', fields, title: 'Bikes', slides: [{ pattern: 'free', blocks: ['text'], arrange: { dense: true }, slots: { title: 'Keep it', b1Text: text } }] };
    for (const [instruction, slide] of changes) {
      const model = scripted(plan(slide), ...Array.from({ length: 4 }, () => ({ s1: { title: 'Keep the bikes', b1Text: text } })));
      const edited = await planEdit(demo, deck, instruction, { llm: model.llm, ...options, stopBeforeCut: true });
      const words = Number(/^Write (\d+) words, as was asked: /.exec(String(slotOf(model.requests[1], 's1', 'b1Text')?.description))?.[1]);
      assert.ok(words > 0 && words <= Math.floor(chars / WORD_CHARS), `${chars} ${instruction}: ${words}`);
      if (chars === 4300) assert.equal(words, 716);
      assert.deepEqual(edited.deck.slides[0]?.arrange, { dense: true });
      assert.deepEqual(edited.warnings.filter(asked), [], `${chars} ${instruction}`);
    }
  }
});

test('a writer inside the range it is told is not sent back, however many paragraphs it writes', async () => {
  // 720 words: 4320 characters on one slide, held at 9 pt, the least that reads on a slide 13.33 inches wide.
  const { llm } = scripted(outlineOf(entry('text-only', { words: 720 })));
  const plan = await planOutline(demo, brief, { llm, ...options });
  assert.deepEqual(planTextFit(demo, plan, options), [{ slide: 1, words: 720, holds: 720, size: 9, templateSize: 18, dense: true, most: 725 }]);
  const first = scripted({ s1: { title: 'Keep it', b1Text: paragraphs(4320, 12) } });
  await writeDeck(demo, plan, { llm: first.llm, ...options });
  const description = String(slotOf(first.requests[0], 's1', 'b1Text')?.description);
  // The range ends where any text of twelve paragraphs or fewer fits, however its paragraphs
  // fall: each may end a character into a line. A budget of what twelve alike paragraphs held
  // let a writer inside it be sent back.
  assert.match(description, /^Write 720 words, as was asked: 4130 to 4350 characters in all\./);
  assert.equal(limitOf(description), 4350);
  for (let count = 1; count <= 12; count++) {
    for (const lead of [0, 1]) {
      // Alike, or one long paragraph and the others each just over a line of 54 characters.
      const text = lead === 0 || count === 1 ? paragraphs(4350, count) : [...paragraphs(4350 - (count - 1) * 56, 1), ...Array.from({ length: count - 1 }, () => paragraphs(56, 1)[0] as string)];
      const writer = scripted({ s1: { title: 'Keep it', b1Text: text } });
      const written = await writeDeck(demo, plan, { llm: writer.llm, ...options });
      assert.equal(writer.requests.length, 1, `${count} paragraphs`);
      assert.deepEqual(written.warnings, [], `${count} paragraphs`);
    }
  }
});
