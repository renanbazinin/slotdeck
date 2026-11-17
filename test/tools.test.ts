import assert from 'node:assert/strict';
import { mkdtempSync, readFileSync, readdirSync, rmSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { test } from 'node:test';
import { placeholderPhoto } from '../examples/placeholder.ts';
import { PptxPackage, TOOL_NAMES, loadTemplate, memoryStore, presentationTools, storeKeys } from '../src/index.ts';
import type { DeckPlan, DeckSpec, ImageRef, JsonRequest, LlmProvider } from '../src/index.ts';
import { directoryStore } from '../src/node.ts';
import { readPresentation } from '../src/pptx/presentation.ts';
import { buildFixturePptx, fixtureManifest } from './fixture.ts';

const template = loadTemplate(buildFixturePptx(), fixtureManifest);
const png = placeholderPhoto(20, 10, [0, 0, 0], [255, 255, 255]);

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

const outline = { title: 'Dogs', slides: [{ pattern: 'content', purpose: 'Why dogs' }, { pattern: 'closing', purpose: 'Thanks' }] };
const revisedOutline = {
  summary: 'Added a slide about cats.',
  language: 'en',
  title: 'Dogs and cats',
  slides: [
    { pattern: 'content', purpose: 'Why dogs' },
    { pattern: 'content', purpose: 'Why cats' },
    { pattern: 'closing', purpose: 'Thanks' },
  ],
};
const slideContent = (title: string, prompt: string) => ({ title, body: ['Loyal'], note: 'A note', photo: { prompt, alt: 'Animal' } });

function setup(...answers: unknown[]) {
  const model = scripted(...answers);
  const store = memoryStore();
  let generated = 0;
  let next = 0;
  const tools = presentationTools({
    template,
    llm: model.llm,
    store,
    fields: { closing: 'Thank you' },
    images: {
      async generateImage() {
        generated++;
        return png;
      },
    },
    thumbnail: (pattern) => `/thumbs/${pattern}.png`,
    newId: (kind) => `${kind}_${++next}`,
  });
  return { ...model, store, tools, generated: () => generated };
}

test('the tools are declared with the arguments a chat model must give', () => {
  const { tools } = setup();
  assert.deepEqual(
    tools.definitions.map((tool) => [tool.name, (tool.parameters as { required: string[] }).required]),
    [
      ['plan_presentation', ['topic']],
      // A change of colour alone needs no instruction, so the instruction is required of neither tool that changes something.
      ['edit_presentation_plan', ['plan_id']],
      ['create_presentation', ['plan_id']],
      ['edit_presentation', ['presentation_id']],
    ],
  );
  for (const tool of tools.definitions) assert.ok(tool.description.length > 40);
});

test('plan, change the plan, create, change the deck', async () => {
  const { tools, store, requests, generated } = setup(
    outline,
    revisedOutline,
    { s1: slideContent('Dogs are great', 'A dog'), s2: slideContent('Cats are great', 'A cat'), s3: {} },
    {
      summary: 'Shortened slide 1.',
      language: 'en',
      title: 'Dogs and cats',
      fields: { closing: 'Thank you' },
      slides: [
        { from: 1, pattern: 'content', change: 'Shorter title', images: false },
        { from: 2, pattern: 'content', change: '', images: false },
        { from: 3, pattern: 'closing', change: '', images: false },
      ],
    },
    { s1: { title: 'Dogs', body: ['Loyal'], note: 'A note' } },
  );

  // 1. Plan: one model call, nothing rendered.
  const planned = await tools.call(TOOL_NAMES.plan, { topic: 'A deck about dogs', language: 'en', slide_count: 2 });
  assert.equal(planned.output.plan_id, 'plan_1');
  assert.deepEqual(planned.output.slides, [
    { number: 1, layout: 'content', layout_description: 'Title, body, photo and icon.', purpose: 'Why dogs', thumbnail: '/thumbs/content.png' },
    { number: 2, layout: 'closing', layout_description: 'Closing words.', purpose: 'Thanks', thumbnail: '/thumbs/closing.png' },
  ]);
  assert.equal(requests.length, 1);
  assert.equal(generated(), 0);
  assert.equal(planned.plan?.brief.slideCount, 2);

  // 2. Change the plan: a new plan under a new id; the old one stays.
  const replanned = await tools.call(TOOL_NAMES.editPlan, { plan_id: 'plan_1', instruction: 'Add a slide about cats' });
  assert.equal(replanned.output.plan_id, 'plan_2');
  assert.equal(replanned.output.previous_plan_id, 'plan_1');
  assert.equal(replanned.output.summary, 'Added a slide about cats.');
  assert.equal((replanned.output.slides as unknown[]).length, 3);
  assert.match(requests[1]?.prompt ?? '', /1\. content: Why dogs\n2\. closing: Thanks/);
  assert.match(requests[1]?.prompt ?? '', /Requested change:\nAdd a slide about cats/);
  assert.equal(replanned.plan?.brief.instructions, 'Add a slide about cats', 'the request travels with the brief to the writing stage');

  // 3. Create: content is written, pictures generated and stored apart, the file rendered.
  const created = await tools.call(TOOL_NAMES.create, { plan_id: 'plan_2' });
  assert.equal(created.output.presentation_id, 'deck_3');
  assert.equal(created.output.file, 'pptx/deck_3.pptx');
  assert.deepEqual(created.output.slides, [
    { number: 1, layout: 'content', title: 'Dogs are great' },
    { number: 2, layout: 'content', title: 'Cats are great' },
    { number: 3, layout: 'closing' },
  ]);
  assert.equal(generated(), 2);
  assert.match(requests[2]?.prompt ?? '', /Also respect:\nAdd a slide about cats/);
  assert.equal(readPresentation(PptxPackage.load(created.pptx as Uint8Array)).slides.length, 3);

  const dogPhoto = (created.deck?.slides[0]?.slots.photo as ImageRef).asset;
  assert.deepEqual(await store.get(storeKeys.asset(dogPhoto)), png, 'the picture is stored on its own');
  const savedDeck = new TextDecoder().decode(await store.get(storeKeys.deck('deck_3')));
  assert.ok(savedDeck.length < 2000 && savedDeck.includes('"prompt":"A dog"'), 'the deck JSON is small text holding the picture description');
  assert.deepEqual(await store.get(storeKeys.pptx('deck_3')), created.pptx);

  // 4. Change the deck: only slide 1 is rewritten and no picture is generated again.
  const edited = await tools.call(TOOL_NAMES.edit, { presentation_id: 'deck_3', instruction: 'Shorten the title of slide 1' });
  assert.equal(edited.output.presentation_id, 'deck_4');
  assert.equal(edited.output.previous_presentation_id, 'deck_3');
  assert.equal(edited.output.summary, 'Shortened slide 1.');
  assert.deepEqual(edited.output.rewritten_slides, [1]);
  assert.equal(generated(), 2, 'both photos came from the store');
  assert.equal(edited.deck?.slides[0]?.slots.title, 'Dogs');
  assert.deepEqual(edited.deck?.slides[1], created.deck?.slides[1]);

  assert.deepEqual(
    store.keys().map((key) => key.split('/')[0]).sort(),
    ['asset', 'asset', 'deck', 'deck', 'plan', 'plan', 'pptx', 'pptx'],
  );
  // What goes back to the model stays small and textual.
  for (const result of [planned, replanned, created, edited]) assert.ok(JSON.stringify(result.output).length < 1500);
});

test('problems come back as an error the model can act on', async () => {
  const { tools } = setup({ title: 'T', slides: [{ pattern: 'made-up', purpose: 'p' }] }, { title: 'T', slides: [{ pattern: 'made-up', purpose: 'p' }] });

  assert.match(String((await tools.call(TOOL_NAMES.create, { plan_id: 'nope' })).output.error), /no plan with id "nope"/);
  assert.match(String((await tools.call(TOOL_NAMES.edit, { presentation_id: 'nope', instruction: 'x' })).output.error), /no presentation with id "nope"/);
  assert.match(String((await tools.call(TOOL_NAMES.plan, {})).output.error), /"topic" is required/);
  assert.match(String((await tools.call('unknown_tool', {})).output.error), /Unknown tool/);
  assert.match(String((await tools.call(TOOL_NAMES.plan, { topic: 'x' })).output.error), /usable outline/);
});

test('only a declared tool is run, whatever is given for a name', async () => {
  const { tools, requests } = setup();
  // "constructor" and "toString" are keys of every object, and no tools.
  for (const name of ['constructor', 'toString', '__proto__', 'hasOwnProperty', 'valueOf', '__defineGetter__', '', 'PLAN_PRESENTATION', 'plan_presentation ']) {
    assert.deepEqual(await tools.call(name, { topic: 'Bikes' }), { output: { error: `Unknown tool "${name}"` } });
  }
  // A name that is no text is no tool either, and the call still answers: a symbol cannot stand in a text as it is.
  for (const [name, said] of [[5, '5'], [null, 'null'], [undefined, 'undefined'], [Symbol('x'), 'Symbol(x)'], [{}, '[object Object]']] as const) {
    assert.deepEqual(await tools.call(name as unknown as string, { topic: 'Bikes' }), { output: { error: `Unknown tool "${said}"` } });
  }
  assert.equal(requests.length, 0);
});

test('a number of slides beyond what the plan tool declares is held to it', async () => {
  const { tools, requests } = setup(outline, outline, outline);
  const declared = (tools.definitions[0]?.parameters as { properties: { slide_count: { minimum: number; maximum: number } } }).properties.slide_count;
  assert.deepEqual([declared.minimum, declared.maximum], [1, 30]);

  for (const [given, taken] of [[30, 30], [31, 30], [1e9, 30]] as const) {
    const planned = await tools.call(TOOL_NAMES.plan, { topic: 'A deck about dogs', language: 'en', slide_count: given });
    assert.equal(planned.plan?.brief.slideCount, taken);
    const request = requests.at(-1);
    assert.match(request?.prompt ?? '', new RegExp(`Plan a deck of exactly ${taken} slides\\.`));
    const slides = (request?.schema as { properties: { slides: { minItems: number; maxItems: number } } }).properties.slides;
    assert.deepEqual([slides.minItems, slides.maxItems], [taken, taken]);
    // The room for the answer is that of thirty slides, not of a thousand million.
    assert.equal(request?.maxOutputTokens, requests[0]?.maxOutputTokens);
  }
});

test('what the software trips over in the store is said in words; what the parts of the app say is passed on', async () => {
  const { tools, store, requests } = setup();
  const put = (key: string, text: string) => store.set(key, new TextEncoder().encode(text));
  const noPlan = 'The stored plan could not be read. Make a new plan with plan_presentation.';
  const noDeck = 'The stored presentation could not be read. Make it again: a new plan with plan_presentation, then create_presentation.';
  // No JSON at all, JSON that is no plan and no presentation, and one with a part missing: a plan
  // without its brief, which is also a presentation whose slide has no slots.
  const partly = '{"version":1,"title":"T","slides":[{"pattern":"content","purpose":"p"}]}';
  const kept = ['not json {', '', '5', '"plan"', '[]', '{}', '{"version":1,"slides":null}', '{"version":1,"title":"T","brief":{"prompt":"x"},"slides":[null]}', partly];
  for (const text of kept) {
    await put(storeKeys.plan('bad'), text);
    await put(storeKeys.deck('bad'), text);
    assert.deepEqual((await tools.call(TOOL_NAMES.create, { plan_id: 'bad' })).output, { error: noPlan }, text);
    assert.deepEqual((await tools.call(TOOL_NAMES.editPlan, { plan_id: 'bad', instruction: 'Add a slide' })).output, { error: noPlan }, text);
    assert.deepEqual((await tools.call(TOOL_NAMES.edit, { presentation_id: 'bad', instruction: 'Shorten slide 1' })).output, { error: noDeck }, text);
    assert.deepEqual((await tools.call(TOOL_NAMES.edit, { presentation_id: 'bad', accent_color: '1E6B3A' })).output, { error: noDeck }, text);
    // A colour alone is set by code, which needs no brief: the plan that lacks one is recoloured as it is.
    if (text !== partly) assert.deepEqual((await tools.call(TOOL_NAMES.editPlan, { plan_id: 'bad', accent_color: '1E6B3A' })).output, { error: noPlan }, text);
  }
  assert.equal(requests.length, 0, 'no model is asked about a plan that cannot be read');
  // What the store does not hold is still said as that.
  await put(storeKeys.plan('bad'), 'null');
  assert.match(String((await tools.call(TOOL_NAMES.create, { plan_id: 'bad' })).output.error), /no plan with id "bad"/);
  // The plan tool reads nothing from the store. Where the software trips there, on arguments that cannot be read, it says that.
  const noPlanning = 'The presentation could not be planned: the software failed on this request. Tell the user that it did not work.';
  const unreadable = Object.defineProperty({}, 'topic', { enumerable: true, get: () => (undefined as unknown as { trim(): string }).trim() });
  assert.deepEqual((await tools.call(TOOL_NAMES.plan, unreadable)).output, { error: noPlanning });
  // The page on the tools quotes the three sentences as they are.
  const page = readFileSync(new URL('../docs/chat-tools.md', import.meta.url), 'utf8').replace(/\s+/g, ' ');
  for (const sentence of [noPlan, noDeck, noPlanning]) assert.ok(page.includes(`\`${sentence}\``), sentence);

  // A network that fails is a TypeError too. What the app's own model says goes to the chat model as it is,
  const down = presentationTools({ template, store: memoryStore(), fields: { closing: 'x' }, llm: { generateJson: () => Promise.reject(new TypeError('fetch failed')) } });
  assert.deepEqual((await down.call(TOOL_NAMES.plan, { topic: 'Dogs' })).output, { error: 'fetch failed' });
  // and so does what its store says,
  const closing = { title: 'T', slides: [{ pattern: 'closing', purpose: 'p' }] };
  const full = presentationTools({ template, llm: scripted(closing).llm, fields: { closing: 'x' }, store: { get: async () => undefined, set: () => Promise.reject(new RangeError('The disk is full')) } });
  assert.deepEqual((await full.call(TOOL_NAMES.plan, { topic: 'Dogs' })).output, { error: 'The disk is full' });
  // and its image provider.
  const blind = setup(outline, { s1: slideContent('Dogs are great', 'A dog'), s2: {} });
  const dark = presentationTools({ template, llm: blind.llm, store: blind.store, fields: { closing: 'x' }, newId: () => 'plan_a', images: { generateImage: () => Promise.reject(new TypeError('fetch failed')) } });
  await dark.call(TOOL_NAMES.plan, { topic: 'Dogs' });
  assert.deepEqual((await dark.call(TOOL_NAMES.create, { plan_id: 'plan_a' })).output, { error: 'fetch failed' });
});

test('without an image provider, plans carry no photos', async () => {
  const model = scripted({ title: 'T', slides: [{ pattern: 'closing', purpose: 'p' }] });
  const tools = presentationTools({ template, llm: model.llm, store: memoryStore(), fields: { closing: 'Bye' } });
  const planned = await tools.call(TOOL_NAMES.plan, { topic: 'x', image_density: 'high' });
  assert.equal(planned.plan?.brief.imageDensity, 'none');
  assert.match(model.requests[0]?.prompt ?? '', /Do not plan any photos/);
});

test('a folder works as the store, and keys cannot leave it', async () => {
  const folder = mkdtempSync(join(tmpdir(), 'slotdeck-store-'));
  try {
    const store = directoryStore(folder);
    assert.equal(await store.get('deck/missing.json'), undefined);
    await store.set('asset/gen-1', png);
    assert.deepEqual(await store.get('asset/gen-1'), png);
    assert.deepEqual(readdirSync(join(folder, 'asset')), ['gen-1']);
    await assert.rejects(store.set('../outside', png), /leaves the store folder/);

    // The tools run on it unchanged.
    const model = scripted(outline);
    const tools = presentationTools({ template, llm: model.llm, store, fields: { closing: 'x' }, images: { generateImage: async () => png }, newId: () => 'plan_a' });
    await tools.call(TOOL_NAMES.plan, { topic: 'Dogs' });
    const saved = JSON.parse(new TextDecoder().decode(await store.get(storeKeys.plan('plan_a')))) as DeckPlan;
    assert.equal(saved.title, 'Dogs');
  } finally {
    rmSync(folder, { recursive: true, force: true });
  }
});

test('a deck kept as JSON renders again from the store alone', async () => {
  const { tools, store } = setup(outline, { s1: slideContent('Dogs are great', 'A dog'), s2: {} });
  await tools.call(TOOL_NAMES.plan, { topic: 'Dogs' });
  const created = await tools.call(TOOL_NAMES.create, { plan_id: 'plan_1' });
  const deck = JSON.parse(new TextDecoder().decode(await store.get(storeKeys.deck('deck_2')))) as DeckSpec;
  assert.deepEqual(deck, created.deck);
  assert.ok(!JSON.stringify(deck).includes('iVBOR'), 'no image bytes in the JSON');
});
