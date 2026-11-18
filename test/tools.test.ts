import assert from 'node:assert/strict';
import { mkdtempSync, readFileSync, readdirSync, rmSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { test } from 'node:test';
import { placeholderPhoto } from '../examples/placeholder.ts';
import { PptxPackage, TOOL_NAMES, describeDocuments, documentPassages, formatIssue, loadTemplate, memoryStore, presentationTools, storeKeys } from '../src/index.ts';
import type { ChatDocument, DeckPlan, DeckSpec, ImageRef, JsonRequest, LlmProvider, PresentationToolsOptions, TemplateManifest, ToolDefinition } from '../src/index.ts';
import { directoryStore } from '../src/node.ts';
import { readPresentation } from '../src/pptx/presentation.ts';
import { buildFixturePptx, fixtureManifest } from './fixture.ts';

const template = loadTemplate(buildFixturePptx(), fixtureManifest);
const png = placeholderPhoto(20, 10, [0, 0, 0], [255, 255, 255]);
/** What `next` ends with in a result that has something for the person. */
const IN_THEIR_LANGUAGE = 'Tell the user what this result says in the language they write in, in your own words and with the same slide numbers and figures; never paste its sentences as they are.';

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

// --- the person's numbers and text --------------------------------------------------------------

const demo = loadTemplate(new Uint8Array(readFileSync(new URL('../templates/demo.pptx', import.meta.url))), JSON.parse(readFileSync(new URL('../templates/demo.manifest.json', import.meta.url), 'utf8')) as TemplateManifest);
const demoFields = { organizationDepartment: 'Retail' };
const noModel: LlmProvider = { generateJson: async () => ({}) };
const painter = { generateImage: async () => png };
type Property = { description: string } & Record<string, unknown>;
const propertiesOf = (tool: ToolDefinition | undefined): Record<string, Property> => (tool?.parameters as { properties: Record<string, Property> }).properties;

/** Tools on a template with these options, over a store of their own, with ids counted from 1. */
function toolsOn(on: PresentationToolsOptions['template'], answers: unknown[], more: Partial<PresentationToolsOptions> = {}) {
  const model = scripted(...answers);
  const store = memoryStore();
  let next = 0;
  const tools = presentationTools({ template: on, llm: model.llm, store, newId: (kind) => `${kind}_${++next}`, ...more });
  return { ...model, store, tools };
}

// The texts of 0.6.2 that make the chat model pass the person's numbers on as they said them.
const SLIDE_COUNT =
  'Number of slides in all. Set it whenever the user names a number of slides ("2 slides", "a five-slide deck"): it is not read from the topic or the instructions. Leave it out when they name none: the planner then chooses the number. ';
const SMALLER = {
  slideCount: 'With it the presentation has exactly that many slides: long text is set smaller to fit, and no slide is added.',
  plan: 'The planner keeps them, setting long text smaller where it must, and the plan says so in text_pt and note. ',
};
const HOLDS = {
  slideCount: 'With it the presentation has exactly that many slides, and no slide is added: a slide holds what its layout holds, and the plan says in holds_words and note what is cut.',
  plan: 'The planner keeps them, and the plan says in holds_words and note where a slide holds less. ',
};
const AS_ASKED = 'Do not split, shorten or sum up what the user asked for yourself: pass the number of slides and the amount of text as they said them. ';
const INSTRUCTIONS =
  'Everything else the user asked for, in their own words: audience, tone, every amount of text they ask for ("500 words on each slide", "a paragraph of about 80 words on pricing"), ' +
  'and every wish about how slides look or are arranged that has no argument of its own ("three columns", "no cover slide"). Pass a wish on even when it seems unusual or long.';
const GIVEN_TEXT = '. Text the user gives for a slide goes in whole and word for word, in quotation marks';
const ACCEPT_CUT = {
  type: 'boolean',
  description:
    "Only after the user accepted that a slide's text is cut, when an earlier call with this instruction answered with a note saying so. Left out, a change that would cut a slide's text is not made, and the answer's note says why.",
};
const NAME_DOCUMENTS = 'When the user wants their own text on the slides and the app lists it, name it in documents and do not copy it into topic: the planner reads it word for word. ';
// A short paste is not listed: its way to the slides.
const UNLISTED_TEXT = " A text of the user's for the slides that the app does not list (a short paste) goes here, whole and word for word, in quotation marks.";

/** The declarations of 0.6.1, taken from its tree. */
const BEFORE = JSON.parse(readFileSync(new URL('./fixtures/tools-0.6.1.json', import.meta.url), 'utf8')) as Record<string, ToolDefinition[]>;
/** Declarations of 0.6.1 with the texts of 0.6.2 put in their place, `accept_cut` added, and from 0.6.4 the plan shown in the user's language. */
function asNow(before: ToolDefinition[] | undefined, forms: typeof SMALLER): ToolDefinition[] {
  const after = structuredClone(before ?? []);
  const [plan, editPlan, , edit] = after;
  const planned = propertiesOf(plan);
  (planned.slide_count as Property).description = SLIDE_COUNT + forms.slideCount;
  (planned.instructions as Property).description = INSTRUCTIONS;
  (plan as ToolDefinition).description = (plan as ToolDefinition).description.replace('Show the plan to the user', `${AS_ASKED}${forms.plan}Show the plan to the user in the language they write in,`);
  for (const tool of [editPlan, edit]) {
    const instruction = propertiesOf(tool).instruction as Property;
    instruction.description = instruction.description.replace('figures they gave', `figures they gave${GIVEN_TEXT}`);
  }
  propertiesOf(edit).accept_cut = ACCEPT_CUT;
  return after;
}

test("without documents the tools are declared as in 0.6.1 but for the texts that pass the person's numbers on", () => {
  const fixture = presentationTools({ template, llm: noModel, store: memoryStore(), fields: { closing: 'Thank you' }, images: painter }).definitions;
  const onDemo = presentationTools({ template: demo, llm: noModel, store: memoryStore() }).definitions;
  assert.deepEqual(fixture, asNow(BEFORE.fixture, SMALLER));
  // The demo template offers free slides only with layouts adapted to their text: here a slide holds what its layout holds.
  assert.deepEqual(onDemo, asNow(BEFORE.demo, HOLDS));
  // The change to a presentation takes accept_cut last; the plan tool takes no documents.
  assert.equal(Object.keys(propertiesOf(onDemo[3])).at(-1), 'accept_cut');
  assert.ok(!('documents' in propertiesOf(onDemo[0])));
});

test('without free slides the tools promise no smaller text', () => {
  const off = presentationTools({ template: demo, llm: noModel, store: memoryStore(), free: false, colors: false }).definitions;
  assert.deepEqual(off, asNow(BEFORE['demo/free-off'], HOLDS));
  // Free slides on (the demo template needs its deck field for them), then off by the option or by the app's list of layouts.
  const promised = (more: Partial<PresentationToolsOptions>) => {
    const [plan] = presentationTools({ template: demo, llm: noModel, store: memoryStore(), layout: 'adaptive', fields: demoFields, ...more }).definitions;
    return [propertiesOf(plan).slide_count?.description, plan?.description.includes(AS_ASKED + SMALLER.plan), plan?.description.includes(AS_ASKED + HOLDS.plan), /set smaller/.test(JSON.stringify(plan))];
  };
  assert.deepEqual(promised({}), [SLIDE_COUNT + SMALLER.slideCount, true, false, true]);
  for (const more of [{ free: false }, { patterns: ['cover', 'text-only', 'closing'] }]) {
    assert.deepEqual(promised(more), [SLIDE_COUNT + HOLDS.slideCount, false, true, false], JSON.stringify(more));
  }
});

test('with documents the plan tool takes names', () => {
  const without = presentationTools({ template: demo, llm: noModel, store: memoryStore() }).definitions;
  const [plan, ...rest] = presentationTools({ template: demo, llm: noModel, store: memoryStore(), documents: true }).definitions;
  const properties = propertiesOf(plan);
  const before = propertiesOf(without[0]);
  assert.deepEqual(Object.keys(properties), ['topic', 'slide_count', 'language', 'image_density', 'text_size', 'text_only', 'instructions', 'documents', 'accent_color', 'layouts']);
  assert.deepEqual(properties.documents, {
    type: 'array',
    items: { type: 'string' },
    description:
      'The texts of this conversation the presentation is made from, by the names the app lists for them (a file name, "pasted text 1"). The slides carry every part of them in their own words and order. ' +
      'Name a text whenever the user wants it on the slides, and also when they want a presentation about it, a summary of it or some parts of it: say which in instructions. ' +
      'Leave it out when the presentation is not made from a text the user gave.',
  });
  assert.equal(properties.topic?.description, `${before.topic?.description} When documents are named, a sentence or two on what the presentation is for is enough: the planner reads the documents themselves.`);
  assert.equal(properties.language?.description, `${before.language?.description} When documents are named, the language of their text, unless the user asks for another.`);
  assert.equal(properties.instructions?.description, `${INSTRUCTIONS}${UNLISTED_TEXT}`);
  assert.equal(plan?.description, without[0]?.description.replace('Show the plan to the user', `${NAME_DOCUMENTS}Show the plan to the user`));
  assert.deepEqual((plan?.parameters as { required: string[] }).required, ['topic']);
  // The other tools take no documents: a new text means a new plan.
  assert.deepEqual(rest, without.slice(1));
});

// A text a person gave in the conversation: four paragraphs under two headings, as four passages.
const GARDEN = [
  'Background',
  'The community garden opened in 2018 on an empty lot behind the station. Forty families now grow vegetables there, and a school class visits every week. Water comes from a tank that collects rain from the roof of the old shed.',
  'Volunteers meet on Saturday mornings to weed the paths, fix the fences and share seeds. The garden lends its tools for free, and a board by the gate lists the tasks of the week.',
  '',
  'Plans',
  'The board wants to add twenty beds on the north side and a greenhouse for winter seedlings. Raised beds will let older members and people in wheelchairs garden without bending down.',
  'The work will cost about forty thousand euros. The town has promised half of it, and the rest will come from a plant sale in May and from local shops.',
].join('\n');
const conversation: ChatDocument[] = [
  { name: 'pasted text 1', text: 'Make two slides of my garden notes for the members meeting, with all of the text.' },
  { name: 'Garden notes.docx', text: GARDEN },
];
const freeText = { shows: [{ kind: 'text', count: 1 }], shows_main: 'none', side_by_side: 0 };
const gardenOutline = {
  title: 'The garden',
  left_out: [],
  asked_slides: 0,
  document_language: 'en',
  slides: [
    { pattern: 'free', purpose: 'Why the garden works', passages: [1, 2], ...freeText },
    { pattern: 'free', purpose: 'What comes next', passages: [3, 4], ...freeText },
  ],
};
/** A writer that gives each slide it is asked for a title and what else it must have. */
const titles = (request: JsonRequest): unknown => {
  type Slide = { properties: Record<string, { type?: string }>; required?: string[] };
  const slides = (request.schema as unknown as { properties: Record<string, Slide> }).properties;
  return Object.fromEntries(
    Object.entries(slides).map(([key, slide]) => [key, Object.fromEntries((slide.required ?? []).map((name) => [name, slide.properties[name]?.type === 'array' ? ['Point'] : name === 'title' ? `Slide ${key}` : 'Line']))]),
  );
};
const gardenTools = (...answers: unknown[]) => toolsOn(demo, answers, { fields: demoFields, layout: 'adaptive', documents: true });
const gardenPlan = { topic: 'The garden notes, for the members meeting', language: 'en', slide_count: 2 };

test('the plan tool reads named texts from the list on the call', async () => {
  const { tools, requests } = gardenTools(gardenOutline);
  const planned = await tools.call(TOOL_NAMES.plan, { ...gardenPlan, documents: ['Garden notes.docx'] }, { documents: conversation });
  assert.equal(planned.output.error, undefined, String(planned.output.error));
  assert.equal(planned.plan?.brief.document, GARDEN);
  // The planner reads the text itself: every passage is in its request.
  for (const passage of documentPassages(GARDEN)) assert.ok(requests[0]?.prompt.includes(passage.text), passage.text);
  // The chat model is told which passages each slide carries and how many words, not the text.
  const slides = planned.output.slides as Array<Record<string, unknown>>;
  assert.deepEqual(
    slides.map((slide) => [slide.passages, slide.text_words, slide.text_pt ?? null, slide.holds_words ?? null]),
    [
      ['p1-p2', 75, null, null],
      ['p3-p4', 61, null, null],
    ],
  );
  assert.deepEqual(planned.output.document, { passages: 4, words: 136 });
  assert.equal(planned.output.note, undefined);
  const said = JSON.stringify(planned.output);
  assert.ok(said.length < 1500, said);
  for (const sentence of GARDEN.split(/\n|(?<=\.) /)) assert.ok(sentence.length < 12 || !said.includes(sentence), sentence);

  // A name is found by case and by its beginning, without quotes, and a text named twice is carried once.
  for (const documents of ['garden notes.DOCX', 'garden notes', ['"Garden notes"', 'Garden notes.docx']]) {
    const again = gardenTools(gardenOutline);
    const found = await again.tools.call(TOOL_NAMES.plan, { ...gardenPlan, documents }, { documents: conversation });
    assert.equal(found.plan?.brief.document, GARDEN, JSON.stringify(documents));
  }
  // Tools built without the option pass the argument over.
  const off = toolsOn(demo, [gardenOutline], { fields: demoFields, layout: 'adaptive' });
  const plain = await off.tools.call(TOOL_NAMES.plan, { ...gardenPlan, documents: ['Garden notes.docx'] }, { documents: conversation });
  assert.equal(plain.plan?.brief.document, undefined);
});

test('every named text is carried, once, under any form of its name', async () => {
  const carried = async (documents: string[], list: ChatDocument[], answer: unknown = gardenOutline) => {
    const { tools } = gardenTools(answer);
    const result = await tools.call(TOOL_NAMES.plan, { ...gardenPlan, documents }, { documents: list });
    assert.equal(result.output.error, undefined, String(result.output.error));
    return result.plan?.brief.document;
  };
  // A file name in decomposed letters, as from a Mac, or with a direction mark, is found by the name a chat model gives back.
  const decomposed = 'Résumé café.docx'.normalize('NFD');
  for (const name of ['Résumé café.docx', 'RÉSUMÉ CAFÉ.DOCX', 'résumé']) assert.equal(await carried([name.normalize('NFC')], [{ name: decomposed, text: GARDEN }]), GARDEN, name);
  assert.equal(await carried(['סיכום.docx'], [{ name: '\u200Fסיכום.docx', text: GARDEN }]), GARDEN);
  // A file and the message it came in, both named, are one text: the message, which holds all of the file.
  const message = `Here are my notes:\r\n\r\n${GARDEN.replace(/\n/g, '\r\n')}`;
  assert.equal(await carried(['Garden notes.docx', 'pasted text 2'], [...conversation, { name: 'pasted text 2', text: message }]), message);
  // One text under two names is carried once.
  const twice: ChatDocument[] = [
    { name: 'a.txt', text: GARDEN },
    { name: 'b.txt', text: GARDEN.replace(/\n/g, '\r\n') },
  ];
  assert.equal(await carried(['a.txt', 'b.txt'], twice), GARDEN);
  // A plan that puts every passage of the joined texts on the two slides.
  const placing = (joined: string) => {
    const numbers = documentPassages(joined).map((passage) => passage.number);
    return {
      ...gardenOutline,
      slides: [
        { pattern: 'free', purpose: 'The first part', passages: numbers.slice(0, 1), ...freeText },
        { pattern: 'free', purpose: 'The rest', passages: numbers.slice(1), ...freeText },
      ],
    };
  };
  // Words of one text that are in another only in the same order, with others between them, are a text of its own.
  const apart: ChatDocument[] = [
    { name: 'a.txt', text: 'The garden opened in 2018.\nVolunteers meet on Saturdays.' },
    { name: 'b.txt', text: 'The garden opened in 2018.\nThe board wants twenty more beds.\nVolunteers meet on Saturdays.' },
  ];
  const both = apart.map((document) => document.text).join('\n\n');
  assert.equal(await carried(['a.txt', 'b.txt'], apart, placing(both)), both);
  // Eleven texts are all carried: what is limited is their characters, not their number.
  const notes = Array.from({ length: 11 }, (_, at) => ({ name: `note-${at + 1}.txt`, text: `Note ${at + 1}. The team in region ${at + 1} met its goal this quarter.` }));
  const joined = notes.map((note) => note.text).join('\n\n');
  assert.equal(
    await carried(
      notes.map((note) => note.name),
      notes,
      placing(joined),
    ),
    joined,
  );
});

test('an unknown, empty or too long text is an error the chat model can act on', async () => {
  const { tools, requests } = gardenTools();
  const plan = async (documents: unknown, list?: ChatDocument[]) => (await tools.call(TOOL_NAMES.plan, { topic: 'The garden', documents }, list ? { documents: list } : {})).output.error;
  // A short paste is not listed and is not to be named: the error says where it goes, and that a file not listed could not be read.
  const unlisted = 'A text the user pasted in a short message is not listed: pass it in instructions, whole and word for word, in quotation marks. A file that is not listed could not be read: tell the user.';
  assert.equal(
    await plan(['budget.xlsx'], conversation),
    `"documents" names a text the app does not have: "budget.xlsx". The texts the app lists for this conversation are "pasted text 1" and "Garden notes.docx": use those names. ${unlisted}`,
  );
  assert.equal(await plan(['a.docx', 'b.docx']), `"documents" names texts the app does not have: "a.docx" and "b.docx". The app lists no text for this conversation. ${unlisted}`);
  // A name is found by its own beginning only: "notes" begins no name here.
  assert.equal(String(await plan('notes', conversation)).startsWith('"documents" names a text the app does not have: "notes".'), true);
  assert.equal(await plan('notes', [{ name: 'notes', text: ' \n ' }]), 'The text "notes" is empty: tell the user it could not be read.');
  assert.equal(
    await plan(['long.txt'], [{ name: 'long.txt', text: 'The garden grows. '.repeat(700) }]),
    'The text is 12599 characters, and a presentation is made from at most 12000 (about 1,800 words). Ask the user to paste the part the presentation is about, and name that.',
  );
  // Two texts, each of half of that: two texts and not one under two names, which is carried once.
  assert.equal(
    await plan(['one', 'two'], [
      { name: 'one', text: 'Seeds are shared. '.repeat(340) },
      { name: 'two', text: 'Beds are watered. '.repeat(340) },
    ]),
    'The texts are 12240 characters, and a presentation is made from at most 12000 (about 1,800 words). Ask the user to paste the part the presentation is about, and name that.',
  );
  assert.equal(requests.length, 0, 'no model is asked');
});

const LIBRARY = readFileSync(new URL('./fixtures/library-he.txt', import.meta.url), 'utf8');
const libraryOutline = { title: 'הספרייה', slides: [{ pattern: 'content', purpose: 'למה לשפץ' }, { pattern: 'closing', purpose: 'תודה' }] };

test('a topic that copies a listed text is an error; one that retells it is a note', async () => {
  // A fictional Hebrew text of 416 words, listed as a long message of the person's.
  const listed: ChatDocument[] = [{ name: 'pasted text 1', text: LIBRARY }];
  const copied =
    '"topic" copies the text "pasted text 1". Name it in documents instead and keep topic to what the presentation is for: the planner reads the text word for word. ' +
    'If the user wants a presentation about it or a summary of it, name it all the same and say so in instructions.';
  const paragraphs = LIBRARY.split('\n').filter((line) => line.split(' ').length > 3);
  const summary =
    'הספרייה העירונית בשכונת הגפן נפתחה לפני עשרים שנה. בשנה שעברה ביקרו בה שמונים אלף איש. ועד השכונה הגיש לעירייה תוכנית שיפוץ. ' +
    'עלות השיפוץ שישה מיליון שקלים, מחציתה מהעירייה ומחציתה מתרומות. שלוש קרנות התחייבו לתרום מיליון וחצי. ' +
    'במהלך השיפוץ הספרייה תפעל במבנה זמני במתנ"ס. הסיכון העיקרי הוא עיכוב בהיתרי הבנייה. העבודות יימשכו כשנה.';
  const unrelated =
    'A presentation for the sales team about the new phone plans of the coming year: the prices, the speeds, what each plan includes, how customers move from an old plan to a new one, ' +
    'which shops will sell them first, how the staff will be trained in the weeks before the launch, and what the team should say when a customer asks about the old prices.';
  const { tools, requests } = toolsOn(template, [libraryOutline, libraryOutline, libraryOutline, libraryOutline], { fields: { closing: 'תודה' }, images: painter, documents: true });
  const plan = (topic: string, list = listed) => tools.call(TOOL_NAMES.plan, { topic, language: 'he' }, { documents: list });

  // A copy, whole or shortened, is an error before any model is asked.
  assert.deepEqual((await plan(LIBRARY)).output, { error: copied });
  assert.deepEqual((await plan(paragraphs.filter((_, at) => at % 2 === 0).join('\n'))).output, { error: copied });
  assert.equal(requests.length, 0);

  // A summary of its facts plans, and says that the slides will not carry the text.
  const retold = await plan(summary);
  assert.equal(retold.output.error, undefined, String(retold.output.error));
  assert.deepEqual(retold.output.not_on_slides, ['pasted text 1']);
  assert.equal(
    retold.output.next,
    'The topic retells "pasted text 1", and the slides will not carry that text word for word. If the user wanted their text on the slides, plan again and name it in documents. ' +
      `Show this plan to the user. Use edit_presentation_plan for changes and create_presentation once they approve. ${IN_THEIR_LANGUAGE}`,
  );
  // The app is given the same as a note, to show in the person's language.
  assert.deepEqual(retold.issues, [
    {
      severity: 'warning',
      message: 'the topic retells "pasted text 1", and the slides will not carry that text word for word',
      code: 'text-retold',
      values: { name: 'pasted text 1' },
    },
  ]);
  assert.equal(retold.plan?.brief.document, undefined);
  // A sentence on what the presentation is for, and a topic of its own, say neither.
  for (const topic of ['מצגת על תוכנית השיפוץ של הספרייה לישיבת הוועד', unrelated]) {
    const output = (await plan(topic)).output;
    assert.equal(output.error, undefined, String(output.error));
    assert.equal(output.not_on_slides, undefined, topic);
    assert.equal(output.next, 'Show this plan to the user. Use edit_presentation_plan for changes and create_presentation once they approve.');
  }
  // A listed message of 700 characters is most often the request itself: a topic that copies or retells it says neither.
  const message = LIBRARY.slice(0, 700);
  const short = await plan(message, [{ name: 'pasted text 1', text: message }]);
  assert.equal(short.output.error, undefined);
  assert.equal(short.output.not_on_slides, undefined);
  assert.equal(requests.length, 4);
});

test('a list on the call and no name gives the requests of a call without one', async () => {
  const args = { topic: 'A deck about dogs', language: 'en', slide_count: 2 };
  const options = { fields: { closing: 'Thank you' }, images: painter, thumbnail: (pattern: string) => `/thumbs/${pattern}.png` };
  const before = toolsOn(template, [outline], options);
  const listing = toolsOn(template, [outline, outline], { ...options, documents: true });
  const planned = await before.tools.call(TOOL_NAMES.plan, args);
  const listed = await listing.tools.call(TOOL_NAMES.plan, args, { documents: [{ name: 'pasted text 1', text: LIBRARY }] });
  const bare = await listing.tools.call(TOOL_NAMES.plan, args);
  assert.deepEqual(listing.requests, [...before.requests, ...before.requests]);
  assert.deepEqual(listed.output, planned.output);
  assert.deepEqual({ ...bare.output, plan_id: 'plan_1' }, planned.output);
  assert.deepEqual(listed.plan, planned.plan);
});

// A long message that only asks for a presentation, listed as a text.
const REQUEST = [
  'I need a presentation for the members meeting of our community garden next Saturday morning.',
  'About forty members will come, and most of them know the garden well, so please skip the history and keep the tone friendly and practical.',
  'Please open with a short cover slide with the name of the meeting and the date, and close with a slide that thanks everyone and says who to ask about the work days.',
  'In between I would like one slide on how the season went, one slide on the three things that went well, and one slide on the two things that did not go well and what we will do about them.',
  'After that please add a slide on the plans for the coming year, with no more than four points, and a slide on the help the board needs from the members.',
  'Keep every slide easy to read from the back of the hall, and avoid long paragraphs: short points work better for this group.',
  'If a slide would look better with a simple table, use one, but do not invent any numbers that I have not given you.',
  'I would also like the presentation to feel like one story from the beginning to the end, so the slides should lead into each other.',
  'Please make sure the slide on the plans stands out a little, because that is the one we will talk about the longest.',
  'One more thing: the hall has a wide screen, so a wide layout is fine.',
  'The meeting starts at ten and ends at noon, and the slides will be shown in the community hall next to the station.',
  'Please write the slides in plain English, since several members are still learning the language.',
  'Thank you very much, and tell me if anything in this request is not clear before you start.',
].join(' ');

test('a long request is planned: a copy of it names the argument, and named, it is taken as instructions', async () => {
  assert.ok(REQUEST.length >= 1500, 'long enough to be checked');
  const listed: ChatDocument[] = [{ name: 'pasted text 1', text: REQUEST }];
  const allOut = { ...gardenOutline, left_out: documentPassages(REQUEST).map((passage) => passage.number), slides: [{ pattern: 'cover', purpose: 'The garden' }, { pattern: 'text-only', purpose: 'The plans' }] };
  const plainOutline = { title: 'The garden', slides: [{ pattern: 'cover', purpose: 'The garden' }, { pattern: 'text-only', purpose: 'The plans' }] };
  const { tools, store, requests } = gardenTools(allOut, plainOutline);
  const args = { topic: 'The garden, for the members meeting', language: 'en' };

  // Passed on in instructions as the user said it, it copies the listed text, and the error says that instructions does.
  const copied = await tools.call(TOOL_NAMES.plan, { ...args, instructions: REQUEST }, { documents: listed });
  assert.equal(
    copied.output.error,
    '"instructions" copies the text "pasted text 1". Name it in documents instead and keep topic to what the presentation is for: the planner reads the text word for word. ' +
      'If the user wants a presentation about it or a summary of it, name it all the same and say so in instructions.',
  );
  assert.equal(requests.length, 0);

  // Named, the planner leaves all of it out. The plan is made again with it as instructions, as
  // a plan of the request in instructions is made with no list, and stored. Nothing sends the chat model back.
  const named = await tools.call(TOOL_NAMES.plan, { ...args, documents: ['pasted text 1'] }, { documents: listed });
  assert.equal(named.output.error, undefined, String(named.output.error));
  const direct = gardenTools(plainOutline);
  await direct.tools.call(TOOL_NAMES.plan, { ...args, instructions: REQUEST });
  assert.equal(requests.length, 2);
  assert.deepEqual(requests[1], direct.requests[0]);
  assert.equal(named.plan?.brief.document, undefined);
  assert.equal(named.plan?.brief.instructions, REQUEST);
  assert.deepEqual(store.keys(), [storeKeys.plan('plan_1')]);
  assert.equal(named.output.document, undefined);
  assert.equal(named.output.next, `"pasted text 1" only asks for a presentation, so the planner took it as instructions, and no slide carries it word for word. Show this plan to the user. Use edit_presentation_plan for changes and create_presentation once they approve.`);
});

test('create_presentation says how many words were kept', async () => {
  const { tools } = gardenTools(gardenOutline, titles);
  await tools.call(TOOL_NAMES.plan, { ...gardenPlan, documents: ['Garden notes.docx'] }, { documents: conversation });
  const created = await tools.call(TOOL_NAMES.create, { plan_id: 'plan_1' });
  assert.equal(created.output.error, undefined, String(created.output.error));
  // Each slide holds its passages, so code set them word for word.
  assert.deepEqual(created.output.document, { words: 136, kept_words: 136 });
  assert.equal(created.output.warnings, undefined);
  assert.deepEqual(
    created.deck?.slides.map((slide) => slide.passages),
    [
      [1, 2],
      [3, 4],
    ],
  );
  const said = JSON.stringify(created.output);
  assert.ok(said.length < 1500, said);
  // A presentation made from no text says nothing of one.
  const { tools: plain } = setup(outline, { s1: slideContent('Dogs are great', 'A dog'), s2: {} });
  await plain.call(TOOL_NAMES.plan, { topic: 'Dogs' });
  assert.equal((await plain.call(TOOL_NAMES.create, { plan_id: 'plan_1' })).output.document, undefined);
});

test('create_presentation says a text it translated is translated, and counts no words kept', async () => {
  const { tools } = gardenTools(gardenOutline, titles, titles);
  await tools.call(TOOL_NAMES.plan, { ...gardenPlan, language: 'fr', documents: ['Garden notes.docx'] }, { documents: conversation });
  const created = await tools.call(TOOL_NAMES.create, { plan_id: 'plan_1' });
  assert.equal(created.output.error, undefined, String(created.output.error));
  assert.deepEqual(created.output.document, { words: 136, translated: true });
});

test('describeDocuments lists the names with their words and first words', () => {
  const listed: ChatDocument[] = [
    { name: 'risk-review.docx', text: 'Quarterly review of the risks we carry this year and what we do about each of them, with owners and dates.' },
    { name: 'pasted text 1', text: 'Our plan for 2027 has three parts: growth, staff, costs and training.' },
  ];
  assert.equal(
    describeDocuments(listed),
    'Texts the user gave in this conversation, which the presentation tools can read by name: "risk-review.docx" (21 words, begins "Quarterly review of the risks we carry\u2026"), ' +
      '"pasted text 1" (12 words, begins "Our plan for 2027 has three parts:\u2026"). When the user wants a presentation of one of them, name it in "documents" of plan_presentation ' +
      'and say in "topic" in a sentence or two what the presentation is for: the planner reads the text itself and puts all of it on the slides, word for word. ' +
      'Do not copy or retell a listed text in "topic". Name it also when the user wants a summary of it or some parts of it, and say which in "instructions". ' +
      'A message that only asks for a presentation is not such a text: pass what it asks in "topic" and "instructions".',
  );
  assert.equal(describeDocuments([]), '');
  assert.equal(
    describeDocuments([
      { name: ' ', text: 'A text' },
      { name: 'blank', text: ' \n' },
    ]),
    '',
  );
  // A text too long to make a presentation of whole is listed all the same, so that the chat model can ask for the part it is about.
  assert.match(describeDocuments([{ name: 'book.pdf', text: 'The "garden" grows. '.repeat(700) }]), /: "book\.pdf" \(2100 words, too long to make a presentation of whole, begins "The 'garden' grows\. The 'garden' grows\.\u2026"\)\. /);
  // At most the last ten.
  const many = Array.from({ length: 12 }, (_, at) => ({ name: `pasted text ${at + 1}`, text: 'Seeds are shared.' }));
  assert.match(describeDocuments(many), /name: "pasted text 3" \(3 words, begins "Seeds are shared\.\u2026"\),.*"pasted text 12"/);
  assert.doesNotMatch(describeDocuments(many), /"pasted text 2"/);
});

const NEXT = 'Show this plan to the user. Use edit_presentation_plan for changes and create_presentation once they approve.';
/** An outline entry of a text slide with the words asked for it. */
const asked = (words: number, purpose = 'Keep the bike share') => ({ pattern: 'text-only', purpose, words, shows: [], shows_main: 'auto', side_by_side: 0 });
const bikes = (...answers: unknown[]) => toolsOn(demo, answers, { fields: demoFields, layout: 'adaptive' });

test('the plan says the size of a long text and what is cut', async () => {
  // Two slides of 500 words each: both set dense, well below the template's 18 pt.
  const two = bikes({ title: 'Bikes', slides: [asked(500), asked(500, 'Grow the bike share')] });
  const planned = await two.tools.call(TOOL_NAMES.plan, { topic: 'The bike share', language: 'en', slide_count: 2, instructions: '500 words on each slide' });
  assert.deepEqual(
    (planned.output.slides as Array<Record<string, unknown>>).map((slide) => [slide.words, slide.text_pt, slide.holds_words ?? null]),
    [
      [500, 10, null],
      [500, 10, null],
    ],
  );
  assert.equal(planned.output.note, "Slides 1 and 2 carry all their text in smaller type than the template's: 10 pt, where the template's is 18 pt. Tell the user, and offer more slides if they want it larger.");
  assert.deepEqual(Object.keys(planned.output), ['plan_id', 'title', 'language', 'slides', 'note', 'next']);
  assert.equal(planned.output.next, `${NEXT} ${IN_THEIR_LANGUAGE}`);
  // The app is given the note slide by slide, to show in the person's language.
  assert.deepEqual(
    planned.issues?.map((issue) => [issue.code, issue.values, issue.message, issue.slide]),
    [1, 2].map((slide) => ['slide-set-small', { slide, size: 10, templateSize: 18 }, `slide ${slide} carries all its text in smaller type than the template's: 10 pt, where the template's is 18 pt`, undefined]),
  );

  // 1000 words on one slide: cut, said before anything is written.
  const long = await bikes({ title: 'Bikes', slides: [asked(1000)] }).tools.call(TOOL_NAMES.plan, { topic: 'The bike share', language: 'en', slide_count: 1 });
  assert.deepEqual((long.output.slides as Array<Record<string, unknown>>)[0], {
    number: 1,
    layout: 'free',
    layout_description: 'A slide laid out for this presentation: a text.',
    purpose: 'Keep the bike share',
    words: 1000,
    text_pt: 9,
    holds_words: 725,
    free_layout: true,
  });
  assert.equal(
    long.output.note,
    'Not all the text fits: slide 1 holds about 725 of its 1000 words (text at 9 pt). The rest will be cut. Tell the user before creating the presentation, and offer more slides or a shorter text.',
  );
  assert.deepEqual(long.issues, [
    {
      severity: 'warning',
      message: 'slide 1 holds about 725 of its 1000 words (text at 9 pt): the rest will be cut',
      code: 'slide-holds-less',
      values: { slide: 1, words: 1000, holdsWords: 725, size: 9, freeOff: false },
    },
  ]);

  // 200 words are set at four fifths of the template's size or more: as in 0.6.1, nothing is said.
  const short = await bikes({ title: 'Bikes', slides: [asked(200)] }).tools.call(TOOL_NAMES.plan, { topic: 'The bike share', language: 'en', slide_count: 1 });
  const line = (short.output.slides as Array<Record<string, unknown>>)[0];
  assert.equal(line?.words, 200);
  assert.equal(line?.text_pt, undefined);
  assert.equal(short.output.note, undefined);

  // A plan without words is said as in 0.6.1.
  const plain = await bikes({ title: 'Bikes', slides: [{ pattern: 'text-only', purpose: 'Keep the bike share' }] }).tools.call(TOOL_NAMES.plan, { topic: 'The bike share', language: 'en', slide_count: 1 });
  assert.deepEqual(plain.output, {
    plan_id: 'plan_1',
    title: 'Bikes',
    language: 'en',
    slides: [{ number: 1, layout: 'text-only', layout_description: (plain.output.slides as Array<Record<string, unknown>>)[0]?.layout_description, purpose: 'Keep the bike share' }],
    next: NEXT,
  });
});

test('the plan asks for slide_count when words came without one', async () => {
  const hint = 'No slide_count was given, so the planner chose 2 slides. If the user named a number of slides, plan again with slide_count. ';
  const words = { title: 'Bikes', slides: [asked(500), asked(500, 'Grow the bike share')] };
  const open = await bikes(words).tools.call(TOOL_NAMES.plan, { topic: 'The bike share', language: 'en', instructions: '2 slides of 500 words' });
  // Two slides of 500 words are set small, which the note says: next ends with the language to say it in.
  assert.equal(open.output.next, `${hint}${NEXT} ${IN_THEIR_LANGUAGE}`);
  const counted = await bikes(words).tools.call(TOOL_NAMES.plan, { topic: 'The bike share', language: 'en', slide_count: 2 });
  assert.equal(counted.output.next, `${NEXT} ${IN_THEIR_LANGUAGE}`);
  // Without words the planner's own number is as it was.
  const plain = await bikes({ title: 'Bikes', slides: [{ pattern: 'text-only', purpose: 'Keep the bike share' }] }).tools.call(TOOL_NAMES.plan, { topic: 'The bike share', language: 'en' });
  assert.equal(plain.output.next, NEXT);
  // A change to a plan gets no hint.
  const { tools } = bikes(words, { summary: 'Done.', language: 'en', ...words });
  await tools.call(TOOL_NAMES.plan, { topic: 'The bike share', language: 'en' });
  const changed = await tools.call(TOOL_NAMES.editPlan, { plan_id: 'plan_1', instruction: 'A better title' });
  assert.equal(changed.output.next, `${NEXT} ${IN_THEIR_LANGUAGE}`);
  assert.deepEqual(changed.issues?.map((issue) => issue.code), ['slide-set-small', 'slide-set-small']);
});

test('edit_presentation_plan keeps the mark of parts a change asked for, and shows the model none of it', async () => {
  const text = { shows: [{ kind: 'text', count: 1 }], shows_main: 'auto', side_by_side: 0 };
  const chart = { shows: [{ kind: 'text', count: 1 }, { kind: 'chart', count: 1 }], shows_main: 'auto', side_by_side: 0 };
  const plan = { title: 'Bikes', slides: [{ pattern: 'free', purpose: 'Growth', ...text }, { pattern: 'free', purpose: 'Asks', ...text }] };
  const change = { summary: 'Added a chart.', language: 'en', title: 'Bikes', slides: [{ pattern: 'free', purpose: 'Growth', ...chart, words: 300 }, { pattern: 'free', purpose: 'Asks', ...text }] };
  const { tools, store } = bikes(plan, change);
  await tools.call(TOOL_NAMES.plan, { topic: 'The bike share', language: 'en', slide_count: 2 });
  const changed = await tools.call(TOOL_NAMES.editPlan, { plan_id: 'plan_1', instruction: 'Add a chart beside the text of slide 1, and make it 300 words' });
  assert.ok(!JSON.stringify(changed.output).includes('partsAsked'), JSON.stringify(changed.output));
  const stored = JSON.parse(new TextDecoder().decode(await store.get(storeKeys.plan(String(changed.output.plan_id))))) as { slides: Array<{ blocks?: string[]; partsAsked?: true }> };
  assert.deepEqual([stored.slides[0]?.blocks, stored.slides[0]?.partsAsked, stored.slides[1]?.partsAsked], [['text', 'chart'], true, undefined]);
});

test('edit_presentation says a cut before writing', async () => {
  const deck: DeckSpec = { version: 1, language: 'en', fields: demoFields, title: 'Bikes', slides: [{ pattern: 'text-only', slots: { title: 'Keep it', body: ['Riders rely on it.'] } }] };
  const change = {
    summary: 'Longer.',
    language: 'en',
    title: 'Bikes',
    slides: [{ from: 1, pattern: 'text-only', change: 'Make the text 1500 words', images: false, words: 1500, parts: [], shows_main: 'auto', side_by_side: 0 }],
    text_size: 'normal',
    fields: demoFields,
  };
  const paragraph = 'Riders rely on the service every day, and the city gains from it in ways that are easy to see. '.repeat(4).trim();
  const { tools, store, requests } = bikes(change, change, { s1: { title: 'Keep it', b1Text: Array.from({ length: 12 }, () => paragraph) } });
  await store.set(storeKeys.deck('deck_0'), new TextEncoder().encode(JSON.stringify(deck)));
  const stopped = await tools.call(TOOL_NAMES.edit, { presentation_id: 'deck_0', instruction: 'Make slide 1 1500 words' });
  assert.deepEqual(stopped.output, {
    note: 'Nothing was changed: slide 1 would hold about 725 of its 1500 words (text at 9 pt). Ask for more slides or a shorter text, or accept that the rest is cut.',
    next:
      'Nothing was written. Tell the user what the note says, and ask them, in their language, whether they accept that the rest is cut. ' +
      `If they accept, call edit_presentation again with the same instruction and accept_cut: true; otherwise change the request (more slides, a shorter text). ${IN_THEIR_LANGUAGE}`,
  });
  // The app is given the cut slide by slide, from the error's fits, to show in the person's language.
  assert.deepEqual(stopped.issues, [
    {
      severity: 'warning',
      message: 'slide 1 would hold about 725 of its 1500 words (text at 9 pt), so the change was not made',
      code: 'change-would-cut',
      values: { slide: 1, words: 1500, holdsWords: 725, size: 9, freeOff: false },
    },
  ]);
  assert.equal(requests.length, 1, 'the plan call alone');
  assert.deepEqual(store.keys(), [storeKeys.deck('deck_0')], 'nothing is stored');

  // Accepted, the change is made: the slide is cut to what it holds, and the presentation says so.
  const accepted = await tools.call(TOOL_NAMES.edit, { presentation_id: 'deck_0', instruction: 'Make slide 1 1500 words', accept_cut: true });
  assert.equal(accepted.output.error, undefined, String(accepted.output.error));
  assert.equal(accepted.output.presentation_id, 'deck_1');
  assert.equal(requests.length, 3);
  assert.deepEqual(
    (accepted.output.warnings as string[]).filter((warning) => /words were asked/.test(warning)),
    ['slide 1: 1500 words were asked for this slide, and it holds about 725 even with its text set small, at 9 pt: the rest is cut. Ask for the text over more slides, or for a shorter one'],
  );
  // Each line is an issue the app can show in the person's language, and the chat model is told to speak theirs.
  assert.deepEqual(accepted.issues?.map((issue) => formatIssue(issue, 'en', { deck: accepted.deck })), accepted.output.warnings);
  assert.ok(
    accepted.issues?.some((issue) => formatIssue(issue, 'he', { deck: accepted.deck }) === "שקף 1: לשקף הזה התבקשו 1500 מילים, והוא מכיל כ-725 מהן גם בגופן קטן של 9 נק': יתר הטקסט נחתך. אפשר לבקש את הטקסט על פני יותר שקפים, או טקסט קצר יותר"),
  );
  assert.equal(accepted.output.next, IN_THEIR_LANGUAGE);
});

test("own-text.md's lines for the chat's system prompt agree with what the tools say", () => {
  const prompt = readFileSync(new URL('../docs/prompts/own-text.md', import.meta.url), 'utf8');
  // The lines are the block of the prompt that begins with the line on the user's own text.
  const all = prompt.split('\n');
  const start = all.findIndex((line) => line.startsWith('- When the user wants their own text on the slides'));
  const end = all.indexOf('```', start);
  assert.ok(start > 0 && all[start - 1] === '```' && end > start, 'own-text.md has the lines');
  const lines = all.slice(start, end);
  // Every name in quotes is an argument of the tools or a field of what they answer.
  const [plan, , , edit] = presentationTools({ template: demo, llm: noModel, store: memoryStore(), documents: true }).definitions;
  const known = new Set([...Object.keys(propertiesOf(plan)), ...Object.keys(propertiesOf(edit)), 'words', 'holds_words', 'note', 'not_on_slides', 'warnings']);
  for (const [, name = ''] of lines.join(' ').matchAll(/"([a-z_]+)"/g)) assert.ok(known.has(name), name);
  // A slide with holds_words was planned with the user's numbers, and planning again gives the same cut: the user is told.
  const again = lines.filter((line) => line.includes('plan again'));
  assert.equal(again.length, 1);
  assert.match(again[0] as string, /A slide with "holds_words" is not such a case: tell the user what "note" says before you create the presentation\./);
  // A short paste is never listed and cannot be named: it has a way of its own, as the plan tool's error for a name not listed says.
  assert.ok(lines.includes('- A shorter text the user pasted is not listed: put it in "instructions", whole and word for word, in quotation marks.'));
});
