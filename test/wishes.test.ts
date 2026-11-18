// What a person asks about the form of a deck: which layouts it may use, its colours, how many
// slides it has. Such wishes come before the planner's habits, code sets what code can set, and a
// plan or a deck without them is planned as it was before. No test calls a live model.
import assert from 'node:assert/strict';
import { readFileSync } from 'node:fs';
import { dirname, join } from 'node:path';
import { test } from 'node:test';
import { fileURLToPath } from 'node:url';
import {
  PlanError,
  PptxPackage,
  TOOL_NAMES,
  composeLayouts,
  editPlan,
  loadTemplate,
  memoryStore,
  planEdit,
  planOutline,
  presentationTools,
  recolorDeck,
  storeKeys,
  validateDeck,
  writeDeck,
} from '../src/index.ts';
import type { Brief, DeckPlan, DeckSpec, JsonRequest, LlmProvider, PlanSettings, PresentationToolsOptions, SlideSpec, TemplateManifest } from '../src/index.ts';
import { readPresentation } from '../src/pptx/presentation.ts';
import { buildFixturePptx, cardsManifest } from './fixture.ts';
import { withoutCodes } from './issues.ts';

const pptx = buildFixturePptx();
// Three patterns: `content` (with a photo), `closing` (its text is a deck field) and `cards`.
const template = loadTemplate(pptx, cardsManifest);
const fields = { closing: 'Thank you' };
const brief: Brief = { prompt: 'A deck about dogs', language: 'en', fields };

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

type Schema = { properties: Record<string, Schema>; required?: string[]; enum?: string[]; items?: Schema; minItems?: number; maxItems?: number; description?: string; type?: string };
const schemaOf = (request: JsonRequest | undefined) => request?.schema as unknown as Schema;
const entryOf = (request: JsonRequest | undefined) => schemaOf(request).properties.slides?.items as Schema;
const offered = (request: JsonRequest | undefined) => entryOf(request).properties.pattern?.enum;
const promptOf = (request: JsonRequest | undefined) => request?.prompt ?? '';
/** The rules of a planning prompt: its last block, one rule to a line. */
const rulesOf = (request: JsonRequest | undefined) => (promptOf(request).split('\n\n').at(-1) ?? '').split('\n');

const WISHES = (asked: string) =>
  `- What the ${asked} asks about the form of the deck comes first, before variety, before the usual choices below and before what a pattern's description advises: ` +
  'how many slides, which kinds of slides, how many things on a slide, how many side by side. Do exactly that, even when it means one slide or one pattern for every slide.';
const OPENING = '- A deck of three or more slides opens with an opening pattern and ends with a closing pattern when the list has them, unless the brief asks for a deck without them.';
const VARY = '- Otherwise vary the patterns across the deck.';
const FREE_PARTS_RULE =
  '- A slide with the layout free names its parts in "shows", in reading order. A number of things the brief asks to see on one slide (twelve charts, forty logos) is that number of parts, however many.';
const limited = (ids: string) => `- This deck was limited to these layouts: ${ids}. Keep to them unless the requested change asks for another.`;

const cardsOutline = { title: 'Dogs', slides: [{ pattern: 'cards', purpose: 'Two things', items: 2 }] };
const twoCards = (title = 'Two'): SlideSpec => ({ pattern: 'cards', items: 2, slots: { title, card1Text: 'One', card2Text: 'Two' } });
const cardsContent = { title: 'Two', card1Text: 'One', card2Text: 'Two' };
const threeSlides: DeckSpec = { version: 1, title: 'Dogs', brief: 'A deck about dogs', language: 'en', direction: 'ltr', fields, slides: [twoCards(), twoCards('More'), { pattern: 'closing', slots: {} }] };
/** The answer of a model that plans a change to `threeSlides`: where each slide comes from, and what changes on it. */
const editAnswer = (slides: Array<[from: number, pattern: string, change?: string]>) => ({
  summary: 'Done.',
  language: 'en',
  title: 'Dogs',
  text_size: 'normal',
  fields,
  slides: slides.map(([from, pattern, change = '']) => ({ from, pattern, change, images: false })),
});

// --- which layouts a deck may use ---------------------------------------------------------

test('a list of layouts in the brief is all the outline is offered, and the app\'s own list narrows it', async () => {
  // The app's list alone, as before: the same prompt, with fewer patterns in it.
  const app = scripted(cardsOutline);
  await planOutline(template, { ...brief, slideCount: 4 }, { llm: app.llm, patterns: ['cards', 'closing'] });
  assert.deepEqual(offered(app.requests[0]), ['closing', 'cards']);
  assert.ok(rulesOf(app.requests[0]).includes(VARY), 'the app\'s list changes no rule');
  assert.match(promptOf(app.requests[0]), /Unless the brief says how many pictures it wants: use a pattern with a photo on about half of the slides\./);

  // The person's list: nothing else is in the list of patterns or in the schema.
  const person = scripted(cardsOutline);
  const plan = await planOutline(template, { ...brief, slideCount: 4, patterns: ['cards', 'content'] }, { llm: person.llm });
  assert.deepEqual(offered(person.requests[0]), ['content', 'cards']);
  assert.match(promptOf(person.requests[0]), /\n- content: /);
  assert.doesNotMatch(promptOf(person.requests[0]), /\n- closing: /);
  assert.deepEqual(plan.brief.patterns, ['cards', 'content'], 'the list stays with the plan as it was given');
  // Variety and the share of photos would pull against the list. The opening and closing rule stays: it speaks of the list.
  const rules = rulesOf(person.requests[0]);
  assert.ok(!rules.includes(VARY));
  assert.doesNotMatch(rules.join('\n'), /photo on|patterns with a photo/);
  assert.ok(rules.includes(OPENING));
  assert.ok(rules.some((rule) => rule.startsWith('- A wish about the kind of slides (text only, no tables, only cards) decides which patterns may be used')));

  // Both lists: what is in both.
  const both = scripted(cardsOutline);
  await planOutline(template, { ...brief, patterns: ['cards', 'content'] }, { llm: both.llm, patterns: ['closing', 'cards'] });
  assert.deepEqual(offered(both.requests[0]), ['cards']);

  // An empty list is no list: every pattern is on offer, and free slides with them.
  const none = scripted(cardsOutline);
  await planOutline(template, { ...brief, patterns: [] }, { llm: none.llm });
  assert.deepEqual(offered(none.requests[0]), ['content', 'closing', 'cards', 'free']);
  assert.ok(rulesOf(none.requests[0]).includes(VARY));
});

test('a list that names a layout the template lacks, or shares nothing with the app\'s, is an error that says so', async () => {
  const unknown = scripted();
  await assert.rejects(planOutline(template, { ...brief, patterns: ['cards', 'made-up'] }, { llm: unknown.llm }), (error: unknown) => {
    assert.ok(error instanceof PlanError);
    assert.equal(error.message, 'Template "fixture" has no pattern called "made-up", so the deck cannot be limited to it. Its patterns are: content, closing, cards');
    return true;
  });
  await assert.rejects(planOutline(template, { ...brief, patterns: ['x', 'y'] }, { llm: unknown.llm }), /has no pattern called "x", "y", so the deck cannot be limited to them\./);

  const apart = scripted();
  await assert.rejects(planOutline(template, { ...brief, patterns: ['content'] }, { llm: apart.llm, patterns: ['closing', 'cards'] }), (error: unknown) => {
    assert.ok(error instanceof PlanError);
    assert.equal(error.message, 'The deck is limited to the patterns content, and the options allow only closing, cards: no pattern is in both lists');
    return true;
  });
  assert.equal(unknown.requests.length + apart.requests.length, 0, 'no model is asked');
});

test('a text-only wish gives way to a list of layouts: the list names them outright', async () => {
  const { llm, requests } = scripted(cardsOutline);
  await planOutline(template, { ...brief, textOnly: true, patterns: ['cards', 'closing'] }, { llm });
  assert.deepEqual(offered(requests[0]), ['closing', 'cards'], 'cards are no plain text, and the person named them');
  assert.doesNotMatch(promptOf(requests[0]), /every pattern in the list is plain text/);
});

test('a stored list that names a layout the template has lost is read without it', async () => {
  const stored: DeckPlan = {
    version: 1,
    title: 'Dogs',
    brief: { ...brief, patterns: ['cards', 'removed-since'] },
    slides: [{ pattern: 'cards', purpose: 'Two things', items: 2 }, { pattern: 'closing', purpose: 'Thanks' }],
  };
  // The deck is written, the slide outside the list too, and the deck keeps what is left of the list.
  const written = await writeDeck(template, stored, { llm: scripted({ s1: cardsContent, s2: {} }).llm });
  assert.deepEqual(written.deck.slides.map((slide) => slide.pattern), ['cards', 'closing']);
  assert.deepEqual(written.deck.patterns, ['cards']);

  // With nothing left, the deck is planned and written as if it had no list.
  const gone: DeckPlan = { ...stored, brief: { ...brief, patterns: ['removed-since'] } };
  const plain = await writeDeck(template, gone, { llm: scripted({ s1: cardsContent, s2: {} }).llm });
  assert.ok(!('patterns' in plain.deck));

  const answer = { summary: 'Done.', language: 'en', title: 'Dogs', slides: stored.slides };
  const change = scripted(answer);
  await editPlan(template, stored, 'Add a slide', { llm: change.llm });
  assert.ok(rulesOf(change.requests[0]).includes(limited('cards')));
  const free = scripted(answer);
  const kept = await editPlan(template, gone, 'Add a slide', { llm: free.llm });
  assert.doesNotMatch(promptOf(free.requests[0]), /was limited to these layouts/);
  assert.match(promptOf(free.requests[0]), /of the slides the change adds, about half use a pattern with a photo/, 'planned as a deck without a list');
  assert.deepEqual(kept.plan.brief.patterns, ['removed-since'], 'the plan itself is not rewritten');

  // A deck too: its own list is a wish, and what the template has lost is passed over.
  const deck: DeckSpec = { version: 1, language: 'en', direction: 'ltr', fields, patterns: ['cards', 'removed-since'], slides: [twoCards(), { pattern: 'closing', slots: {} }] };
  const keep = { summary: 'Done.', language: 'en', title: '', text_size: 'normal', fields, slides: [{ from: 1, pattern: 'cards', change: '', images: false }, { from: 2, pattern: 'closing', change: '', images: false }] };
  const edit = scripted(keep);
  const result = await planEdit(template, deck, 'Nothing to change', { llm: edit.llm });
  assert.match(promptOf(edit.requests[0]), /\n- This deck was limited to these layouts: cards\. Keep to them unless the requested change asks for another\.\n/);
  assert.deepEqual(result.deck.patterns, deck.patterns, 'an edit keeps the deck\'s list as it is');
  const lost = scripted(keep);
  await planEdit(template, { ...deck, patterns: ['removed-since'] }, 'Nothing to change', { llm: lost.llm });
  assert.doesNotMatch(promptOf(lost.requests[0]), /was limited to these layouts/);
});

test('a list of layouts is a wish when a plan or a deck is changed: everything stays on offer, with one rule', async () => {
  const plan: DeckPlan = { version: 1, title: 'Dogs', brief: { ...brief, patterns: ['cards'] }, slides: [{ pattern: 'cards', purpose: 'Two things', items: 2 }] };
  const answer = { summary: 'Added a closing slide.', language: 'en', title: 'Dogs', slides: [...plan.slides, { pattern: 'closing', purpose: 'Thanks' }] };
  const { llm, requests } = scripted(answer);
  const changed = await editPlan(template, plan, 'End with a closing slide', { llm });
  assert.deepEqual(offered(requests[0]), ['content', 'closing', 'cards']);
  assert.ok(rulesOf(requests[0]).includes(limited('cards')));
  assert.deepEqual(changed.plan.slides.map((slide) => slide.pattern), ['cards', 'closing'], 'the requested change asked for another layout, and got it');
  assert.deepEqual(changed.plan.brief.patterns, ['cards']);
  // The slide outside the list is written like any other, and the list goes to the deck.
  const written = await writeDeck(template, changed.plan, { llm: scripted({ s1: cardsContent, s2: {} }).llm });
  assert.deepEqual(written.deck.patterns, ['cards']);
  assert.equal(written.deck.slides.length, 2);
});

test('no layout is made for a deck that is limited to a list, whatever composeLayouts says', async () => {
  const root = join(dirname(fileURLToPath(import.meta.url)), '..');
  const demo = loadTemplate(readFileSync(join(root, 'templates', 'demo.pptx')), JSON.parse(readFileSync(join(root, 'templates', 'demo.manifest.json'), 'utf8')) as TemplateManifest);
  const outline = {
    title: 'The year ahead',
    slides: [
      { pattern: 'cover', purpose: 'Open' },
      { pattern: 'text-only', purpose: 'Where we stand' },
      { pattern: 'text-only', purpose: 'Our three goals' },
      { pattern: 'text-only', purpose: 'What it costs' },
      { pattern: 'closing', purpose: 'Thanks' },
    ],
  };
  const model = () => {
    const requests: JsonRequest[] = [];
    const llm: LlmProvider = { generateJson: async (request) => (requests.push(request), request.system.startsWith('You review the plan') ? { layouts: [] } : outline) };
    return { llm, requests };
  };
  const year: Brief = { prompt: 'The year ahead', slideCount: 5, language: 'en', fields: { organizationDepartment: 'Strategy office' } };

  // Layouts are made where free slides are not on offer, so this deck is planned without them, as before there were any.
  const open = model();
  await planOutline(demo, year, { llm: open.llm, layout: 'adaptive', composeLayouts: 2, free: false });
  assert.equal(open.requests.length, 2, 'the outline, then the look for slides that want a new layout');

  const closed = model();
  const plan = await planOutline(demo, { ...year, patterns: ['cover', 'text-only', 'closing'] }, { llm: closed.llm, layout: 'adaptive', composeLayouts: 2 });
  assert.equal(closed.requests.length, 1, 'the outline alone');
  assert.deepEqual(offered(closed.requests[0]), ['cover', 'text-only', 'closing'], 'a list that does not name free slides offers none');
  assert.equal(plan.layouts, undefined);
  // The step on its own makes nothing for such a plan either.
  assert.deepEqual(await composeLayouts(demo, plan, { llm: closed.llm, layout: 'adaptive' }), { plan, made: [], failed: [] });
  assert.equal(closed.requests.length, 1);
});

// --- wishes come first --------------------------------------------------------------------

test('what the brief asks about the form of the deck is the first rule, and the habits give way to it', async () => {
  const { llm, requests } = scripted(cardsOutline);
  await planOutline(template, brief, { llm });
  const rules = rulesOf(requests[0]);
  assert.equal(rules[0], WISHES('brief'));
  // Free slides are on offer, and the rule about their parts follows the wishes: a number of things asked for on one slide is that number of parts.
  assert.equal(rules[1], FREE_PARTS_RULE);
  assert.equal(rules[2], OPENING);
  assert.equal(rules[3], '- For a pattern with a row of items, set "items" to the number the brief asks for, or else the number the content really has. Do not pad a list to fill a row.');
  assert.ok(rules.includes(VARY));
  assert.equal(rules.at(-1), '- Unless the brief says how many pictures it wants: use a pattern with a photo on about half of the slides (a free slide with a photo counts).');
  // With free slides turned off the rules are those of before: the opening rule second, and the share of photos as it was.
  const before = scripted(cardsOutline);
  await planOutline(template, brief, { llm: before.llm, free: false });
  assert.equal(rulesOf(before.requests[0])[1], OPENING);
  assert.equal(rulesOf(before.requests[0]).at(-1), '- Unless the brief says how many pictures it wants: use a pattern with a photo on about half of the slides.');
  assert.doesNotMatch(JSON.stringify(before.requests[0]), /free|shows/);
  // The usual length is as it was, in the prompt and in the schema.
  assert.match(promptOf(requests[0]), /\n\nPlan a deck of 6 to 10 slides\. Choose one pattern for each slide/);
  const slides = schemaOf(requests[0]).properties.slides as Schema;
  assert.deepEqual([slides.minItems, slides.maxItems], [6, 10]);

  for (const [density, sentence] of [
    ['low', '- Unless the brief says how many pictures it wants: use a pattern with a photo on at most a quarter of the slides (a free slide with a photo counts).'],
    ['high', '- Unless the brief says how many pictures it wants: prefer patterns with a photo wherever they suit the content (a free slide with a photo counts).'],
    ['none', '- Do not plan any photos.'],
  ] as const) {
    const asked = scripted(cardsOutline);
    await planOutline(template, { ...brief, imageDensity: density }, { llm: asked.llm });
    assert.equal(rulesOf(asked.requests[0]).at(-1), sentence);
  }
});

test('a deck of one slide is asked for as one slide, and is not told to open and close', async () => {
  const one = scripted(cardsOutline);
  await planOutline(template, { ...brief, slideCount: 1 }, { llm: one.llm });
  assert.match(promptOf(one.requests[0]), /\n\nPlan a deck of exactly 1 slide\. Choose one pattern for each slide/);
  assert.doesNotMatch(promptOf(one.requests[0]), /opening pattern|closing pattern/);
  assert.equal(rulesOf(one.requests[0])[0], WISHES('brief'));

  const two = scripted(cardsOutline);
  await planOutline(template, { ...brief, slideCount: { min: 1, max: 2 } }, { llm: two.llm });
  assert.match(promptOf(two.requests[0]), /Plan a deck of 1 to 2 slides\./);
  assert.ok(!rulesOf(two.requests[0]).includes(OPENING));

  const three = scripted(cardsOutline);
  await planOutline(template, { ...brief, slideCount: 3 }, { llm: three.llm });
  assert.match(promptOf(three.requests[0]), /Plan a deck of exactly 3 slides\./);
  assert.ok(rulesOf(three.requests[0]).includes(OPENING));
});

test('the writer is told it places nothing, and how much to write', async () => {
  const { llm, requests } = scripted(cardsOutline, { s1: cardsContent });
  await writeDeck(template, await planOutline(template, { ...brief, slideCount: 1 }, { llm }), { llm });
  for (const request of requests) {
    assert.ok(request.system.startsWith("You write presentation decks for a fixed slide template. You do not place or size anything: you choose from the template's slide patterns and write the content for their slots.\n"));
    assert.ok(request.system.includes('\n- Character limits are hard and count spaces. Stay under them. Where a slot says how much to write, write about that much; elsewhere keep slide text short: a phrase beats a sentence.\n'));
  }
});

test('a change to a plan puts the request first, asks for the same entry as a new plan, and speaks of photos for added slides only', async () => {
  const outline = scripted(cardsOutline);
  const plan = await planOutline(template, { ...brief, slideCount: 1 }, { llm: outline.llm });
  const answer = { summary: 'Done.', language: 'en', title: 'Dogs', slides: plan.slides };

  const change = scripted(answer);
  await editPlan(template, plan, 'Add a slide', { llm: change.llm });
  const rules = rulesOf(change.requests[0]);
  assert.equal(rules[0], WISHES('request'));
  assert.equal(rules[1], '- Change only what the request asks for. Keep every other slide, its pattern, its purpose and its place.');
  // Free slides are on offer here too: a slide that has parts keeps them unless the request is about them.
  assert.equal(rules[2], '- A slide with the layout free names its parts in "shows". Repeat them exactly as the plan has them unless the request changes what that slide shows.');
  assert.ok(rules.includes('- Colours are set by the software, not by you: a wish about colour changes no slide.'));
  assert.equal(rules.at(-1), '- Unless the request says how many pictures it wants: of the slides the change adds, about half use a pattern with a photo (a free slide with a photo counts).');
  assert.deepEqual(entryOf(change.requests[0]), entryOf(outline.requests[0]), 'one schema for a slide of a plan');
  assert.deepEqual(entryOf(change.requests[0]).required, ['pattern', 'purpose', 'shows', 'shows_main', 'side_by_side']);

  // The other shares of photos speak of added slides too. "None" is the sentence it always was.
  for (const [density, sentence] of [
    ['low', '- Unless the request says how many pictures it wants: of the slides the change adds, at most a quarter use a pattern with a photo (a free slide with a photo counts).'],
    ['high', '- Unless the request says how many pictures it wants: for the slides the change adds, prefer patterns with a photo wherever they suit the content (a free slide with a photo counts).'],
    ['none', '- Do not plan any photos.'],
  ] as const) {
    const asked = scripted(answer);
    await editPlan(template, { ...plan, brief: { ...plan.brief, imageDensity: density } }, 'Add a slide', { llm: asked.llm });
    assert.equal(rulesOf(asked.requests[0]).at(-1), sentence);
  }

  // A plan of one or two slides is not told to open and close when it is changed, as it was not when it was made.
  for (const slides of [plan.slides, [...plan.slides, ...plan.slides]]) {
    const small = scripted({ ...answer, slides });
    await editPlan(template, { ...plan, slides }, 'A better title', { llm: small.llm });
    assert.doesNotMatch(promptOf(small.requests[0]), /opening pattern|closing pattern/);
  }

  // Nothing about photos for a deck of text, nor for one limited to a list of layouts.
  const textOnly = scripted({ ...answer, slides: [{ pattern: 'closing', purpose: 'Thanks' }] });
  await editPlan(template, { ...plan, brief: { ...plan.brief, textOnly: true } }, 'Add a slide', { llm: textOnly.llm });
  assert.doesNotMatch(rulesOf(textOnly.requests[0]).join('\n'), /photo/);
  assert.match(rulesOf(textOnly.requests[0]).join('\n'), /This deck was asked to be text only/);
  const listed = scripted(answer);
  await editPlan(template, { ...plan, brief: { ...plan.brief, patterns: ['cards'] } }, 'Add a slide', { llm: listed.llm });
  assert.doesNotMatch(rulesOf(listed.requests[0]).join('\n'), /photo/);
});

// --- colours and the list, kept with the plan and the deck -----------------------------------

const oldPlan = (): DeckPlan =>
  // A plan as 0.5.2 stored it: no key that came later.
  JSON.parse(JSON.stringify({ version: 1, title: 'Dogs', brief, slides: [{ pattern: 'cards', purpose: 'Two things', items: 2 }, { pattern: 'closing', purpose: 'Thanks' }] })) as DeckPlan;
const oldDeck: DeckSpec = {
  version: 1,
  language: 'en',
  direction: 'ltr',
  brief: 'A deck about dogs',
  slides: [twoCards(), { pattern: 'closing', slots: {} }],
  title: 'Dogs',
  fields,
};

test('writeDeck copies the brief\'s colours and list to the deck, and adds no key when there are none', async () => {
  const before = await writeDeck(template, oldPlan(), { llm: scripted({ s1: cardsContent, s2: {} }).llm });
  assert.deepEqual(before.deck, oldDeck, 'a plan from before gives the deck it always gave');
  assert.deepEqual(Object.keys(before.deck), ['version', 'language', 'direction', 'brief', 'slides', 'title', 'fields']);

  const plan = oldPlan();
  plan.brief = { ...plan.brief, colors: { accent: '#1e6b3a' }, patterns: ['cards', 'closing', 'cards'] };
  const { llm, requests } = scripted({ s1: cardsContent, s2: {} });
  const after = await writeDeck(template, plan, { llm });
  assert.deepEqual(after.deck, { ...oldDeck, colors: { accent: '1E6B3A' }, patterns: ['cards', 'closing'] });
  assert.deepEqual(after.warnings, []);
  assert.equal(requests.length, 1, 'code copies them: no model is asked about a colour');
  assert.doesNotMatch(JSON.stringify(requests[0]), /1e6b3a/i, 'and none is told one');
});

test('editPlan writes colours and the list to the brief when the caller sets them, and removes them on null or an empty list', async () => {
  const plan = oldPlan();
  const answer = { summary: 'Done.', language: 'en', title: 'Dogs', slides: plan.slides };

  // The person's list has a name of its own, `deckPatterns`: `patterns` in the same options is the app's list, as it was.
  const settings: PlanSettings = { colors: { accent: '#1e6b3a' }, deckPatterns: ['cards', 'closing'] };
  const set = scripted(answer);
  const first = await editPlan(template, plan, 'Only cards and the closing slide from now on, in dark green', { llm: set.llm, ...settings });
  assert.deepEqual(first.plan.brief.colors, { accent: '1E6B3A' });
  assert.deepEqual(first.plan.brief.patterns, ['cards', 'closing']);
  assert.ok(rulesOf(set.requests[0]).includes(limited('cards, closing')), 'the model is told of the list set with this change');
  assert.deepEqual(offered(set.requests[0]), ['content', 'closing', 'cards']);
  assert.ok(!('colors' in plan.brief) && !('patterns' in plan.brief), 'the plan given is not changed');

  // A later change that sets neither keeps both.
  const later = scripted(answer);
  const second = await editPlan(template, first.plan, 'Add a slide', { llm: later.llm });
  assert.deepEqual([second.plan.brief.colors, second.plan.brief.patterns], [{ accent: '1E6B3A' }, ['cards', 'closing']]);

  const lift = scripted(answer, answer);
  const third = await editPlan(template, second.plan, 'Any layout again, in the usual colours', { llm: lift.llm, colors: null, deckPatterns: [] });
  assert.ok(!('colors' in third.plan.brief) && !('patterns' in third.plan.brief));
  assert.doesNotMatch(promptOf(lift.requests[0]), /was limited to these layouts/);
  const fourth = await editPlan(template, second.plan, 'Any layout again', { llm: lift.llm, deckPatterns: null });
  assert.ok(!('patterns' in fourth.plan.brief));
  assert.deepEqual(fourth.plan.brief.colors, { accent: '1E6B3A' });

  // No id is an error when a plan is changed: only a new brief is held to its ids. One the template
  // does not have is dropped, and a list that names none of its patterns sets nothing.
  const partly = scripted(answer);
  const some = await editPlan(template, plan, 'Only cards and tables', { llm: partly.llm, deckPatterns: ['cards', 'tables'] });
  assert.deepEqual(some.plan.brief.patterns, ['cards']);
  assert.ok(rulesOf(partly.requests[0]).includes(limited('cards')));
  const slip = scripted(answer);
  const unchanged = await editPlan(template, first.plan, 'Only tables', { llm: slip.llm, deckPatterns: ['tables'] });
  assert.deepEqual(unchanged.plan.brief.patterns, ['cards', 'closing'], 'a slip never lifts a limit');
  assert.ok(rulesOf(slip.requests[0]).includes(limited('cards, closing')));
  // Nor is a list that shares nothing with the app's: it is kept as the person's wish, and no limit is in force.
  const apart = scripted(answer);
  const wished = await editPlan(template, plan, 'Only the photo layout', { llm: apart.llm, patterns: ['cards', 'closing'], deckPatterns: ['content'] });
  assert.deepEqual(wished.plan.brief.patterns, ['content']);
  assert.doesNotMatch(promptOf(apart.requests[0]), /was limited to these layouts/);

  // The app's own list is another thing: it narrows what is offered, and is not kept.
  const app = scripted({ ...answer, slides: [plan.slides[0]] });
  const narrowed = await editPlan(template, plan, 'Drop the last slide', { llm: app.llm, patterns: ['cards'] });
  assert.deepEqual(offered(app.requests[0]), ['cards']);
  assert.ok(!('patterns' in narrowed.plan.brief));
});

test('a model that plans a change is told whether the software set a colour with it', async () => {
  const set = '- The software sets the colour that was asked for with this change: say nothing about colour in summary.';
  const unset = '- No colour was given with this change: if the request asks for one, say in summary that the colour was not changed.';
  const plan = oldPlan();
  const answer = { summary: 'Done.', language: 'en', title: 'Dogs', slides: plan.slides };
  const keep = editAnswer([[1, 'cards'], [2, 'cards'], [3, 'closing']]);

  // Without a colour the summary must not claim one: the chat model reads it, and can then call again with a colour.
  const plain = scripted(answer, keep);
  await editPlan(template, plan, 'Make it dark green', { llm: plain.llm });
  await planEdit(template, threeSlides, 'Make it dark green', { llm: plain.llm });
  for (const request of plain.requests) {
    assert.ok(promptOf(request).split('\n').includes(unset));
    assert.ok(!promptOf(request).split('\n').includes(set));
  }

  // With one, also the one that returns to the template's colours, the model leaves colour out of its summary.
  const given = scripted(answer, answer, keep, keep);
  await editPlan(template, plan, 'Make it dark green', { llm: given.llm, colors: { accent: '1E6B3A' } });
  await editPlan(template, plan, 'The usual colours', { llm: given.llm, colors: null });
  await planEdit(template, threeSlides, 'Make it dark green', { llm: given.llm, colors: { accent: '1E6B3A' } });
  await planEdit(template, threeSlides, 'The usual colours', { llm: given.llm, colors: null });
  for (const request of given.requests) {
    assert.ok(promptOf(request).split('\n').includes(set));
    assert.ok(!promptOf(request).split('\n').includes(unset));
    assert.doesNotMatch(JSON.stringify(request), /1E6B3A/, 'the model is not told the colour itself');
  }
  // The note stands right after the rule that colours are not the model's work.
  const lines = promptOf(given.requests[2]).split('\n');
  assert.equal(lines[lines.indexOf(set) - 1], '- Colours are set by the software, not by you: for a wish about colour leave change empty.');
});

test('the code of a colour that was given is no figure of the brief', async () => {
  const root = join(dirname(fileURLToPath(import.meta.url)), '..');
  const demo = loadTemplate(readFileSync(join(root, 'templates', 'demo.pptx')), JSON.parse(readFileSync(join(root, 'templates', 'demo.manifest.json'), 'utf8')) as TemplateManifest);
  const slides = [{ pattern: 'cover', purpose: 'Open' }, { pattern: 'text-only', purpose: 'Why teams work' }, { pattern: 'closing', purpose: 'Thanks' }];
  const plan: DeckPlan = { version: 1, title: 'Teamwork', brief: { prompt: 'A deck about teamwork', language: 'en', fields: { organizationDepartment: 'Office' } }, slides };
  const numbers = (request: JsonRequest | undefined) => (offered(request) ?? []).filter((id) => id === 'stat-cards' || id === 'chart-native');
  const ask = async (words: string, colors?: { accent: string }) => {
    const model = scripted({ summary: 'Done.', language: 'en', title: 'Teamwork', slides });
    await editPlan(demo, plan, words, { llm: model.llm, layout: 'adaptive', ...(colors ? { colors } : {}) });
    return numbers(model.requests[0]);
  };
  assert.deepEqual(await ask('Make it #1e6b3a', { accent: '1E6B3A' }), [], 'the colour named in the request brings no number layouts');
  assert.deepEqual(await ask('Make it #abc', { accent: 'AABBCC' }), []);
  // A number that is no colour of the deck still counts, and so does a code when no colour was given.
  assert.deepEqual(await ask('Make it 1E6B3A and add our 3 sites', { accent: '1E6B3A' }), ['stat-cards', 'chart-native']);
  assert.deepEqual(await ask('Make it 1E6B3A'), ['stat-cards', 'chart-native'], 'as before when the caller sets no colour');
});

test('recolorDeck sets and removes colours by code, on the deck or on some slides', () => {
  const deck: DeckSpec = { ...oldDeck, slides: [twoCards(), twoCards('More'), { pattern: 'closing', slots: {} }] };
  const frozen = JSON.stringify(deck);

  const green = recolorDeck(deck, { accent: '#1e6b3a' });
  assert.deepEqual(green, { ...deck, colors: { accent: '1E6B3A' } });
  // Some slides only: the deck's own colours stay as they are.
  const red = recolorDeck(green, { accent: 'C62828' }, [2, 9]);
  assert.deepEqual(red.colors, { accent: '1E6B3A' });
  assert.deepEqual(red.slides.map((slide) => slide.colors), [undefined, { accent: 'C62828' }, undefined]);
  assert.ok(!('colors' in (red.slides[0] as SlideSpec)), 'no key on a slide that has none');
  // null removes: a slide follows the deck again, a deck its template.
  assert.deepEqual(recolorDeck(red, null, [2]), green);
  assert.deepEqual(recolorDeck(green, null), deck);
  // The deck's colours and a slide's own are set apart: one does not touch the other.
  assert.deepEqual(recolorDeck(red, { accent: '1565C0' }), { ...red, colors: { accent: '1565C0' } });
  assert.deepEqual(recolorDeck(red, null).slides, red.slides);
  assert.ok(!('colors' in recolorDeck(red, null)));
  // Colours that ask for nothing are no colours.
  assert.deepEqual(recolorDeck(green, {}), deck);
  // The other colours a deck can state go the same way.
  assert.deepEqual(recolorDeck(deck, { accent: '1565c0', accent2: '#F4A259', replace: { '0e7c7b': '#1565c0' } }).colors, { accent: '1565C0', accent2: 'F4A259', replace: { '0E7C7B': '1565C0' } });
  assert.equal(JSON.stringify(deck), frozen, 'the deck given is not changed');
});

test('an edit is told that colours and look are not wording, and what change is for', async () => {
  const { llm, requests } = scripted(editAnswer([[1, 'cards'], [2, 'cards'], [3, 'closing']]));
  const result = await planEdit(template, { ...threeSlides, patterns: ['cards', 'closing'] }, 'Make it look more modern, and blue', { llm });
  const prompt = promptOf(requests[0]);
  assert.match(prompt, /Return the slides of the deck after the change, in their new order\.\n- What the request asks about the form of the deck comes first, before variety,/);
  assert.ok(prompt.includes(`\n${WISHES('request')}\n- from: `));
  assert.ok(prompt.includes(`\n${limited('cards, closing')}\n`), 'in the order the person gave them');
  assert.ok(prompt.includes('\n- change: what to write or change in the content of that slide. Leave it empty for every slide whose content the request leaves alone.\n'));
  assert.ok(prompt.includes('\n- Colours are set by the software, not by you: for a wish about colour leave change empty.\n'));
  assert.ok(
    prompt.includes(
      '\n- A wish about how the deck looks (more modern, less empty, less text) is met with another pattern for a slide, another number of items, or the text size. It changes layouts and settings, not wording: leave change empty unless the slide gets another pattern.\n' +
        // A person who asks for other words on a slide still gets them: "shorten slide 3" is no wish about look.
        '- A request to shorten, cut or rewrite the text of a slide is a change of its content: say it in change.\n',
    ),
  );
  const entry = entryOf(requests[0]);
  assert.equal(
    entry.properties.change?.description,
    'What to write or change in the content of this slide: its text, numbers, pictures or parts. Empty when the content stays as it is; a new colour or a new place in the deck is not a change of content.',
  );
  assert.deepEqual(entry.properties.pattern?.enum, ['content', 'closing', 'cards'], 'the list is a wish here: every pattern stays on offer');
  assert.equal(requests.length, 1, 'nothing was rewritten');
  assert.deepEqual(result.deck.slides, threeSlides.slides);
  assert.deepEqual(result.deck.patterns, ['cards', 'closing']);
});

test('an edit sets colours and the list by code, and a slide keeps its own colours when it is moved or rewritten', async () => {
  const own = { accent: 'C62828' };
  const deck: DeckSpec = { ...threeSlides, slides: [{ ...twoCards(), colors: own }, twoCards('More'), { pattern: 'closing', slots: {} }] };

  // Slide 2 of the deck as it is gets the colour; the edit moves it to the end.
  const move = scripted(editAnswer([[3, 'closing'], [1, 'cards'], [2, 'cards']]));
  const moved = await planEdit(template, deck, 'Put the closing slide first', { llm: move.llm, colors: { accent: '#1565c0' }, colorSlides: [2], deckPatterns: ['cards', 'closing'] });
  assert.equal(move.requests.length, 1, 'a colour needs no writing');
  assert.deepEqual(moved.rewritten, []);
  assert.deepEqual(moved.deck.slides.map((slide) => slide.pattern), ['closing', 'cards', 'cards']);
  assert.deepEqual(moved.deck.slides.map((slide) => slide.colors), [undefined, own, { accent: '1565C0' }]);
  assert.ok(!('colors' in moved.deck), 'the colour went to the slide, not to the deck');
  assert.deepEqual(moved.deck.patterns, ['cards', 'closing']);

  // A slide that is written again keeps the colours it had.
  const rewrite = scripted(editAnswer([[1, 'cards', 'A shorter title'], [2, 'cards'], [3, 'closing']]), { s1: { ...cardsContent, title: 'Two!' } });
  const rewritten = await planEdit(template, deck, 'Shorten the title of slide 1', { llm: rewrite.llm });
  assert.deepEqual(rewritten.rewritten, [1]);
  assert.deepEqual(rewritten.deck.slides[0], { ...twoCards('Two!'), colors: own });
  assert.ok(!('colors' in rewritten.deck) && !('patterns' in rewritten.deck), 'nothing is set that the caller did not set');

  // Without slides named the colours are the deck's, and null takes them away again. A slide's own stay.
  const keep = () => scripted(editAnswer([[1, 'cards'], [2, 'cards'], [3, 'closing']]));
  const blue = await planEdit(template, deck, 'Blue, please', { llm: keep().llm, colors: { accent: '1565C0' } });
  assert.deepEqual(blue.deck.colors, { accent: '1565C0' });
  assert.deepEqual(blue.deck.slides, deck.slides);
  const back = await planEdit(template, { ...blue.deck, patterns: ['cards'] }, 'The usual colours again', { llm: keep().llm, colors: null, deckPatterns: null });
  assert.ok(!('colors' in back.deck) && !('patterns' in back.deck));
  assert.deepEqual(back.deck.slides[0]?.colors, own);

  // The app's own list still narrows what a new slide may use, and is not written to the deck.
  const app = keep();
  const narrowed = await planEdit(template, threeSlides, 'Nothing to change', { llm: app.llm, patterns: ['cards'] });
  assert.deepEqual(entryOf(app.requests[0]).properties.pattern?.enum, ['cards', 'closing'], 'what is offered, and what the deck already uses');
  assert.ok(!('patterns' in narrowed.deck));

  // No id is an error when a deck is changed: one the template does not have is dropped, and a
  // list that names none of its patterns sets nothing, so that a slip never lifts a limit.
  const partly = keep();
  const some = await planEdit(template, threeSlides, 'Only cards and tables', { llm: partly.llm, deckPatterns: ['cards', 'tables'] });
  assert.deepEqual(some.deck.patterns, ['cards']);
  assert.match(promptOf(partly.requests[0]), /\n- This deck was limited to these layouts: cards\./);
  const slip = keep();
  const unchanged = await planEdit(template, { ...threeSlides, patterns: ['cards', 'closing'] }, 'Only tables', { llm: slip.llm, deckPatterns: ['tables'] });
  assert.deepEqual(unchanged.deck.patterns, ['cards', 'closing']);
  assert.match(promptOf(slip.requests[0]), /\n- This deck was limited to these layouts: cards, closing\./);
});

test('a list that shares nothing with the app\'s is kept as the wish it was, and limits nothing', async () => {
  // Stored with the plan or the deck, such a list fails nothing: the app's list alone decides what is offered.
  const plan: DeckPlan = { ...oldPlan(), brief: { ...brief, patterns: ['content'] } };
  const answer = { summary: 'Done.', language: 'en', title: 'Dogs', slides: plan.slides };
  const app = { patterns: ['closing', 'cards'] };

  const change = scripted(answer);
  const changed = await editPlan(template, plan, 'Add a slide', { llm: change.llm, ...app });
  assert.deepEqual(offered(change.requests[0]), ['closing', 'cards']);
  assert.doesNotMatch(promptOf(change.requests[0]), /was limited to these layouts/);
  assert.deepEqual(changed.plan.brief.patterns, ['content']);

  const written = await writeDeck(template, plan, { llm: scripted({ s1: cardsContent, s2: {} }).llm, ...app });
  assert.deepEqual(written.deck.patterns, ['content'], 'the deck keeps the person\'s list, not the app\'s');

  const edit = scripted(editAnswer([[1, 'cards'], [2, 'cards'], [3, 'closing']]));
  const edited = await planEdit(template, { ...threeSlides, patterns: ['content'] }, 'Nothing to change', { llm: edit.llm, ...app });
  assert.doesNotMatch(promptOf(edit.requests[0]), /was limited to these layouts/);
  assert.deepEqual(edited.deck.patterns, ['content']);

  // The check of a deck does not look at the list: it is a wish for planning, not a rule for a deck.
  for (const patterns of [['made-up'], ['content'], []]) {
    assert.deepEqual(validateDeck(template.manifest, { ...threeSlides, patterns }, { layout: 'adaptive' }), validateDeck(template.manifest, threeSlides, { layout: 'adaptive' }));
  }
});

test('colours for some slides go to those slides alone: none to a slide the edit adds, and a removed slide is named', async () => {
  const red = { accent: 'C62828' };
  const added = () => scripted(editAnswer([[0, 'cards', 'A new slide'], [1, 'cards'], [2, 'cards'], [3, 'closing']]), { s1: { ...cardsContent, card3Text: 'Three' } });

  // 0 is how the plan of an edit marks a new slide. It is no slide number, and colours nothing.
  const zero = await planEdit(template, threeSlides, 'Add a slide first', { llm: added().llm, colors: red, colorSlides: [0, 2] });
  assert.deepEqual(zero.deck.slides.map((slide) => slide.colors), [undefined, undefined, red, undefined]);
  assert.deepEqual(zero.recolored, [3], 'the number the slide has now');
  assert.deepEqual(zero.warnings, []);
  assert.ok(!('colors' in zero.deck));

  // An empty list names no slide: the colours are the deck's, as if the list were left out. The tool reads it so too.
  const none = await planEdit(template, threeSlides, 'Add a slide first', { llm: added().llm, colors: red, colorSlides: [] });
  assert.deepEqual(none.deck.colors, red);
  assert.deepEqual(none.deck.slides.map((slide) => slide.colors), [undefined, undefined, undefined, undefined]);
  assert.ok(!('recolored' in none));
  assert.deepEqual(recolorDeck(threeSlides, red, []), { ...threeSlides, colors: red });

  // A slide the edit removes cannot be coloured. The deck is not coloured in its place, and a warning says so.
  const drop = scripted(editAnswer([[1, 'cards'], [3, 'closing']]));
  const dropped = await planEdit(template, threeSlides, 'Remove slide 2', { llm: drop.llm, colors: red, colorSlides: [2] });
  assert.ok(!('colors' in dropped.deck));
  assert.deepEqual(dropped.deck.slides.map((slide) => slide.colors), [undefined, undefined]);
  assert.deepEqual(dropped.recolored, []);
  assert.deepEqual(withoutCodes(dropped.warnings), [{ severity: 'warning', message: 'the change removed slide 2, which was to be recoloured' }]);
  assert.doesNotMatch(dropped.warnings[0]?.message ?? '', /characters|paragraphs|lines/, 'not a message a rewrite would be started for');

  // An edit that sets no colours for slides has no such key in its result.
  const plain = await planEdit(template, threeSlides, 'Remove slide 2', { llm: scripted(editAnswer([[1, 'cards'], [3, 'closing']])).llm });
  assert.deepEqual(Object.keys(plain), ['deck', 'summary', 'rewritten', 'warnings']);
});

test('a colour that is no colour is reported, also when nothing is written', async () => {
  const keep = () => scripted(editAnswer([[1, 'cards'], [2, 'cards'], [3, 'closing']]));
  const said = (deck: DeckSpec) => validateDeck(template.manifest, deck).filter((issue) => issue.slide === undefined);

  // An edit that rewrites no slide asks no second model, and still says what the check of the deck says.
  const edited = await planEdit(template, threeSlides, 'Blue, please', { llm: keep().llm, colors: { accent: 'blue' } });
  assert.deepEqual(edited.deck.colors, { accent: 'blue' });
  assert.equal(edited.warnings.length, 1);
  assert.deepEqual(edited.warnings, said(edited.deck));
  assert.match(edited.warnings[0]?.message ?? '', /"blue", which is not six hex digits/);
  // A deck that states nothing new gets no warning from such an edit, as before.
  assert.deepEqual((await planEdit(template, threeSlides, 'Nothing to change', { llm: keep().llm })).warnings, []);

  // Colours that are no set of colours at all are kept as given too, so that the same check names them.
  const plan = oldPlan();
  plan.brief = { ...plan.brief, colors: 'blue' as never };
  const written = await writeDeck(template, plan, { llm: scripted({ s1: cardsContent, s2: {} }).llm });
  assert.equal(written.deck.colors, 'blue');
  assert.deepEqual(written.warnings.map((issue) => issue.message), ['"colors" must be an object such as { "accent": "1565C0" }; it is ignored']);
  const wrong = recolorDeck({ ...threeSlides, colors: { accent: 'C62828' } }, '1E6B3A' as never);
  assert.equal(wrong.colors, '1E6B3A', 'not taken for a wish to remove the colours');
  assert.equal(said(wrong).length, 1);
});

// --- the chat tools ---------------------------------------------------------------------------

const encoder = new TextEncoder();
const decoder = new TextDecoder();
type Definition = { name: string; description: string; parameters: { properties: Record<string, Schema>; required: string[] } };
const definitionsOf = (tools: { definitions: unknown[] }) => tools.definitions as Definition[];
const toolNamed = (tools: { definitions: unknown[] }, name: string) => definitionsOf(tools).find((tool) => tool.name === name) as Definition;

/** The tools over a store that already holds a presentation, `deck_0`. */
async function toolsWith(deck: DeckSpec, answers: unknown[] = [], options: Partial<PresentationToolsOptions> = {}) {
  const model = scripted(...answers);
  const store = memoryStore();
  await store.set(storeKeys.deck('deck_0'), encoder.encode(JSON.stringify(deck)));
  let next = 0;
  const tools = presentationTools({ template, llm: model.llm, store, fields, newId: (kind) => `${kind}_${++next}`, ...options });
  const saved = async (id: string) => JSON.parse(decoder.decode(await store.get(storeKeys.deck(id)))) as DeckSpec;
  return { ...model, store, tools, saved };
}

/** What an app must supply for every layout of the fixture to be usable: the deck field of `closing`, and photos for `content`. */
const everything = { fields, images: { generateImage: async () => new Uint8Array([1]) } };
/** How free slides are named among the layouts, where the tools can plan them. */
const FREE_LAYOUT = '- free: A slide laid out by the software from the parts asked for.';

test('the tools are declared for their template: four tools, each with its arguments, the layouts by id and description', () => {
  const tools = presentationTools({ template, llm: scripted().llm, store: memoryStore(), ...everything });
  assert.deepEqual(tools.definitions.map((tool) => tool.name), Object.values(TOOL_NAMES));
  for (const tool of definitionsOf(tools)) {
    assert.equal(typeof tool.parameters.properties, 'object');
    assert.ok(Object.keys(tool.parameters.properties).length > 0);
    for (const name of tool.parameters.required) assert.ok(name in tool.parameters.properties, `${tool.name} requires ${name}`);
  }
  assert.deepEqual(Object.keys(toolNamed(tools, TOOL_NAMES.plan).parameters.properties), ['topic', 'slide_count', 'language', 'image_density', 'text_size', 'text_only', 'instructions', 'accent_color', 'layouts']);
  assert.deepEqual(Object.keys(toolNamed(tools, TOOL_NAMES.editPlan).parameters.properties), ['plan_id', 'instruction', 'accent_color', 'layouts']);
  assert.deepEqual(Object.keys(toolNamed(tools, TOOL_NAMES.create).parameters.properties), ['plan_id']);
  // A change that would cut a slide's text is made only once the person accepts it: accept_cut comes last.
  assert.deepEqual(Object.keys(toolNamed(tools, TOOL_NAMES.edit).parameters.properties), ['presentation_id', 'instruction', 'accent_color', 'color_slides', 'layouts', 'accept_cut']);
  assert.deepEqual(toolNamed(tools, TOOL_NAMES.edit).parameters.required, ['presentation_id']);

  // Each layout is named with the first sentence of its description, so that an id is not all a chat model has.
  // Free slides can be planned on this template, so `free` stands beside its layouts as one more to name.
  const layouts = toolNamed(tools, TOOL_NAMES.plan).parameters.properties.layouts as Schema;
  assert.deepEqual(layouts.items?.enum, ['content', 'closing', 'cards', 'free']);
  assert.equal(
    layouts.description,
    'Only when the user limits the layouts. Every slide then uses one of them, so name the opening and closing layouts too unless the user wants none. The layouts:\n' +
      `- content: Title, body, photo and icon.\n- closing: Closing words.\n- cards: A row of cards.\n${FREE_LAYOUT}`,
  );
  for (const name of [TOOL_NAMES.editPlan, TOOL_NAMES.edit]) {
    const changed = toolNamed(tools, name).parameters.properties.layouts as Schema;
    assert.deepEqual(changed.items?.enum, ['content', 'closing', 'cards', 'free']);
    // A new list is for what is planned from now on: the slides that are there stay unless the person says more.
    assert.equal(
      changed.description,
      'Only when the user changes which layouts may be used: the only layouts allowed from now on. They are described under plan_presentation. ' +
        'Slides that are there already keep their layout unless the instruction says otherwise. An empty list lifts the limit.',
    );
  }
  // The planner cannot take a number of slides from the topic, so the chat model is told to pass it,
  // and that the planner keeps it: long text is set smaller on a free slide of text, never spread.
  assert.equal(
    toolNamed(tools, TOOL_NAMES.plan).parameters.properties.slide_count?.description,
    'Number of slides in all. Set it whenever the user names a number of slides ("2 slides", "a five-slide deck"): it is not read from the topic or the instructions. ' +
      'Leave it out when they name none: the planner then chooses the number. With it the presentation has exactly that many slides: long text is set smaller to fit, and no slide is added.',
  );
  assert.match(toolNamed(tools, TOOL_NAMES.plan).parameters.properties.text_only?.description ?? '', /or charts\. Not together with layouts: the layouts named then decide\.$/);
  assert.match(
    toolNamed(tools, TOOL_NAMES.editPlan).parameters.properties.instruction?.description ?? '',
    /figures they gave\. Text the user gives for a slide goes in whole and word for word, in quotation marks\. Leave it empty when only the colour changes$/,
  );
  // A colour alone changes a plan too, so the instruction is required of neither tool that changes something.
  assert.deepEqual(toolNamed(tools, TOOL_NAMES.editPlan).parameters.required, ['plan_id']);
  assert.equal(
    toolNamed(tools, TOOL_NAMES.plan).parameters.properties.accent_color?.description,
    "Only when the user asks for a colour: the main colour as six hex digits (RRGGBB), in place of the template's. For a colour given by name choose a shade that suits a presentation (dark green 1E6B3A, not 00FF00).",
  );
  assert.match(toolNamed(tools, TOOL_NAMES.edit).parameters.properties.accent_color?.description ?? '', /"template" returns to the template's colours\.$/);
  assert.equal(
    toolNamed(tools, TOOL_NAMES.edit).parameters.properties.color_slides?.description,
    'The slides to colour, by number, when the user wants the colour on some slides only. Left out, the whole presentation takes it.',
  );
  assert.equal(
    toolNamed(tools, TOOL_NAMES.plan).parameters.properties.instructions?.description,
    'Everything else the user asked for, in their own words: audience, tone, every amount of text they ask for ("500 words on each slide", "a paragraph of about 80 words on pricing"), ' +
      'and every wish about how slides look or are arranged that has no argument of its own ("three columns", "no cover slide"). Pass a wish on even when it seems unusual or long.',
  );
  // The descriptions say what can be asked for.
  assert.match(toolNamed(tools, TOOL_NAMES.plan).description, /Pass on every wish the user has about how it looks or is built: in slide_count, layouts, text_only, text_size, image_density and accent_color where one fits/);
  assert.match(toolNamed(tools, TOOL_NAMES.editPlan).description, /change the layout of a slide or what it shows, the length, the language or the colour\./);
  assert.match(toolNamed(tools, TOOL_NAMES.edit).description, /change the colour of the presentation or of one slide\./);

  // What a chat model's API refuses, or reads badly, is kept out.
  const declared = JSON.stringify(tools.definitions);
  for (const word of ['anyOf', 'oneOf', '"pattern"', '"format"', '"enum":[]']) assert.ok(!declared.includes(word), word);
  for (const tool of definitionsOf(tools)) {
    for (const name of ['layouts', 'color_slides']) {
      const list = tool.parameters.properties[name];
      if (list) assert.ok(!('minItems' in list) && !('maxItems' in list));
    }
  }
  // Built for each set of tools, so that one template's layouts are never offered for another.
  assert.notEqual(presentationTools({ template, llm: scripted().llm, store: memoryStore(), ...everything }).definitions, tools.definitions);
});

/** The fixture with other descriptions for its three layouts, in the order content, closing, cards. */
const described = (...descriptions: [string, string, string]): TemplateManifest => ({
  ...cardsManifest,
  patterns: cardsManifest.patterns.map((pattern, index) => ({ ...pattern, description: descriptions[index] as string })),
});
const layoutLinesOf = (manifest: TemplateManifest) => {
  const tools = presentationTools({ template: loadTemplate(pptx, manifest), llm: scripted().llm, store: memoryStore(), ...everything });
  return (toolNamed(tools, TOOL_NAMES.plan).parameters.properties.layouts?.description ?? '').split('\n').slice(1);
};

test('a long description is cut to its first sentence, and to 80 characters', () => {
  const lines = layoutLinesOf(
    described('Title, body, photo and icon.', 'Closing words. Use it last.', 'A row of cards, each with a heading and a short text that says what the card is about, for features or steps. Use it often.'),
  );
  // The line for free slides follows the template's layouts, in words of its own.
  assert.deepEqual(lines, ['- content: Title, body, photo and icon.', '- closing: Closing words.', '- cards: A row of cards, each with a heading and a short text that says what the card...', FREE_LAYOUT]);
  assert.ok((lines[2] as string).length <= '- cards: '.length + 80);

  // A full stop inside "e.g." or "i.e." ends no sentence.
  assert.deepEqual(layoutLinesOf(described('Use for e.g. a list of names. Then more.', 'Closing words, i.e. a thank you. Last.', 'A row of cards.')), [
    '- content: Use for e.g. a list of names.',
    '- closing: Closing words, i.e. a thank you.',
    '- cards: A row of cards.',
    FREE_LAYOUT,
  ]);
});

test('two layouts that the cut would leave with one line are told apart', () => {
  // A layout and its mirror differ in their last words, which the cut takes away.
  const mirrored = 'A title, an optional subtitle and one or two paragraphs of running text, beside a photo on the';
  const lines = layoutLinesOf(described(`${mirrored} left. Use it for a story.`, `${mirrored} right. Use it for a story.`, 'A row of cards.'));
  assert.deepEqual(lines, [`- content: ${mirrored} left.`, `- closing: ${mirrored} right.`, '- cards: A row of cards.', FREE_LAYOUT], 'those two get their whole first sentence; the third keeps its own');
  // Where the first sentences are one too, the description itself is given.
  assert.deepEqual(layoutLinesOf(described('Text beside a photo. The photo is on the left.', 'Text beside a photo. The photo is on the right.', 'A row of cards.')).slice(0, 2), [
    '- content: Text beside a photo. The photo is on the left.',
    '- closing: Text beside a photo. The photo is on the right.',
  ]);
});

test('only the layouts these tools can plan with are offered, and one that is named all the same is an error that says why', async () => {
  // No photo can be made and no deck field was given: of the three layouts, one is left. A free slide needs neither.
  const bare = presentationTools({ template, llm: scripted().llm, store: memoryStore() });
  for (const name of [TOOL_NAMES.plan, TOOL_NAMES.editPlan, TOOL_NAMES.edit]) {
    assert.deepEqual(toolNamed(bare, name).parameters.properties.layouts?.items?.enum, ['cards', 'free']);
  }
  assert.ok((toolNamed(bare, TOOL_NAMES.plan).parameters.properties.layouts?.description ?? '').endsWith(`The layouts:\n- cards: A row of cards.\n${FREE_LAYOUT}`));
  const refused = await bare.call(TOOL_NAMES.plan, { topic: 'Dogs', layouts: ['cards', 'content', 'closing', 'tables'] });
  assert.equal(
    refused.output.error,
    '"layouts" names layouts that cannot be used here: content needs a photo, and images are turned off; closing needs the deck field "closing"; tables is no layout of this template. The layouts to choose from: cards, free',
  );
  // The tools that change something say the same, before any model is asked and without touching the limit.
  const { tools, requests, store } = await toolsWith({ ...threeSlides, patterns: ['cards'] });
  await store.set(storeKeys.plan('plan_0'), encoder.encode(JSON.stringify({ ...oldPlan(), brief: { ...brief, patterns: ['cards'] } })));
  const one = '"layouts" names a layout that cannot be used here: content needs a photo, and images are turned off. The layouts to choose from: closing, cards, free';
  assert.equal((await tools.call(TOOL_NAMES.editPlan, { plan_id: 'plan_0', instruction: 'Photos too', layouts: ['cards', 'content'] })).output.error, one);
  assert.equal((await tools.call(TOOL_NAMES.edit, { presentation_id: 'deck_0', instruction: 'Photos too', layouts: ['content'] })).output.error, one);
  assert.equal((await tools.call(TOOL_NAMES.edit, { presentation_id: 'deck_0', accent_color: '1E6B3A', layouts: ['content'] })).output.error, one);
  assert.equal(requests.length, 0);
  assert.deepEqual(store.keys().sort(), ['deck/deck_0.json', 'plan/plan_0.json']);

  // Number layouts count as usable when the tools are declared: whether a brief holds figures is known only when it is planned.
  const root = join(dirname(fileURLToPath(import.meta.url)), '..');
  const demo = loadTemplate(readFileSync(join(root, 'templates', 'demo.pptx')), JSON.parse(readFileSync(join(root, 'templates', 'demo.manifest.json'), 'utf8')) as TemplateManifest);
  const model = scripted();
  const options = { template: demo, llm: model.llm, store: memoryStore(), layout: 'adaptive' as const, fields: { organizationDepartment: 'Office' } };
  const declared = toolNamed(presentationTools(options), TOOL_NAMES.plan).parameters.properties.layouts?.items?.enum ?? [];
  assert.ok(declared.includes('stat-cards') && declared.includes('cover') && declared.includes('closing'));
  for (const id of ['icon-cards', 'text-and-photo', 'wide-diagram']) assert.ok(!declared.includes(id), `${id} needs what was not supplied`);
  assert.ok(!(toolNamed(presentationTools({ ...options, numberPatterns: 'never' }), TOOL_NAMES.plan).parameters.properties.layouts?.items?.enum ?? []).includes('stat-cards'));
  // A number layout named for a brief without a figure cannot be planned, and the plan is not made from the rest of the list.
  const planned = await presentationTools(options).call(TOOL_NAMES.plan, { topic: 'Our key figures', layouts: ['cover', 'stat-cards'] });
  assert.equal(planned.output.error, 'The deck cannot be limited to a pattern that cannot be filled here: stat-cards is built around numbers, and the brief gives none');
  assert.equal(model.requests.length, 0);
});

test('a new brief that names a pattern which cannot be filled is an error that says which and why', async () => {
  const { llm, requests } = scripted();
  // `closing` shows a deck field, and this brief has none.
  await assert.rejects(planOutline(template, { prompt: 'A deck about dogs', language: 'en', patterns: ['cards', 'closing'] }, { llm }), (error: unknown) => {
    assert.ok(error instanceof PlanError);
    assert.equal(error.message, 'The deck cannot be limited to a pattern that cannot be filled here: closing needs the deck field "closing"');
    return true;
  });
  await assert.rejects(
    planOutline(template, { prompt: 'A deck about dogs', language: 'en', imageDensity: 'none', patterns: ['content', 'closing'] }, { llm }),
    /cannot be limited to patterns that cannot be filled here: content needs a photo, and images are turned off; closing needs the deck field "closing"$/,
  );
  assert.equal(requests.length, 0, 'no model is asked');
  // Without a list of its own, a brief is planned from what can be filled, as before.
  const open = scripted(cardsOutline);
  await planOutline(template, { prompt: 'A deck about dogs', language: 'en' }, { llm: open.llm });
  assert.deepEqual(offered(open.requests[0]), ['content', 'cards', 'free']);
});

test('an app that allows some layouts only, or no other colour, gets tools that say so', async () => {
  const some = presentationTools({ template, llm: scripted().llm, store: memoryStore(), fields, patterns: ['cards', 'closing', 'not-there'] });
  for (const name of [TOOL_NAMES.plan, TOOL_NAMES.editPlan, TOOL_NAMES.edit]) {
    assert.deepEqual(toolNamed(some, name).parameters.properties.layouts?.items?.enum, ['closing', 'cards']);
  }
  assert.doesNotMatch(toolNamed(some, TOOL_NAMES.plan).parameters.properties.layouts?.description ?? '', /content/);
  // A layout the app does not allow is not the person's to name.
  const { tools: narrow } = await toolsWith(threeSlides, [], { patterns: ['cards', 'closing'], ...everything });
  assert.equal(
    (await narrow.call(TOOL_NAMES.plan, { topic: 'Dogs', layouts: ['cards', 'content'] })).output.error,
    '"layouts" names a layout that cannot be used here: content is not among the layouts allowed here. The layouts to choose from: closing, cards',
  );
  // A list that names nothing of the template leaves no layout to choose: the argument is not declared, and one that comes is passed over.
  const nothing = await toolsWith(threeSlides, [], { patterns: ['not-there'] });
  assert.ok(!JSON.stringify(nothing.tools.definitions).includes('layouts'));
  assert.doesNotMatch(toolNamed(nothing.tools, TOOL_NAMES.plan).parameters.properties.text_only?.description ?? '', /layouts/);
  assert.equal((await nothing.tools.call(TOOL_NAMES.edit, { presentation_id: 'deck_0', accent_color: '1E6B3A', layouts: ['cards'] })).output.error, undefined);

  const fixed = loadTemplate(pptx, { ...cardsManifest, colors: false });
  for (const tools of [
    presentationTools({ template, llm: scripted().llm, store: memoryStore(), colors: false }),
    presentationTools({ template: fixed, llm: scripted().llm, store: memoryStore() }),
  ]) {
    assert.deepEqual(tools.definitions.map((tool) => tool.name), Object.values(TOOL_NAMES));
    assert.doesNotMatch(JSON.stringify(tools.definitions), /accent_color|color_slides|colour/);
    assert.deepEqual(toolNamed(tools, TOOL_NAMES.edit).parameters.required, ['presentation_id', 'instruction'], 'with no colour to give, an edit needs its instruction');
    assert.ok('layouts' in toolNamed(tools, TOOL_NAMES.plan).parameters.properties);
  }
  // Such tools pass a colour over, should one come all the same.
  const off = await toolsWith(threeSlides, [cardsOutline], { colors: false });
  const planned = await off.tools.call(TOOL_NAMES.plan, { topic: 'Dogs', accent_color: '1E6B3A' });
  assert.ok(!('colors' in (planned.plan?.brief ?? {})) && !('accent_color' in planned.output));
  assert.equal((await off.tools.call(TOOL_NAMES.edit, { presentation_id: 'deck_0', accent_color: '1E6B3A' })).output.error, '"instruction" is required');
});

test('a colour and no instruction recolours the presentation with no model call', async () => {
  const { tools, requests, saved, store } = await toolsWith(threeSlides);
  const green = await tools.call(TOOL_NAMES.edit, { presentation_id: 'deck_0', accent_color: '#1e6b3a' });
  assert.equal(green.output.error, undefined, String(green.output.error));
  assert.equal(requests.length, 0, 'no model is asked');
  assert.deepEqual(green.deck, { ...threeSlides, colors: { accent: '1E6B3A' } });
  assert.deepEqual(await saved('deck_1'), green.deck);
  assert.deepEqual(green.output, {
    presentation_id: 'deck_1',
    title: 'Dogs',
    file: 'pptx/deck_1.pptx',
    accent_color: '1E6B3A',
    slides: [
      { number: 1, layout: 'cards', title: 'Two' },
      { number: 2, layout: 'cards', title: 'More' },
      { number: 3, layout: 'closing' },
    ],
    summary: 'The presentation now has the colour 1E6B3A.',
    rewritten_slides: [],
    previous_presentation_id: 'deck_0',
  });
  assert.equal(readPresentation(PptxPackage.load(green.pptx as Uint8Array)).slides.length, 3, 'the file is rendered again');
  assert.deepEqual(await store.get(storeKeys.pptx('deck_1')), green.pptx);
  assert.deepEqual(await saved('deck_0'), threeSlides, 'the presentation it came from stays');

  // Three digits are doubled.
  const short = await tools.call(TOOL_NAMES.edit, { presentation_id: 'deck_1', accent_color: '#abc' });
  assert.deepEqual(short.deck?.colors, { accent: 'AABBCC' });
  assert.equal(requests.length, 0);
});

test('the colour reaches the file: the template\'s main colour gives way to the one asked for', async () => {
  const root = join(dirname(fileURLToPath(import.meta.url)), '..');
  const demo = loadTemplate(readFileSync(join(root, 'templates', 'demo.pptx')), JSON.parse(readFileSync(join(root, 'templates', 'demo.manifest.json'), 'utf8')) as TemplateManifest);
  const deck: DeckSpec = {
    version: 1,
    fields: { organizationDepartment: 'Riverbend' },
    slides: [
      { pattern: 'cover', slots: { title: 'A year of rides', subtitle: 'What the season taught us' } },
      { pattern: 'closing', slots: { title: 'Thank you' } },
    ],
  };
  const store = memoryStore();
  await store.set(storeKeys.deck('deck_0'), encoder.encode(JSON.stringify(deck)));
  const model = scripted();
  const tools = presentationTools({ template: demo, llm: model.llm, store, layout: 'adaptive', newId: () => 'deck_1' });
  const slidesOf = (file: Uint8Array) => {
    const pkg = PptxPackage.load(file);
    return readPresentation(pkg).slides.map((slide) => pkg.readText(slide.part)).join('');
  };

  const blue = await tools.call(TOOL_NAMES.edit, { presentation_id: 'deck_0', accent_color: '1565C0' });
  assert.equal(blue.output.error, undefined, String(blue.output.error));
  assert.equal(model.requests.length, 0);
  assert.ok(slidesOf(blue.pptx as Uint8Array).includes('1565C0'), 'the slides are drawn in the colour');
  // Back again: the file has the template's own colours, and none of the blue.
  await store.set(storeKeys.deck('deck_0'), encoder.encode(JSON.stringify(blue.deck)));
  const back = await tools.call(TOOL_NAMES.edit, { presentation_id: 'deck_0', accent_color: 'template' });
  assert.deepEqual(back.deck, deck);
  assert.ok(!slidesOf(back.pptx as Uint8Array).includes('1565C0'));
  assert.ok(slidesOf(back.pptx as Uint8Array).includes('0E7C7B'));
});

test('a colour close to the page is told to the chat model with the presentation; one that reads is not', async () => {
  const root = join(dirname(fileURLToPath(import.meta.url)), '..');
  const demo = loadTemplate(readFileSync(join(root, 'templates', 'demo.pptx')), JSON.parse(readFileSync(join(root, 'templates', 'demo.manifest.json'), 'utf8')) as TemplateManifest);
  // Only the file knows the page a colour lands on: a slide of text on the white page is needed for it.
  const deck: DeckSpec = {
    version: 1,
    fields: { organizationDepartment: 'Riverbend' },
    slides: [
      { pattern: 'cover', slots: { title: 'A year of rides', subtitle: 'What the season taught us' } },
      { pattern: 'text-only', slots: { title: 'What we saw', body: 'Rides grew in every month.' } },
      { pattern: 'closing', slots: { title: 'Thank you' } },
    ],
  };
  const store = memoryStore();
  await store.set(storeKeys.deck('deck_0'), encoder.encode(JSON.stringify(deck)));
  let next = 0;
  const model = scripted();
  const tools = presentationTools({ template: demo, llm: model.llm, store, layout: 'adaptive', newId: () => `deck_${++next}` });
  const white = await tools.call(TOOL_NAMES.edit, { presentation_id: 'deck_0', accent_color: 'FFFFFF' });
  assert.equal(white.output.error, undefined, String(white.output.error));
  assert.deepEqual(white.output.warnings, ["the colour FFFFFF is close to the page's own; fills, chart bars and icons in it will hardly show"]);
  const blue = await tools.call(TOOL_NAMES.edit, { presentation_id: 'deck_0', accent_color: '1565C0' });
  assert.equal(blue.output.warnings, undefined);
  assert.equal(model.requests.length, 0);
});

test('color_slides sets the colour on those slides and leaves the deck\'s own alone', async () => {
  const { tools, requests } = await toolsWith({ ...threeSlides, colors: { accent: '1E6B3A' } });
  const red = await tools.call(TOOL_NAMES.edit, { presentation_id: 'deck_0', accent_color: 'C62828', color_slides: [2] });
  assert.equal(requests.length, 0);
  assert.deepEqual(red.deck?.colors, { accent: '1E6B3A' });
  assert.deepEqual(red.deck?.slides.map((slide) => slide.colors), [undefined, { accent: 'C62828' }, undefined]);
  assert.equal(red.output.accent_color, '1E6B3A');
  assert.deepEqual(red.output.slides, [
    { number: 1, layout: 'cards', title: 'Two' },
    { number: 2, layout: 'cards', title: 'More', accent_color: 'C62828' },
    { number: 3, layout: 'closing' },
  ]);
  assert.equal(red.output.summary, 'Slide 2 now has the colour C62828.');

  // On a presentation with no colour of its own, the slides alone get one.
  const plain = await toolsWith(threeSlides);
  const two = await plain.tools.call(TOOL_NAMES.edit, { presentation_id: 'deck_0', accent_color: 'C62828', color_slides: [3, 1] });
  assert.ok(!('colors' in (two.deck as DeckSpec)) && !('accent_color' in two.output));
  assert.deepEqual(two.deck?.slides.map((slide) => slide.colors), [{ accent: 'C62828' }, undefined, { accent: 'C62828' }]);
  assert.equal(two.output.summary, 'Slides 1 and 3 now have the colour C62828.');
  // An empty list names no slide: the whole presentation is meant.
  const all = await plain.tools.call(TOOL_NAMES.edit, { presentation_id: 'deck_0', accent_color: 'C62828', color_slides: [] });
  assert.deepEqual(all.deck?.colors, { accent: 'C62828' });

  // A slide number is a whole number from 1 to the number of slides, or digits that say one. Nothing else is read as one.
  const wrong = '"color_slides" must be slide numbers from 1 to 3';
  for (const slides of [[2, 9], [0], [-1], [4], [1.5], [true], true, false, [null], [[2]], ['0x2'], ['1e0'], ['two'], 'all', '1,2', {}]) {
    assert.equal((await plain.tools.call(TOOL_NAMES.edit, { presentation_id: 'deck_0', accent_color: 'C62828', color_slides: slides })).output.error, wrong, JSON.stringify(slides));
  }
  const digits = await plain.tools.call(TOOL_NAMES.edit, { presentation_id: 'deck_0', accent_color: 'C62828', color_slides: ['2', ' 3 ', 2] });
  assert.deepEqual(digits.deck?.slides.map((slide) => slide.colors), [undefined, { accent: 'C62828' }, { accent: 'C62828' }]);
  assert.equal(plain.requests.length, 0);
});

test('color_slides without a colour is an error before any model is asked', async () => {
  const { tools, requests, store } = await toolsWith(threeSlides, [editAnswer([[1, 'cards'], [2, 'cards'], [3, 'closing']])]);
  // The instruction may well ask for a colour: without accent_color nothing would be coloured, and the answer could not show it.
  const needs = '"color_slides" needs "accent_color"';
  assert.equal((await tools.call(TOOL_NAMES.edit, { presentation_id: 'deck_0', instruction: 'Make slide 2 red', color_slides: [2] })).output.error, needs);
  assert.equal((await tools.call(TOOL_NAMES.edit, { presentation_id: 'deck_0', instruction: 'Make slide 2 red', color_slides: '2', accent_color: '' })).output.error, needs);
  assert.equal((await tools.call(TOOL_NAMES.edit, { presentation_id: 'deck_0', instruction: 'Make slide 9 red', color_slides: [9] })).output.error, '"color_slides" must be slide numbers from 1 to 3');
  assert.equal(requests.length, 0);
  assert.deepEqual(store.keys(), ['deck/deck_0.json']);
  // A list that names no slide asks for nothing, and the edit goes on.
  const plain = await tools.call(TOOL_NAMES.edit, { presentation_id: 'deck_0', instruction: 'Nothing to change', color_slides: [] });
  assert.equal(plain.output.error, undefined, String(plain.output.error));
  assert.equal(requests.length, 1);
});

test('"template" returns to the template\'s colours', async () => {
  const colored: DeckSpec = { ...threeSlides, colors: { accent: '1E6B3A' }, slides: [twoCards(), { ...twoCards('More'), colors: { accent: 'C62828' } }, { pattern: 'closing', slots: {} }] };
  const { tools, requests } = await toolsWith(colored);
  // One slide of a presentation that has a colour: the slide says that it keeps the template's own, or it would follow the presentation.
  const one = await tools.call(TOOL_NAMES.edit, { presentation_id: 'deck_0', accent_color: 'template', color_slides: [2, 3] });
  assert.deepEqual(one.deck, { ...colored, slides: [twoCards(), { ...twoCards('More'), colors: {} }, { pattern: 'closing', slots: {}, colors: {} }] });
  assert.equal(one.output.summary, "Slides 2 and 3 are back in the template's colours.");
  assert.equal(one.output.accent_color, '1E6B3A', 'the presentation keeps its colour');
  // The chat model is told which slides do not have the presentation's colour, in the word the argument takes.
  assert.deepEqual(one.output.slides, [
    { number: 1, layout: 'cards', title: 'Two' },
    { number: 2, layout: 'cards', title: 'More', accent_color: 'template' },
    { number: 3, layout: 'closing', accent_color: 'template' },
  ]);
  assert.deepEqual(validateDeck(template.manifest, one.deck as DeckSpec, { layout: 'adaptive' }), [], 'an empty set of colours is a valid thing for a slide to state');
  // In a presentation without a colour of its own, the slide's colours simply go.
  const { colors: _deckColors, ...uncolored } = colored;
  const plain = await toolsWith(uncolored);
  const back = await plain.tools.call(TOOL_NAMES.edit, { presentation_id: 'deck_0', accent_color: 'template', color_slides: [2] });
  assert.deepEqual(back.deck, threeSlides);
  assert.equal(back.output.summary, "Slide 2 is back in the template's colours.");
  // With an instruction too, the slide is found where the edit put it.
  const moved = await toolsWith(colored, [editAnswer([[2, 'cards'], [1, 'cards'], [3, 'closing']])]);
  const first = await moved.tools.call(TOOL_NAMES.edit, { presentation_id: 'deck_0', instruction: 'Put slide 2 first', accent_color: 'template', color_slides: [2] });
  assert.deepEqual(first.deck?.slides.map((slide) => slide.colors), [{}, undefined, undefined]);
  assert.equal(first.output.summary, "Done. Slide 1 is back in the template's colours.");
  // The presentation: every colour that was asked for goes.
  const all = await tools.call(TOOL_NAMES.edit, { presentation_id: 'deck_0', accent_color: 'Template' });
  assert.deepEqual(all.deck, threeSlides);
  assert.ok(!('accent_color' in all.output));
  assert.equal(all.output.summary, "The presentation is back in the template's colours.");
  // A colour for the whole presentation is likewise for every slide: the one that had its own takes it too.
  const blue = await tools.call(TOOL_NAMES.edit, { presentation_id: 'deck_0', accent_color: '1565C0' });
  assert.deepEqual(blue.deck, { ...threeSlides, colors: { accent: '1565C0' } });
  assert.equal(requests.length, 0);
});

test('a colour that is no colour, or nothing to do, comes back as an error the chat model can act on', async () => {
  const { tools, requests, store } = await toolsWith(threeSlides);
  await store.set(storeKeys.plan('plan_0'), encoder.encode(JSON.stringify(oldPlan())));
  const wrong = '"accent_color" must be six hex digits such as 1E6B3A';
  for (const color of ['dark green', '#1E6B3', '1E6B3A00', 'GGGGGG', 12]) {
    assert.equal((await tools.call(TOOL_NAMES.edit, { presentation_id: 'deck_0', accent_color: color })).output.error, wrong, String(color));
  }
  assert.equal((await tools.call(TOOL_NAMES.edit, { presentation_id: 'deck_0', instruction: 'Shorten slide 1', accent_color: 'green' })).output.error, wrong);
  assert.equal((await tools.call(TOOL_NAMES.plan, { topic: 'Dogs', accent_color: 'green' })).output.error, wrong);
  assert.equal((await tools.call(TOOL_NAMES.editPlan, { plan_id: 'plan_0', instruction: 'Greener', accent_color: 'green' })).output.error, wrong);

  assert.equal((await tools.call(TOOL_NAMES.edit, { presentation_id: 'deck_0' })).output.error, '"instruction" is required unless a colour is given');
  assert.equal((await tools.call(TOOL_NAMES.edit, { presentation_id: 'deck_0', instruction: '  ', layouts: ['cards'] })).output.error, '"instruction" is required unless a colour is given');
  // A layout the template does not have is said by the tool, in its own words, with the layouts there are.
  for (const call of [
    tools.call(TOOL_NAMES.plan, { topic: 'Dogs', layouts: ['tables'] }),
    tools.call(TOOL_NAMES.editPlan, { plan_id: 'plan_0', instruction: 'Only tables', layouts: ['tables'] }),
    tools.call(TOOL_NAMES.edit, { presentation_id: 'deck_0', instruction: 'Only tables', layouts: ['tables'] }),
  ]) {
    assert.equal((await call).output.error, '"layouts" names a layout that cannot be used here: tables is no layout of this template. The layouts to choose from: closing, cards, free');
  }
  assert.equal(requests.length, 0, 'none of these reached a model');
  assert.deepEqual(store.keys().sort(), ['deck/deck_0.json', 'plan/plan_0.json'], 'and nothing was saved');
});

test('a layouts argument that names nothing sets nothing: only an empty list lifts a limit', async () => {
  const revised = { summary: 'Done.', language: 'en', title: 'Dogs', slides: oldPlan().slides };
  const keep = editAnswer([[1, 'cards'], [2, 'cards'], [3, 'closing']]);
  const odd = [[7], [null], [' '], [{}], 7, {}, null, ''];
  const { tools, store } = await toolsWith({ ...threeSlides, patterns: ['cards', 'closing'] }, [...odd.map(() => revised), ...odd.map(() => keep), revised, keep]);
  await store.set(storeKeys.plan('plan_0'), encoder.encode(JSON.stringify({ ...oldPlan(), brief: { ...brief, patterns: ['cards', 'closing'] } })));

  for (const layouts of odd) {
    const plan = await tools.call(TOOL_NAMES.editPlan, { plan_id: 'plan_0', instruction: 'A better title', layouts });
    assert.deepEqual(plan.plan?.brief.patterns, ['cards', 'closing'], JSON.stringify(layouts));
    assert.deepEqual(plan.output.layouts, ['cards', 'closing']);
  }
  for (const layouts of odd) {
    const deck = await tools.call(TOOL_NAMES.edit, { presentation_id: 'deck_0', instruction: 'Nothing to change', layouts });
    assert.deepEqual(deck.deck?.patterns, ['cards', 'closing'], JSON.stringify(layouts));
    // Nor without a model: a colour alone, with such an argument beside it.
    const recolored = await tools.call(TOOL_NAMES.edit, { presentation_id: 'deck_0', accent_color: '1E6B3A', layouts });
    assert.deepEqual(recolored.deck?.patterns, ['cards', 'closing']);
    assert.equal(recolored.output.summary, 'The presentation now has the colour 1E6B3A.');
  }
  // The empty list is the one way to lift it. A single id may come as a string.
  const lifted = await tools.call(TOOL_NAMES.editPlan, { plan_id: 'plan_0', instruction: 'Any layout', layouts: [] });
  assert.ok(!('patterns' in (lifted.plan?.brief ?? {})) && !('layouts' in lifted.output));
  const single = await tools.call(TOOL_NAMES.edit, { presentation_id: 'deck_0', instruction: 'Only cards', layouts: 'cards' });
  assert.deepEqual(single.deck?.patterns, ['cards']);
  const free = await tools.call(TOOL_NAMES.edit, { presentation_id: 'deck_0', accent_color: '1E6B3A', layouts: [] });
  assert.ok(!('patterns' in (free.deck as DeckSpec)));
  assert.equal(free.output.summary, 'The presentation now has the colour 1E6B3A. The limit on the layouts is lifted.');
});

test('the layouts said back are the limit in force, not whatever the stored list holds', async () => {
  const revised = { summary: 'Done.', language: 'en', title: 'Dogs', slides: oldPlan().slides };
  // A list that names only what the template has lost limits nothing, and the chat model is not told of a limit.
  const lost = await toolsWith({ ...threeSlides, patterns: ['gone'] }, [revised]);
  await lost.store.set(storeKeys.plan('plan_0'), encoder.encode(JSON.stringify({ ...oldPlan(), brief: { ...brief, patterns: ['gone'] } })));
  const plan = await lost.tools.call(TOOL_NAMES.editPlan, { plan_id: 'plan_0', instruction: 'A better title' });
  assert.doesNotMatch(promptOf(lost.requests[0]), /was limited to these layouts/);
  assert.ok(!('layouts' in plan.output));
  const deck = await lost.tools.call(TOOL_NAMES.edit, { presentation_id: 'deck_0', accent_color: '1E6B3A' });
  assert.deepEqual(deck.deck?.patterns, ['gone'], 'the deck itself is not rewritten');
  assert.ok(!('layouts' in deck.output));

  // The app allows cards and closing; the person had named the photo layout and cards. What holds is what both allow.
  const app = await toolsWith({ ...threeSlides, patterns: ['content', 'cards'] }, [revised], { patterns: ['cards', 'closing'], ...everything });
  await app.store.set(storeKeys.plan('plan_0'), encoder.encode(JSON.stringify({ ...oldPlan(), brief: { ...brief, patterns: ['content', 'cards'] } })));
  const narrowed = await app.tools.call(TOOL_NAMES.editPlan, { plan_id: 'plan_0', instruction: 'A better title' });
  assert.ok(rulesOf(app.requests[0]).includes(limited('cards')));
  assert.deepEqual(narrowed.output.layouts, ['cards']);
  assert.deepEqual(narrowed.plan?.brief.patterns, ['content', 'cards'], 'the person\'s wish is kept as it was');
  assert.deepEqual((await app.tools.call(TOOL_NAMES.edit, { presentation_id: 'deck_0', accent_color: '1E6B3A' })).output.layouts, ['cards']);
});

test('an instruction with a colour is planned by the model and coloured by code', async () => {
  const { tools, requests } = await toolsWith(threeSlides, [editAnswer([[3, 'closing'], [1, 'cards'], [2, 'cards']])]);
  const edited = await tools.call(TOOL_NAMES.edit, { presentation_id: 'deck_0', instruction: 'Put the closing slide first', accent_color: '1565C0', color_slides: [1], layouts: ['cards', 'closing'] });
  assert.equal(edited.output.error, undefined, String(edited.output.error));
  assert.equal(requests.length, 1, 'the plan of the edit; nothing is rewritten');
  assert.doesNotMatch(JSON.stringify(requests[0]), /1565C0/, 'the model is not told the colour');
  assert.deepEqual(edited.deck?.slides.map((slide) => [slide.pattern, slide.colors]), [['closing', undefined], ['cards', { accent: '1565C0' }], ['cards', undefined]]);
  assert.deepEqual(edited.deck?.patterns, ['cards', 'closing']);
  assert.deepEqual(edited.output.layouts, ['cards', 'closing']);
  // The model's sentence, then what code did: the model was told that colours are not its work. The slide is named by its new number.
  assert.equal(edited.output.summary, 'Done. Slide 2 now has the colour 1565C0.');
  assert.match(promptOf(requests[0]), /\n- This deck was limited to these layouts: cards, closing\./);

  // With no slides named, the whole presentation takes the colour, the slide that had its own too.
  const again = await toolsWith(edited.deck as DeckSpec, [editAnswer([[1, 'closing'], [2, 'cards'], [3, 'cards']])]);
  const whole = await again.tools.call(TOOL_NAMES.edit, { presentation_id: 'deck_0', instruction: 'Nothing else', accent_color: 'C62828' });
  assert.deepEqual(whole.deck?.colors, { accent: 'C62828' });
  assert.deepEqual(whole.deck?.slides.map((slide) => slide.colors), [undefined, undefined, undefined]);
  assert.deepEqual(whole.deck?.patterns, ['cards', 'closing'], 'the list the deck had stays');
  assert.equal(whole.output.summary, 'Done. The presentation now has the colour C62828.');

  // A slide the same edit removes cannot be coloured. The summary claims no colour, and a warning says what happened.
  const drop = await toolsWith(threeSlides, [{ ...editAnswer([[1, 'cards'], [3, 'closing']]), summary: 'Removed slide 2.' }]);
  const dropped = await drop.tools.call(TOOL_NAMES.edit, { presentation_id: 'deck_0', instruction: 'Remove slide 2', accent_color: '1565C0', color_slides: [2] });
  assert.equal(dropped.output.summary, 'Removed slide 2.');
  assert.deepEqual(dropped.output.warnings, ['the change removed slide 2, which was to be recoloured']);
  assert.ok(!('colors' in (dropped.deck as DeckSpec)) && dropped.deck?.slides.every((slide) => slide.colors === undefined));

  // Without a colour the summary is the model's alone: nothing was set that code could speak of.
  const plain = await toolsWith(threeSlides, [{ ...editAnswer([[1, 'cards'], [2, 'cards'], [3, 'closing']]), summary: 'The colour was not changed.' }]);
  assert.equal((await plain.tools.call(TOOL_NAMES.edit, { presentation_id: 'deck_0', instruction: 'Make it dark green' })).output.summary, 'The colour was not changed.');
});

test('a plan carries the layouts and the colour the person asked for, says so, and hands them to the presentation', async () => {
  const outline = { title: 'Dogs', slides: [{ pattern: 'cards', purpose: 'Two things', items: 2 }, { pattern: 'closing', purpose: 'Thanks' }] };
  const revised = { summary: 'Done.', language: 'en', title: 'Dogs', slides: outline.slides };
  const { tools, requests, store } = await toolsWith(threeSlides, [outline, revised, { s1: cardsContent, s2: {} }, revised]);

  const planned = await tools.call(TOOL_NAMES.plan, { topic: 'A deck about dogs', language: 'en', layouts: ['cards', 'closing'], accent_color: '1e6b3a' });
  assert.deepEqual(offered(requests[0]), ['closing', 'cards']);
  assert.deepEqual(planned.plan?.brief.patterns, ['cards', 'closing']);
  assert.deepEqual(planned.plan?.brief.colors, { accent: '1E6B3A' });
  assert.equal(planned.output.accent_color, '1E6B3A');
  assert.deepEqual(planned.output.layouts, ['cards', 'closing']);
  assert.deepEqual(Object.keys(planned.output), ['plan_id', 'title', 'language', 'accent_color', 'layouts', 'slides', 'next']);

  // A change to the plan that sets neither keeps both. An argument that names no layouts sets nothing.
  const kept = await tools.call(TOOL_NAMES.editPlan, { plan_id: 'plan_1', instruction: 'A better title', layouts: 7, accent_color: '' });
  assert.equal(kept.output.accent_color, '1E6B3A');
  assert.deepEqual(kept.output.layouts, ['cards', 'closing']);
  assert.match(promptOf(requests[1]), /\n- This deck was limited to these layouts: cards, closing\./);

  const created = await tools.call(TOOL_NAMES.create, { plan_id: 'plan_2' });
  assert.equal(created.output.error, undefined, String(created.output.error));
  assert.deepEqual(created.deck?.colors, { accent: '1E6B3A' });
  assert.deepEqual(created.deck?.patterns, ['cards', 'closing']);
  assert.equal(created.output.accent_color, '1E6B3A');
  assert.deepEqual(created.output.layouts, ['cards', 'closing']);
  assert.equal(created.output.warnings, undefined);

  // An empty list lifts the limit, and "template" takes the colour away.
  const lifted = await tools.call(TOOL_NAMES.editPlan, { plan_id: 'plan_2', instruction: 'Any layout, usual colours', layouts: [], accent_color: 'template' });
  assert.ok(!('patterns' in (lifted.plan?.brief ?? {})) && !('colors' in (lifted.plan?.brief ?? {})));
  assert.deepEqual(Object.keys(lifted.output), ['plan_id', 'title', 'language', 'slides', 'next', 'summary', 'previous_plan_id']);
  assert.deepEqual(offered(requests[3]), ['closing', 'cards', 'free'], 'no photos can be made here, so the photo layout is not on offer; with the limit lifted, free slides are');
  assert.ok((await store.get(storeKeys.plan('plan_4'))) !== undefined);
  // The model's sentence, then what code set.
  assert.equal(lifted.output.summary, "Done. The presentation will be in the template's colours.");
});

test('a colour for a plan needs no model: an empty instruction with a colour is set by code', async () => {
  const revised = { summary: 'Done.', language: 'en', title: 'Dogs', slides: oldPlan().slides };
  const { tools, requests, store } = await toolsWith(threeSlides, [revised]);
  await store.set(storeKeys.plan('plan_0'), encoder.encode(JSON.stringify(oldPlan())));

  for (const instruction of ['', '   ', undefined]) {
    const green = await tools.call(TOOL_NAMES.editPlan, { plan_id: 'plan_0', ...(instruction === undefined ? {} : { instruction }), accent_color: '#1e6b3a', layouts: ['cards', 'closing'] });
    assert.equal(green.output.error, undefined, String(green.output.error));
    // The plan is the one that was stored, with the colour and the list: no word about colour joins its brief.
    assert.deepEqual(green.plan, { ...oldPlan(), brief: { ...brief, colors: { accent: '1E6B3A' }, patterns: ['cards', 'closing'] } });
    assert.equal(green.output.accent_color, '1E6B3A');
    assert.deepEqual(green.output.layouts, ['cards', 'closing']);
    assert.equal(green.output.summary, 'The presentation will have the colour 1E6B3A. The presentation is now limited to the layouts cards, closing. Slides that are there already keep their layout.');
    assert.equal(green.output.previous_plan_id, 'plan_0');
    assert.deepEqual(JSON.parse(decoder.decode(await store.get(storeKeys.plan(green.output.plan_id as string)))), green.plan);
  }
  const back = await tools.call(TOOL_NAMES.editPlan, { plan_id: 'plan_1', instruction: '', accent_color: 'template' });
  assert.ok(!('colors' in (back.plan?.brief ?? {})));
  assert.deepEqual(back.plan?.brief.patterns, ['cards', 'closing'], 'the list is left as it is');
  assert.equal(back.output.summary, "The presentation will be in the template's colours.");
  assert.equal(requests.length, 0, 'no model is asked');

  // With nothing to do, and with layouts alone, the instruction is asked for, in the words the tool that changes a presentation uses.
  assert.equal((await tools.call(TOOL_NAMES.editPlan, { plan_id: 'plan_0' })).output.error, '"instruction" is required unless a colour is given');
  assert.equal((await tools.call(TOOL_NAMES.editPlan, { plan_id: 'plan_0', instruction: ' ', layouts: ['cards'] })).output.error, '"instruction" is required unless a colour is given');
  // With an instruction the model plans the change, and the colour is still code's: its sentence follows the model's.
  const both = await tools.call(TOOL_NAMES.editPlan, { plan_id: 'plan_0', instruction: 'A better title', accent_color: '1E6B3A' });
  assert.equal(requests.length, 1);
  assert.equal(both.plan?.brief.instructions, 'A better title');
  assert.equal(both.output.summary, 'Done. The presentation will have the colour 1E6B3A.');
});

test('text_only gives way to layouts in the tools too, and the plan is not called text only', async () => {
  const { tools, requests } = await toolsWith(threeSlides, [cardsOutline, { title: 'Dogs', slides: [{ pattern: 'closing', purpose: 'Thanks' }] }]);
  const listed = await tools.call(TOOL_NAMES.plan, { topic: 'Dogs', language: 'en', text_only: true, layouts: ['cards', 'closing'] });
  assert.deepEqual(offered(requests[0]), ['closing', 'cards']);
  assert.ok(!('textOnly' in (listed.plan?.brief ?? {})), 'cards are no plain text: the plan does not promise it');
  assert.deepEqual(listed.plan?.brief.patterns, ['cards', 'closing']);
  // Alone, or with an argument that names no layouts, it is what it was.
  const plain = await tools.call(TOOL_NAMES.plan, { topic: 'Dogs', language: 'en', text_only: true, layouts: [] });
  assert.equal(plain.plan?.brief.textOnly, true);
  assert.deepEqual(offered(requests[1]), ['closing', 'free'], 'the plain text layout, and a free slide of text');
});

test('the app\'s own list reaches every tool', async () => {
  const outline = { title: 'Dogs', slides: [{ pattern: 'cards', purpose: 'Two things', items: 2 }] };
  const revised = { summary: 'Done.', language: 'en', title: 'Dogs', slides: outline.slides };
  const images = { generateImage: async () => new Uint8Array([1]) };
  const { tools, requests } = await toolsWith(threeSlides, [outline, revised, { s1: cardsContent }, editAnswer([[1, 'cards']])], { patterns: ['cards'], images });

  await tools.call(TOOL_NAMES.plan, { topic: 'A deck about dogs', language: 'en' });
  assert.deepEqual(offered(requests[0]), ['cards']);
  await tools.call(TOOL_NAMES.editPlan, { plan_id: 'plan_1', instruction: 'A better title' });
  assert.deepEqual(offered(requests[1]), ['cards']);
  const created = await tools.call(TOOL_NAMES.create, { plan_id: 'plan_2' });
  assert.equal(created.output.error, undefined, String(created.output.error));
  assert.ok(!('patterns' in (created.deck as DeckSpec)) && !('layouts' in created.output), 'the app\'s list is not the person\'s: the deck does not carry it');
  await tools.call(TOOL_NAMES.edit, { presentation_id: 'deck_3', instruction: 'Nothing to change' });
  assert.deepEqual(entryOf(requests[3]).properties.pattern?.enum, ['cards']);

  // A plan that was stored with a layout the app no longer allows cannot be written, and says which.
  const narrow = await toolsWith(threeSlides, [], { patterns: ['closing'] });
  await narrow.store.set(storeKeys.plan('plan_0'), encoder.encode(JSON.stringify(oldPlan())));
  assert.match(String((await narrow.tools.call(TOOL_NAMES.create, { plan_id: 'plan_0' })).output.error), /cannot be filled with what was supplied: cards/);
});

test('a warning about the deck as a whole is given as its message, with no slide named', async () => {
  // A stored plan whose colour is no colour: the deck is made, and the check of the deck says what was left out.
  const plan = oldPlan();
  plan.brief = { ...plan.brief, colors: { accent: 'blue' } };
  const { tools, store } = await toolsWith(threeSlides, [{ s1: cardsContent, s2: {} }]);
  await store.set(storeKeys.plan('plan_0'), encoder.encode(JSON.stringify(plan)));
  const created = await tools.call(TOOL_NAMES.create, { plan_id: 'plan_0' });
  assert.equal(created.output.error, undefined, String(created.output.error));
  const expected = validateDeck(template.manifest, created.deck as DeckSpec).filter((issue) => issue.slide === undefined).map((issue) => issue.message);
  assert.equal(expected.length, 1);
  assert.deepEqual(created.output.warnings, expected);
  assert.doesNotMatch(JSON.stringify(created.output.warnings), /undefined/);
  assert.ok(!('accent_color' in created.output), 'a colour that is none is not said back as one');
});
