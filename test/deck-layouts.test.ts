// Layouts made for one deck: the designs travel in the plan and the deck, and are built onto the
// template whenever the deck is rendered, checked or edited. The template is never changed. No
// test calls a live model.
import assert from 'node:assert/strict';
import { readFileSync } from 'node:fs';
import { dirname, join } from 'node:path';
import { test } from 'node:test';
import { fileURLToPath } from 'node:url';
import {
  PptxPackage,
  buildPattern,
  composeLayouts,
  designKit,
  editDeck,
  loadTemplate,
  planDeck,
  planOutline,
  renderDeck,
  validateDeck,
  withLayouts,
  writeDeck,
} from '../src/index.ts';
import type { DeckPlan, DeckSpec, DesignElement, JsonRequest, LlmProvider, PatternDesign, Template, TemplateManifest } from '../src/index.ts';
import { readPresentation } from '../src/pptx/presentation.ts';

const root = join(dirname(fileURLToPath(import.meta.url)), '..');
const pptx = readFileSync(join(root, 'templates', 'demo.pptx'));
const manifest = JSON.parse(readFileSync(join(root, 'templates', 'demo.manifest.json'), 'utf8')) as TemplateManifest;
const demo = (): Template => loadTemplate(pptx, manifest);
const fields = { organizationDepartment: 'Strategy office' };

function styleFor(template: Template, use: string): string {
  const style = designKit(template)?.styles.find((candidate) => candidate.text.includes(`used for ${use}`));
  assert.ok(style, `the kit has a style used for ${use}`);
  return style.key;
}
const panel = (template: Template) => (designKit(template)?.styles.find((style) => style.kind === 'decoration' && style.shape === 'area' && style.text.includes('panel')) as { key: string }).key;

/** Three goals on cards under the title of the text-only layout, which sits a little lower than most. */
function threeGoals(template: Template, id = 'three-goals'): PatternDesign {
  const heading = styleFor(template, 'card1Heading');
  const body = styleFor(template, 'card1Text');
  const card = panel(template);
  const elements: DesignElement[] = [];
  const part = (name: string, style: string, role: DesignElement['role'], column: number, row: number, height: number, sample: string): DesignElement =>
    ({ name, style, role, optional: false, sample, column, row, width: 4, height });
  for (const [n, column] of [[1, 1], [2, 5], [3, 9]] as const) {
    elements.push(part('', card, 'decoration', column, 6, 7, ''));
    elements.push(part(`goal${n}Heading`, heading, 'subtitle', column, 6, 2, `Goal number ${n}`));
    elements.push(part(`goal${n}Text`, body, 'text', column, 8, 5, 'A sentence or two on what reaching this goal means.'));
  }
  return { id, description: 'Three goals side by side on cards, each with a heading and a sentence.', keep: ['title'], base: 'text-only', elements };
}

const goalSlots = { title: 'Three goals for the year', goal1Heading: 'Grow', goal1Text: 'Win two new regions.', goal2Heading: 'Keep', goal2Text: 'Keep every client we have.', goal3Heading: 'Learn', goal3Text: 'Train every team twice.' };

function deckWith(design: PatternDesign): DeckSpec {
  return {
    version: 1,
    fields,
    layouts: [design],
    slides: [
      { pattern: 'cover', slots: { title: 'The year ahead' } },
      { pattern: design.id, slots: goalSlots },
      { pattern: 'closing', slots: { title: 'Thank you' } },
    ],
  };
}

const slideXml = (bytes: Uint8Array, number: number): string => {
  const pkg = PptxPackage.load(bytes);
  return pkg.readText((readPresentation(pkg).slides[number - 1] as { part: string }).part);
};

test('a deck that carries its own layout renders, and renders the same again, with the template unchanged', () => {
  const template = demo();
  const design = threeGoals(template);
  const deck = deckWith(design);
  const first = renderDeck(template, deck, { layout: 'adaptive' });
  assert.equal(readPresentation(PptxPackage.load(first.pptx)).slides.length, 3);
  const xml = slideXml(first.pptx, 2);
  for (const value of Object.values(goalSlots)) assert.ok(xml.includes(value), `slide 2 shows "${value}"`);
  assert.ok(!xml.includes('Goal number'), 'no sample text is left');

  assert.ok(!template.manifest.patterns.some((pattern) => pattern.id === design.id), 'the template has not gained the layout');
  assert.equal(readPresentation(template.pkg).slides.length, 13, 'nor a slide');
  // Built once for the template and the list, so a chat that renders a deck many times builds it once.
  assert.equal(withLayouts(template, deck.layouts), withLayouts(template, structuredClone(deck.layouts)));
  // Another load of the template, as after a restart, builds the same slide.
  const again = renderDeck(demo(), structuredClone(deck), { layout: 'adaptive' });
  assert.equal(slideXml(again.pptx, 2), xml);
});

test('the layout stands on the frame of the layout it was asked for, and says so in its design', () => {
  const template = demo();
  assert.equal(designKit(template)?.base, 'list-with-photo', 'the usual base');
  const kit = designKit(template, undefined, 'text-only');
  assert.equal(kit?.base, 'text-only');
  assert.deepEqual(kit?.keepable.find((entry) => entry.slot === 'title')?.cells.row, 2, 'the title is where text-only moves it');
  assert.equal(designKit(template, undefined, 'icon-cards')?.base, 'list-with-photo', 'a row of items is no base: the usual one is used');
  assert.equal(designKit(template, undefined, 'no-such-layout')?.base, 'list-with-photo');

  const built = buildPattern(template, threeGoals(template));
  const pattern = built.manifest.patterns.find((entry) => entry.id === 'three-goals');
  assert.ok(built.summary.includes('on the frame of "text-only"'), built.summary);
  const textOnly = manifest.patterns.find((entry) => entry.id === 'text-only');
  assert.deepEqual(pattern?.slots.find((slot) => slot.name === 'title'), textOnly?.slots.find((slot) => slot.name === 'title'), 'the title is the one text-only shows, box and all');
});

test('validating a deck against the bare manifest says to build its layouts first', () => {
  const template = demo();
  const deck = deckWith(threeGoals(template));
  const issues = validateDeck(manifest, deck, { layout: 'adaptive' });
  assert.match(issues.map((issue) => issue.message).join('\n'), /"three-goals" is one of the deck's own layouts: check the deck against withLayouts/);
  assert.deepEqual(validateDeck(withLayouts(template, deck.layouts).manifest, deck, { layout: 'adaptive' }).filter((issue) => issue.severity === 'error'), []);
});

test('once the template has a layout of that id, the template\'s own is used', () => {
  const template = demo();
  const design = threeGoals(template);
  const built = buildPattern(template, design);
  const promoted = loadTemplate(built.pptx, built.manifest);
  assert.equal(withLayouts(promoted, [design]), promoted, 'nothing is built again');
  const rendered = renderDeck(promoted, deckWith(design), { layout: 'adaptive' });
  assert.equal(readPresentation(PptxPackage.load(rendered.pptx)).slides.length, 3);
});

/** A model that answers each kind of request the planner makes, and records them. */
function planner(answers: { outline: unknown; decision: unknown; designs?: unknown[]; content?: unknown }): { llm: LlmProvider; requests: JsonRequest[] } {
  const requests: JsonRequest[] = [];
  const designs = [...(answers.designs ?? [])];
  return {
    requests,
    llm: {
      async generateJson(request) {
        requests.push(request);
        if (request.system.startsWith('You review the plan')) return answers.decision;
        if (request.system.startsWith('You design one new slide layout')) {
          if (designs.length === 0) throw new Error('no design left');
          return designs.shift();
        }
        const properties = (request.schema as { properties: Record<string, unknown> }).properties;
        if ('slides' in properties && 'title' in properties) return answers.outline;
        return answers.content;
      },
    },
  };
}

const textOnlyOutline = {
  title: 'The year ahead',
  slides: [
    { pattern: 'cover', purpose: 'Open' },
    { pattern: 'text-only', purpose: 'Where we stand' },
    { pattern: 'text-only', purpose: 'Our three goals: grow, keep, learn' },
    { pattern: 'text-only', purpose: 'What it costs' },
    { pattern: 'closing', purpose: 'Thanks' },
  ],
};
const brief = { prompt: 'The year ahead: three goals, grow, keep and learn', slideCount: 5, textOnly: true, language: 'en', fields };

test('planOutline with composeLayouts makes the layout a slide asks for and gives it to that slide', async () => {
  const template = demo();
  const design = threeGoals(template);
  const { llm, requests } = planner({
    outline: textOnlyOutline,
    decision: { layouts: [{ instruction: 'Three goals side by side, each a heading and a sentence.', slides: [3, 1, 5, 9] }] },
    designs: [{ ...design, base: undefined }],
  });
  const stages: string[] = [];
  const plan = await planOutline(template, brief, { llm, layout: 'adaptive', composeLayouts: 2, onProgress: (event) => stages.push(event.stage) });

  assert.deepEqual(plan.slides.map((slide) => slide.pattern), ['cover', 'text-only', 'three-goals', 'text-only', 'closing'], 'only slide 3: the first and last slides and unknown numbers are left alone');
  assert.equal(plan.layouts?.length, 1);
  assert.equal(plan.layouts?.[0]?.base, 'text-only', 'it stands on the frame of the layout slide 3 had');
  assert.deepEqual(stages, ['outline', 'layouts']);
  const decision = requests.find((request) => request.system.startsWith('You review the plan'));
  assert.equal(decision?.effort, 'low');
  assert.ok((decision?.maxOutputTokens ?? Infinity) <= 1500, 'the decision has a cap on its answer');
  assert.match(decision?.prompt ?? '', /text-only on 3 slides/, 'the decision is told which layout repeats');
  const compose = requests.find((request) => request.system.startsWith('You design one new slide layout'));
  assert.match(compose?.prompt ?? '', /stands on the frame of the layout "text-only"/);

  // Writing the deck: the new layout is offered like the template's own, and the deck keeps its design.
  const content = {
    s1: { title: 'The year ahead' },
    s2: { title: 'Where we stand', body: ['We grew in two regions.'] },
    s3: goalSlots,
    s4: { title: 'What it costs', body: ['Two hires and a budget for training.'] },
    s5: { title: 'Thank you' },
  };
  const written = await writeDeck(template, plan, { llm: planner({ outline: null, decision: null, content }).llm, layout: 'adaptive' });
  assert.deepEqual(written.deck.layouts, plan.layouts);
  assert.equal(written.deck.slides[2]?.slots.goal2Heading, 'Keep');
  const rendered = renderDeck(template, written.deck, { layout: 'adaptive' });
  assert.ok(slideXml(rendered.pptx, 3).includes('Train every team twice.'));
});

test('a layout that cannot be made leaves its slides as planned, and the deck is still made', async () => {
  const template = demo();
  const broken = { id: 'three-goals', description: 'x', keep: ['title'], elements: [{ name: 'goal1Text', style: 'T99', role: 'text', optional: false, sample: 'x', column: 1, row: 6, width: 4, height: 4 }] };
  const { llm } = planner({ outline: textOnlyOutline, decision: { layouts: [{ instruction: 'Three goals.', slides: [3] }] }, designs: [broken, broken] });
  const result = await composeLayouts(template, await planOutline(template, brief, { llm, layout: 'adaptive' }), { llm, layout: 'adaptive' });
  assert.deepEqual(result.made, []);
  assert.equal(result.failed.length, 1);
  assert.deepEqual(result.failed[0]?.slides, [3]);
  assert.match(result.failed[0]?.reason ?? '', /T99/);
  assert.equal(result.plan.layouts, undefined);
  assert.equal(result.plan.slides[2]?.pattern, 'text-only');

  // A failing decision is no reason to fail the deck either.
  const failing: LlmProvider = {
    async generateJson(request) {
      if (request.system.startsWith('You review the plan')) throw new Error('the model is down');
      return textOnlyOutline;
    },
  };
  const plan = await planOutline(template, brief, { llm: failing, layout: 'adaptive', composeLayouts: 1 });
  assert.deepEqual(plan.slides.map((slide) => slide.pattern), textOnlyOutline.slides.map((slide) => slide.pattern));
});

test('an empty answer from the decision makes nothing and calls no designer', async () => {
  const template = demo();
  const { llm, requests } = planner({ outline: textOnlyOutline, decision: { layouts: [] } });
  const plan = await planOutline(template, brief, { llm, layout: 'adaptive', composeLayouts: 2 });
  assert.equal(plan.layouts, undefined);
  assert.equal(requests.length, 2, 'the outline and the decision');
});

test('two layouts that chose the same id are both kept, the second under a free id', async () => {
  const template = demo();
  const design = { ...threeGoals(template), base: undefined };
  const { llm } = planner({
    outline: textOnlyOutline,
    decision: { layouts: [{ instruction: 'Three goals.', slides: [2] }, { instruction: 'Three costs.', slides: [4] }] },
    designs: [design, design],
  });
  const plan = await planOutline(template, brief, { llm, layout: 'adaptive', composeLayouts: 2 });
  assert.deepEqual(plan.layouts?.map((layout) => layout.id), ['three-goals', 'three-goals-2']);
  assert.deepEqual(plan.slides.map((slide) => slide.pattern), ['cover', 'three-goals', 'text-only', 'three-goals-2', 'closing']);
  // Both build onto the template together, as rendering will build them.
  const both = withLayouts(template, plan.layouts);
  assert.ok(['three-goals', 'three-goals-2'].every((id) => both.manifest.patterns.some((pattern) => pattern.id === id)));
});

test('an edit keeps the deck\'s layouts while a slide uses them, and drops one no slide uses', async () => {
  const template = demo();
  const deck = deckWith(threeGoals(template));
  const keep: LlmProvider = {
    async generateJson() {
      return { summary: 'Moved the goals', language: 'en', title: 'T', textSize: 'normal', fields, slides: [{ from: 1, pattern: 'cover', change: '' }, { from: 3, pattern: 'closing', change: '' }, { from: 2, pattern: 'three-goals', change: '' }] };
    },
  };
  const kept = await editDeck(template, deck, 'Put the goals last', { llm: keep, layout: 'adaptive' });
  assert.deepEqual(kept.deck.slides.map((slide) => slide.pattern), ['cover', 'closing', 'three-goals']);
  assert.deepEqual(kept.deck.layouts, deck.layouts);

  const drop: LlmProvider = {
    async generateJson() {
      return { summary: 'Removed the goals', language: 'en', title: 'T', textSize: 'normal', fields, slides: [{ from: 1, pattern: 'cover', change: '' }, { from: 3, pattern: 'closing', change: '' }] };
    },
  };
  const dropped = await editDeck(template, deck, 'Remove the goals', { llm: drop, layout: 'adaptive' });
  assert.equal(dropped.deck.layouts, undefined);
});

test('planDeck passes composeLayouts through, and the option is off unless set', async () => {
  const template = demo();
  const { llm, requests } = planner({
    outline: { title: 'T', slides: [{ pattern: 'cover', purpose: 'Open' }, { pattern: 'text-only', purpose: 'Goals' }, { pattern: 'closing', purpose: 'Thanks' }] },
    decision: { layouts: [] },
    content: { s1: { title: 'T' }, s2: { title: 'Goals', body: ['Grow'] }, s3: { title: 'Thanks' } },
  });
  await planDeck(template, { ...brief, slideCount: 3 }, { llm, layout: 'adaptive' });
  assert.ok(!requests.some((request) => request.system.startsWith('You review the plan')), 'no decision without the option');
  await planDeck(template, { ...brief, slideCount: 3 }, { llm, layout: 'adaptive', composeLayouts: 1 });
  assert.ok(requests.some((request) => request.system.startsWith('You review the plan')));
});

test('a plan\'s own layouts survive editPlan while a slide uses them', async () => {
  const template = demo();
  const plan: DeckPlan = {
    version: 1,
    title: 'T',
    brief,
    layouts: [threeGoals(template)],
    slides: [{ pattern: 'cover', purpose: 'Open' }, { pattern: 'three-goals', purpose: 'Goals' }, { pattern: 'closing', purpose: 'Thanks' }],
  };
  const { editPlan } = await import('../src/index.ts');
  const answer = (slides: unknown[]) => ({ generateJson: async () => ({ summary: 's', language: 'en', title: 'T', slides }) });
  const kept = await editPlan(template, plan, 'Add a slide on costs', { llm: answer([...plan.slides.slice(0, 2), { pattern: 'text-only', purpose: 'Costs' }, plan.slides[2]]), layout: 'adaptive' });
  assert.deepEqual(kept.plan.layouts, plan.layouts);
  const dropped = await editPlan(template, plan, 'Drop the goals', { llm: answer([plan.slides[0], plan.slides[2]]), layout: 'adaptive' });
  assert.equal(dropped.plan.layouts, undefined);
});

test('the chat tools make new layouts for a plan when asked to, mark them, and render the deck', async () => {
  const template = demo();
  const design = threeGoals(template);
  const content = { s1: { title: 'The year ahead' }, s2: { title: 'Where we stand', body: ['We grew.'] }, s3: goalSlots, s4: { title: 'Costs', body: ['Two hires.'] }, s5: { title: 'Thank you' } };
  const { llm } = planner({ outline: textOnlyOutline, decision: { layouts: [{ instruction: 'Three goals.', slides: [3] }] }, designs: [design], content });
  const { presentationTools, memoryStore, TOOL_NAMES } = await import('../src/index.ts');
  const tools = presentationTools({ template, llm, store: memoryStore(), fields, layout: 'adaptive', composeLayouts: 1 });
  const planned = await tools.call(TOOL_NAMES.plan, { topic: brief.prompt, text_only: true, slide_count: 5, language: 'en' });
  const slides = planned.output.slides as Array<{ layout: string; layout_description?: string; new_layout?: boolean }>;
  assert.equal(slides[2]?.layout, 'three-goals');
  assert.equal(slides[2]?.new_layout, true);
  assert.equal(slides[2]?.layout_description, design.description);
  assert.equal(slides[1]?.new_layout, undefined);
  const created = await tools.call(TOOL_NAMES.create, { plan_id: planned.output.plan_id });
  assert.equal(created.output.error, undefined, String(created.output.error));
  assert.equal(created.deck?.layouts?.[0]?.id, 'three-goals');
  assert.ok(created.pptx && slideXml(created.pptx, 3).includes('Win two new regions.'));
});

test('only slides of content may get a new layout: a section divider keeps its own', async () => {
  const template = demo();
  const outline = { title: 'T', slides: [{ pattern: 'cover', purpose: 'Open' }, { pattern: 'text-only', purpose: 'Goals' }, { pattern: 'section', purpose: 'Part two' }, { pattern: 'quote', purpose: 'A word' }, { pattern: 'text-only', purpose: 'Costs' }, { pattern: 'closing', purpose: 'Thanks' }] };
  const { llm, requests } = planner({ outline, decision: { layouts: [{ instruction: 'Three goals.', slides: [3, 4, 2] }] }, designs: [{ ...threeGoals(template), base: undefined }] });
  const plan = await planOutline(template, { ...brief, slideCount: 6 }, { llm, layout: 'adaptive', composeLayouts: 1 });
  assert.match(requests[1]?.prompt ?? '', /Only slides 2, 5 may get one/);
  assert.deepEqual(plan.slides.map((slide) => slide.pattern), ['cover', 'three-goals', 'section', 'quote', 'text-only', 'closing']);

  // A deck with no slide of content asks nothing.
  const bare = planner({ outline: { title: 'T', slides: [outline.slides[0], outline.slides[2], outline.slides[5]] }, decision: { layouts: [] } });
  await planOutline(template, { ...brief, slideCount: 3 }, { llm: bare.llm, layout: 'adaptive', composeLayouts: 1 });
  assert.equal(bare.requests.length, 1);
});

test('a composed text drawn larger than its sample tells the writer the length it was drawn for', () => {
  const template = demo();
  const built = buildPattern(template, threeGoals(template));
  const slot = built.manifest.patterns.find((pattern) => pattern.id === 'three-goals')?.slots.find((candidate) => candidate.name === 'goal1Text');
  assert.ok(slot && slot.kind === 'text' && (slot.maxChars ?? 0) >= 75, 'the box holds well over the 51 characters of the sample');
  assert.equal(slot.description, 'Drawn for about 50 characters');
  const heading = built.manifest.patterns.find((pattern) => pattern.id === 'three-goals')?.slots.find((candidate) => candidate.name === 'goal1Heading');
  assert.equal(heading?.description, undefined, 'a heading is filled by its nature');
});
