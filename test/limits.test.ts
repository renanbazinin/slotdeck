// What the engine does with a deck, a manifest or an asset that is far outside what a slide
// takes: it answers with a sentence, in time, and never writes a file PowerPoint cannot open.
import assert from 'node:assert/strict';
import { readFileSync } from 'node:fs';
import { dirname, join } from 'node:path';
import { test } from 'node:test';
import { fileURLToPath } from 'node:url';
import { strFromU8, unzipSync } from 'fflate';
import { placeholderPhoto } from '../examples/placeholder.ts';
import { textScaleProblem } from '../src/deck/validate.ts';
import { freeSlide } from '../src/free/blocks.ts';
import { freeFor } from '../src/free/resolve.ts';
import {
  DeckValidationError,
  PlanError,
  buildPattern,
  changeManifest,
  checkDesign,
  checkManifest,
  composeLayouts,
  editPlan,
  loadTemplate,
  planEdit,
  planOutline,
  renderDeck,
  resolveAssets,
  sampleDeck,
  validateDeck,
  writeDeck,
} from '../src/index.ts';
import type { Brief, ChartValue, DeckPlan, DeckSpec, FontMetrics, JsonRequest, LlmProvider, ManifestChange, PatternDesign, SlideSpec, TemplateManifest } from '../src/index.ts';
import { openableSizes } from '../src/pptx/xml.ts';
import { breakLines } from '../src/text/layout.ts';
import { buildFixturePptx, fixtureManifest, tableManifest } from './fixture.ts';

const root = join(dirname(fileURLToPath(import.meta.url)), '..');
const demoManifest = JSON.parse(readFileSync(join(root, 'templates', 'demo.manifest.json'), 'utf8')) as TemplateManifest;
const demoPptx = new Uint8Array(readFileSync(join(root, 'templates', 'demo.pptx')));
const demo = loadTemplate(demoPptx, demoManifest);
const fixture = buildFixturePptx();
const square = placeholderPhoto(100, 100, [10, 20, 30], [200, 210, 220]);

const contentSlide = (slots: SlideSpec['slots'] = {}): SlideSpec => ({ pattern: 'content', slots: { title: 'Title', body: ['One', 'Two'], photo: { asset: 'square' }, note: 'Note', ...slots } });
const demoDeck = (slides: SlideSpec[], more: Record<string, unknown> = {}): DeckSpec => ({ version: 1, fields: { organizationDepartment: 'Riverbend' }, ...more, slides }) as DeckSpec;
const series = (count: number): ChartValue => ({ type: 'column', categories: ['Q1'], series: Array.from({ length: count }, (_, index) => ({ name: `S${index + 1}`, values: [index + 1] })) });
const table = { header: ['Station', 'Docks'], rows: [['Tram Square', '24'], ['River Gate', '18']] };
/** How long a call takes, in milliseconds. */
const timed = (run: () => unknown): number => {
  const started = performance.now();
  run();
  return performance.now() - started;
};

test('a chart of more series than PowerPoint opens is an error, on a layout and on a free slide', () => {
  const onLayout = (count: number) => validateDeck(demo, demoDeck([{ pattern: 'chart-native', slots: { title: 'Rides', chart: series(count), chartDescription: 'Rides in thousands.' } }]), { layout: 'adaptive' });
  const onFree = (count: number) => validateDeck(demo, demoDeck([freeSlide({ title: 'Rides', blocks: [{ chart: series(count) }] })]));
  for (const check of [onLayout, onFree]) {
    // 255 series are far more than read well, which is a warning; one more is a file that does not open.
    assert.deepEqual(check(255).filter((issue) => issue.severity === 'error'), []);
    assert.deepEqual(check(256).filter((issue) => issue.severity === 'error').map((issue) => [issue.slide, issue.message]), [[1, 'the chart has 256 series; a chart takes at most 255']]);
  }
  assert.throws(() => renderDeck(demo, demoDeck([freeSlide({ blocks: [{ chart: series(300) }] })])), (error: unknown) => error instanceof DeckValidationError && /300 series; a chart takes at most 255/.test(error.message));
});

test('every text size written is one PowerPoint opens, whatever the template and the deck multiply', () => {
  // The size written is the template's, times the manifest's factor, times the deck's. At the ends
  // of what a deck may ask, with the ends of what a manifest may say, it would be past 4000 pt or
  // under 1 pt, and PowerPoint would not open the file.
  const sizesOf = (pptx: Uint8Array): number[] => {
    const parts = unzipSync(pptx);
    return Object.entries(parts)
      .filter(([name]) => /^ppt\/(?:slides\/slide|charts\/chart)\d+\.xml$/.test(name))
      .flatMap(([, data]) => [...strFromU8(data).matchAll(/<a:(?:rPr|defRPr|endParaRPr)\b[^>]*?\ssz="([^"]*)"/g)].map((match) => Number(match[1])));
  };
  const free = freeSlide({ title: 'Free', blocks: [{ text: 'A text.' }, { card: { heading: 'Head', text: 'Text' } }, { table }, { chart: series(2) }, { figure: { value: '87%', label: 'would recommend it' } }] });
  for (const [factor, textScale] of [[3, 300], [0.5, 0.1]] as const) {
    const template = loadTemplate(demoPptx, { ...demoManifest, textScale: factor });
    const { deck, assets } = sampleDeck(template);
    const sizes = sizesOf(renderDeck(template, { ...deck, textScale, slides: [...deck.slides, free] }, { assets, layout: 'adaptive' }).pptx);
    assert.ok(sizes.length > 100);
    assert.deepEqual(sizes.filter((size) => !(Number.isInteger(size) && size >= 100 && size <= 400_000)), [], `manifest ${factor}, deck ${textScale}`);
    // Held at the edge, not dropped: the largest is 4000 pt, the smallest 1 pt.
    assert.ok(factor === 3 ? sizes.includes(400_000) : sizes.includes(100));
  }
  // A template factor that is no number has no nearest size: the deck is stopped, and checkManifest says so.
  for (const bad of [Number.NaN, Infinity, 0, -1, 'abc', {}]) {
    const manifest = { ...demoManifest, textScale: bad } as unknown as TemplateManifest;
    const template = loadTemplate(demoPptx, manifest);
    const deck = sampleDeck(loadTemplate(demoPptx, demoManifest)).deck;
    const expected = `the template's textScale is ${typeof bad === 'number' ? String(bad) : JSON.stringify(bad)}; it is a factor on body text such as 1.25, which changeManifest sets with "textScale"`;
    assert.deepEqual(validateDeck(template, deck).map((issue) => [issue.severity, issue.slide, issue.message]), [['error', undefined, expected]], String(bad));
    assert.throws(() => renderDeck(template, deck, { assets: sampleDeck(template).assets }), DeckValidationError);
    assert.ok(checkManifest(template).some((issue) => issue.severity === 'error' && issue.message === expected), String(bad));
  }
  // Left out, null, or a number written as text: the template renders as it always did.
  for (const fine of [undefined, null, '1.25', 1.25]) assert.deepEqual(validateDeck(loadTemplate(demoPptx, { ...demoManifest, textScale: fine } as unknown as TemplateManifest), sampleDeck(demo).deck).filter((issue) => issue.severity === 'error'), [], String(fine));
  // Sizes inside the range are left as written: a part with none outside it is the same text.
  assert.equal(openableSizes('<a:rPr lang="en-US" sz="1200"/><a:defRPr sz="100"/><a:endParaRPr sz="400000"/>'), '<a:rPr lang="en-US" sz="1200"/><a:defRPr sz="100"/><a:endParaRPr sz="400000"/>');
  assert.equal(openableSizes('<a:rPr lang="en-US" sz="546000" b="1"/><a:defRPr sz="50"/>'), '<a:rPr lang="en-US" sz="400000" b="1"/><a:defRPr sz="100"/>');
});

test('a text size that is no size stops the deck; one that is not set is the usual one', () => {
  const template = loadTemplate(fixture, fixtureManifest);
  const check = (textScale: unknown) => validateDeck(fixtureManifest, { version: 1, textScale, slides: [contentSlide()] } as unknown as DeckSpec, { assets: ['square'] });
  for (const none of [Number.NaN, 0, -1, 0.05, 301, 1000, 1e9, Infinity, {}, [], true, 'large', '', '1e9']) {
    const issues = check(none);
    assert.equal(issues.length, 1, JSON.stringify(none));
    assert.deepEqual([issues[0]?.severity, issues[0]?.slide, issues[0]?.slot], ['error', undefined, undefined]);
    assert.match(issues[0]?.message ?? '', /^the deck's textScale is .+; it is a factor on body text between 0\.1 and 300, such as 1\.25$/);
  }
  assert.equal(textScaleProblem(Number.NaN), "the deck's textScale is NaN; it is a factor on body text between 0.1 and 300, such as 1.25");
  assert.equal(textScaleProblem('large'), 'the deck\'s textScale is "large"; it is a factor on body text between 0.1 and 300, such as 1.25');
  const render = (textScale: unknown) => renderDeck(template, { version: 1, textScale, slides: [contentSlide()] } as unknown as DeckSpec, { assets: { square } }).pptx;
  assert.throws(() => render(Number.NaN), DeckValidationError);
  assert.throws(() => render(1e9), DeckValidationError);
  // Factors a deck can be made with, also ones nobody would choose: nothing of them is an error.
  // 0.1 and 300 are the ends at which PowerPoint opens the file, and 0.5 rendered them, so a stored deck with them still renders.
  for (const fine of [0.1, 0.25, 0.3, 1, 1.25, 4, 10, 30, 300, '2']) assert.deepEqual(check(fine).filter((issue) => issue.severity === 'error'), [], String(fine));
  // Null is how an application stores a number that is not set: the deck is the deck without one.
  assert.deepEqual(check(null), []);
  assert.deepEqual(render(null), render(undefined));
  // A number written as text is read as that number, as it always was.
  assert.deepEqual(render('2'), render(2));
  // The same on a free slide, whose sizes are worked out from the factor.
  const free = (textScale: unknown) => renderDeck(demo, demoDeck([freeSlide({ title: 'Three', blocks: [{ text: 'A text.' }, { table }] })], { textScale })).pptx;
  assert.deepEqual(free(null), free(undefined));
  assert.deepEqual(free('1.5'), free(1.5));
  assert.throws(() => free(Infinity), DeckValidationError);
});

test('a free slide with a table is worked out in time whatever the text size, and blocks that are no blocks give a sentence', () => {
  const slide = freeSlide({ title: 'Stations', blocks: [{ table }, { text: 'A text.' }, { chart: series(2) }, { figure: { value: '87%', label: 'would recommend it' } }] });
  // Asked without the deck's check, as the design checks ask: a factor of a million, and one that is no number at all.
  for (const textScale of [1e6, 1e9, Infinity]) {
    const took = timed(() => assert.equal(typeof freeFor(demo, { textScale }, slide), 'object'));
    assert.ok(took < 5000, `a text size of ${textScale} is worked out in ${Math.round(took)} ms`);
  }
  // A table style of a size that is no size, in a manifest written by hand.
  const huge = structuredClone(demoManifest);
  const slot = huge.patterns.find((pattern) => pattern.id === 'table')?.slots.find((entry) => entry.kind === 'table');
  assert.ok(slot?.kind === 'table');
  slot.bodyText = { ...(slot.bodyText ?? {}), size: 1e9 };
  const took = timed(() => assert.equal(typeof freeFor(loadTemplate(demoPptx, huge), {}, slide), 'object'));
  assert.ok(took < 5000, `a table style of a thousand million points is worked out in ${Math.round(took)} ms`);

  // Blocks that are no list of kinds are answered with a sentence, by whoever asks.
  for (const blocks of [undefined, null, 'text', 7, {}, [], ['__proto__'], ['constructor'], ['toString'], [7], [null], [{}], ['TEXT']]) {
    assert.equal(freeFor(demo, {}, { blocks, slots: {} } as unknown as SlideSpec), 'its "blocks" are not a list of kinds of block', JSON.stringify(blocks));
  }
  assert.match(freeFor(demo, {}, { blocks: Array.from({ length: 2001 }, () => 'note'), slots: {} } as unknown as SlideSpec) as string, /^it names 2001 blocks, and a free slide holds at most 2000$/);
  // The deck's check says what is wrong, as an error of the slide, and in no time.
  const many = demoDeck([{ pattern: 'free', blocks: Array.from({ length: 150000 }, () => 'note' as const), slots: {} }]);
  let issues: ReturnType<typeof validateDeck> = [];
  assert.ok(timed(() => (issues = validateDeck(demo, many))) < 2000);
  assert.deepEqual(issues.map((issue) => [issue.severity, issue.slide, issue.message]), [['error', 1, 'a free slide holds at most 2000 blocks, not 150000']]);
  assert.deepEqual(validateDeck(demoManifest, many).map((issue) => issue.message), ['a free slide holds at most 2000 blocks, not 150000']);
  // Two thousand are still a slide.
  assert.deepEqual(validateDeck(demoManifest, demoDeck([{ pattern: 'free', blocks: Array.from({ length: 2000 }, () => 'note' as const), slots: {} }])).filter((issue) => !/required/.test(issue.message)), []);
});

test('the design checks pass over a slide that cannot be rendered: items that are no number of items, blocks that are no blocks', () => {
  const fixedWidth: FontMetrics = { face: () => ({ has: () => true, advances: (text) => [...text].map(() => 0.5) }) };
  const measured = loadTemplate(demoPptx, demoManifest, { fonts: fixedWidth });
  const cards = (items: unknown): SlideSpec => ({ pattern: 'icon-cards', items, slots: { title: 'Cards', card1Heading: 'One', card1Text: 'A text.', card2Heading: 'Two', card2Text: 'A text.' } }) as unknown as SlideSpec;
  for (const items of [Infinity, 1e9, 1e6, -1, 2.5, Number.NaN, 'two']) {
    let issues: unknown[] = ['not run'];
    const took = timed(() => (issues = checkDesign(measured, demoDeck([cards(items)]), { layout: 'adaptive' })));
    assert.deepEqual(issues, [], String(items));
    assert.ok(took < 2000, `${items} items are passed over in ${Math.round(took)} ms`);
  }
  for (const blocks of [['__proto__'], ['constructor'], [7], [null], 'text', {}]) {
    assert.deepEqual(checkDesign(measured, demoDeck([{ pattern: 'free', blocks, slots: {} } as unknown as SlideSpec])), [], JSON.stringify(blocks));
  }
  // A slide that can be rendered is still checked.
  assert.ok(Array.isArray(checkDesign(measured, demoDeck([cards(2)]), { layout: 'adaptive' })));
});

test('a table of a hundred thousand cells is checked like any other', () => {
  const header = Array.from({ length: 501 }, () => 'h');
  const rows = Array.from({ length: 300 }, () => Array.from({ length: 501 }, () => 'x'));
  const issues = validateDeck(tableManifest, { version: 1, slides: [{ pattern: 'table', slots: { title: 'Breeds', table: { header, rows } } }] });
  assert.deepEqual([...new Set(issues.map((issue) => issue.severity))], ['warning']);
  assert.ok(issues.some((issue) => /the table has 501 columns/.test(issue.message)) && issues.some((issue) => /the table has 300 rows/.test(issue.message)));
  // On a free slide too, where its cells are counted for what its box holds.
  const wide = { header: Array.from({ length: 400 }, () => 'h'), rows: Array.from({ length: 400 }, () => Array.from({ length: 400 }, () => 'x')) };
  assert.equal(typeof freeFor(demo, { textScale: 1.25 }, freeSlide({ blocks: [{ table: wide }] })), 'object');
});

test('a chart of two hundred thousand values and a table of two hundred thousand rows on a free slide are checked, not thrown', () => {
  // A list that long is more arguments than a call takes: the most and least of it are counted in a loop.
  const count = 200000;
  const chart: ChartValue = { type: 'column', categories: Array.from({ length: count }, (_, index) => `c${index}`), series: [{ name: 'Rides', values: Array.from({ length: count }, (_, index) => index) }] };
  const long = { header: ['Station'], rows: Array.from({ length: count }, (_, index) => [`r${index}`]) };
  const words = { header: ['Words'], rows: [[Array.from({ length: count }, () => 'a').join(' ')]] };
  for (const blocks of [[{ chart }], [{ table: long }], [{ table: words }]]) {
    const issues = validateDeck(demo, demoDeck([freeSlide({ title: 'Rides', blocks })]));
    assert.deepEqual(issues.filter((issue) => issue.severity === 'error'), []);
    assert.ok(issues.length > 0, 'what the slide does not hold is said');
  }
  assert.ok(renderDeck(demo, demoDeck([freeSlide({ title: 'Rides', blocks: [{ chart }] })])).pptx.length > 0);
});

test('a deck of the wrong shape is told what is wrong by the check and the renderer, and passed over by the design checks', () => {
  const cover: SlideSpec = { pattern: 'cover', slots: { title: 'A year of rides', subtitle: 'What the season taught us' } };
  const notDeck = /^the deck is no object with a list of slides, such as \{ "version": 1, "slides": \[\.\.\.\] \}$/;
  const notSlide = /^the slide is no object with a pattern and slots, such as \{ "pattern": "cover", "slots": \{ \.\.\. \} \}$/;
  const layouts = (given: string) => new RegExp(`^"layouts" is ${given}; it is a list of the deck's own layouts, such as \\[\\{ "id": "three-goals", "elements": \\[\\.\\.\\.\\] \\}\\]$`);
  const cases: Array<[unknown, number | undefined, RegExp]> = [
    [null, undefined, notDeck],
    [undefined, undefined, notDeck],
    [5, undefined, notDeck],
    [[cover], undefined, notDeck],
    [{ version: 1 }, undefined, /^the deck has no slides$/],
    [{ version: 1, slides: [null] }, 1, notSlide],
    [demoDeck([cover, null as unknown as SlideSpec]), 2, notSlide],
    [demoDeck([cover, 'cover' as unknown as SlideSpec]), 2, notSlide],
    [demoDeck([cover], { layouts: 5 }), undefined, layouts('a number')],
    [demoDeck([cover], { layouts: 'three-goals' }), undefined, layouts('a string')],
    [demoDeck([cover], { layouts: {} }), undefined, layouts('an object')],
    [demoDeck([cover], { layouts: [null] }), undefined, layouts('a list that holds something other than layouts')],
  ];
  const fixedWidth: FontMetrics = { face: () => ({ has: () => true, advances: (text) => [...text].map(() => 0.5) }) };
  const measured = loadTemplate(demoPptx, demoManifest, { fonts: fixedWidth });
  for (const [deck, slide, message] of cases) {
    const label = JSON.stringify(deck) ?? String(deck);
    for (const target of [demoManifest, demo]) {
      const issues = validateDeck(target, deck as DeckSpec);
      assert.equal(issues.length, 1, label);
      assert.deepEqual([issues[0]?.severity, issues[0]?.slide], ['error', slide], label);
      assert.match(issues[0]?.message ?? '', message, label);
    }
    assert.throws(() => renderDeck(demo, deck as DeckSpec), (error: unknown) => error instanceof DeckValidationError && error.issues.length === 1 && message.test(error.issues[0]?.message ?? ''), label);
    assert.deepEqual(checkDesign(measured, deck as DeckSpec), [], label);
  }
  // Layouts left out, or null as an application stores a list that is not set, are no layouts.
  assert.deepEqual(validateDeck(demo, demoDeck([cover], { layouts: null })), validateDeck(demo, demoDeck([cover])));
  assert.deepEqual(renderDeck(demo, demoDeck([cover], { layouts: null })).pptx, renderDeck(demo, demoDeck([cover])).pptx);
});

test('a brief, a plan or a stored deck of the wrong shape gets an error of the library, or is read for what it holds', async () => {
  const asked: JsonRequest[] = [];
  const llm: LlmProvider = {
    async generateJson(request) {
      asked.push(request);
      throw new Error('stop here');
    },
  };
  // A brief whose prompt is no text: nothing is asked.
  for (const prompt of [5, null, undefined, {}, ['Rides']]) {
    const isPlanError = (error: unknown) => error instanceof PlanError && error.message === 'The brief has no prompt: "prompt" is the text of what the deck is about';
    await assert.rejects(planOutline(demo, { prompt, language: 'en' } as unknown as Brief, { llm }), isPlanError, String(prompt));
    await assert.rejects(writeDeck(demo, { version: 1, title: 'T', brief: { prompt }, slides: [{ pattern: 'cover', purpose: 'p' }] } as unknown as DeckPlan, { llm }), isPlanError, String(prompt));
  }
  // A plan whose slides are no list of objects: nothing is asked.
  for (const slides of [[null], [{ pattern: 'cover', purpose: 'p' }, 5], 'cover', undefined]) {
    const plan = { version: 1, title: 'T', brief: { prompt: 'Bike rides' }, slides } as unknown as DeckPlan;
    for (const run of [() => writeDeck(demo, plan, { llm }), () => editPlan(demo, plan, 'Add a slide', { llm }), () => composeLayouts(demo, plan, { llm })]) {
      await assert.rejects(run(), (error: unknown) => error instanceof PlanError && /^The plan (is no object with a list of slides|has slides that are no object with a pattern and a purpose: 1|has slides that are no object with a pattern and a purpose: 2)$/.test(error.message), JSON.stringify(slides));
    }
  }
  assert.equal(asked.length, 0);

  // A share of pictures that is not one of the list is the usual one: the model is asked as for "medium".
  const outline = async (imageDensity: unknown) => {
    const brief = { prompt: 'Bike rides', language: 'en', fields: { organizationDepartment: 'Riverbend' }, imageDensity } as unknown as Brief;
    await assert.rejects(planOutline(demo, brief, { llm, layout: 'adaptive' }), /^Error: stop here$/, String(imageDensity));
    return asked.at(-1)?.prompt;
  };
  const usual = await outline('medium');
  assert.match(usual ?? '', /a free slide with a photo counts/, 'free slides are on offer, and with them what is said of their photos');
  for (const density of ['many', 'toString', 'constructor', 5, {}]) assert.equal(await outline(density), usual, String(density));

  // A stored deck that names kinds of block that are none, or a slide stored without slots, is described as it is.
  const cover: SlideSpec = { pattern: 'cover', slots: { title: 'A year of rides', subtitle: 'What the season taught us' } };
  const slides = [{ pattern: 'free', blocks: ['constructor'], slots: { title: 'T' } }, { pattern: 'free', blocks: ['__proto__', 'text', 5], slots: { title: 'T' } }, { pattern: 'text-only' }, { pattern: 'text-only', slots: null }];
  for (const slide of slides) {
    await assert.rejects(planEdit(demo, demoDeck([cover, slide as unknown as SlideSpec]), 'Shorten the text', { llm, layout: 'adaptive' }), /^Error: stop here$/, JSON.stringify(slide));
  }
  assert.deepEqual(await resolveAssets(demo, demoDeck([cover, { pattern: 'text-and-photo' } as unknown as SlideSpec]), {}), {});
});

test('checkManifest says when patterns or slots are no lists of objects, and reads no further', () => {
  const check = (patterns: unknown) => checkManifest(loadTemplate(demoPptx, { ...demoManifest, patterns } as unknown as TemplateManifest)).map((issue) => [issue.severity, issue.pattern, issue.message]);
  const listOf = 'patterns is no list of patterns, such as [{ "id": "cover", "sourceSlide": 1, "slots": [...] }]';
  for (const patterns of [null, 'cover', {}, 5]) assert.deepEqual(check(patterns), [['error', undefined, listOf]], JSON.stringify(patterns));
  assert.deepEqual(check([null, 5]), [
    ['error', undefined, 'pattern 1 is no object with an id, a slide and slots'],
    ['error', undefined, 'pattern 2 is no object with an id, a slide and slots'],
  ]);
  const withSlots = (slots: unknown) => check(demoManifest.patterns.map((pattern) => (pattern.id === 'cover' ? { ...pattern, slots } : pattern)));
  for (const slots of [null, 5, {}, 'title']) assert.deepEqual(withSlots(slots), [['error', 'cover', 'slots is no list of slots, such as [{ "name": "title", "kind": "text", "shapeId": "2" }]']], JSON.stringify(slots));
  for (const slots of [[null], ['title'], [[]]]) assert.deepEqual(withSlots(slots), [['error', 'cover', 'slots holds something that is no slot: each is an object with a name, a kind and a shape']], JSON.stringify(slots));
  // A manifest of the right shape is checked as before.
  assert.deepEqual(checkManifest(demo).filter((issue) => issue.severity === 'error'), []);
});

test('an asset that is no picture is an error of the deck, with its slide and its slot', () => {
  const template = loadTemplate(fixture, fixtureManifest);
  for (const bad of [new Uint8Array(0), new Uint8Array([1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12]), 'text', null, new ArrayBuffer(8)]) {
    assert.throws(
      () => renderDeck(template, { version: 1, slides: [contentSlide(), contentSlide()] }, { assets: { square: bad as unknown as Uint8Array } }),
      (error: unknown) => error instanceof DeckValidationError && error.issues.length === 1 && error.issues[0]?.slide === 1 && error.issues[0].slot === 'photo' && error.issues[0].message === 'asset "square" is not a PNG, JPEG or GIF picture',
    );
  }
  // An icon of a deck in other colours is read for its colours first: the same answer.
  const iconDeck = demoDeck([{ pattern: 'icon-cards', items: 2, slots: { title: 'Two', card1Icon: { asset: 'clock' }, card1Heading: 'One', card1Text: 'A text.', card2Icon: { asset: 'clock' }, card2Heading: 'Two', card2Text: 'A text.' } }], { colors: { accent: '1565C0' } });
  assert.throws(() => renderDeck(demo, iconDeck, { assets: { clock: new Uint8Array([137, 80, 78, 71, 1, 2, 3]) } }), (error: unknown) => error instanceof DeckValidationError && /slide 1 \/ card1Icon: asset "clock" is not a PNG, JPEG or GIF picture/.test(error.message));
});

test('the look of a chart and of a table is checked when a manifest is changed, and a colour is stored as six hex digits', () => {
  const chartManifest: TemplateManifest = {
    ...tableManifest,
    patterns: [...tableManifest.patterns, { id: 'chart', sourceSlide: 1, description: 'A chart.', slots: [{ name: 'title', kind: 'text', shapeId: '2' }, { name: 'chart', kind: 'chart', shapeId: '4', text: { size: 9, color: '0C0C0C' } }] }],
  };
  const template = loadTemplate(fixture, chartManifest);
  const chartSlot = (patch: Record<string, unknown>): ManifestChange => ({ op: 'setSlot', pattern: 'chart', slot: 'chart', ...patch }) as ManifestChange;
  const tableSlot = (patch: Record<string, unknown>): ManifestChange => ({ op: 'setSlot', pattern: 'table', slot: 'table', ...patch }) as ManifestChange;
  const textSlot = (patch: Record<string, unknown>): ManifestChange => ({ op: 'setSlot', pattern: 'content', slot: 'body', ...patch }) as ManifestChange;
  const refused = changeManifest(template, [
    chartSlot({ text: { color: 'red' } }),
    chartSlot({ text: { color: '0C0C0C"/><x' } }),
    chartSlot({ text: { size: 0 } }),
    chartSlot({ text: { size: 1e9 } }),
    chartSlot({ text: { size: '12' } }),
    chartSlot({ text: 'bold' }),
    chartSlot({ lineColor: 'red' }),
    chartSlot({ colors: ['0E7C7B', 'red'] }),
    chartSlot({ colors: '0E7C7B' }),
    tableSlot({ bodyText: { color: 'red' } }),
    tableSlot({ bodyText: { size: 0 } }),
    tableSlot({ headerText: { size: Infinity } }),
    textSlot({ paragraphGap: Number.NaN }),
    textSlot({ paragraphGap: -1 }),
  ]);
  assert.deepEqual(refused.applied, []);
  assert.deepEqual(refused.rejected.map((entry) => entry.reason), [
    'the colour of text is six hex digits such as "0E7C7B", not "red"',
    'the colour of text is six hex digits such as "0E7C7B", not "0C0C0C\\"/><x"',
    'the size of text is between 4 and 200 points',
    'the size of text is between 4 and 200 points',
    'the size of text is between 4 and 200 points',
    'text is the look of a text, such as { "size": 12, "color": "1B2A41" }',
    'lineColor is six hex digits such as "0E7C7B", not "red"',
    'a colour of a chart is six hex digits such as "0E7C7B", not "red"',
    'colors is a list of colours, such as ["0E7C7B", "F4A259"]',
    'the colour of bodyText is six hex digits such as "0E7C7B", not "red"',
    'the size of bodyText is between 4 and 200 points',
    'the size of headerText is between 4 and 200 points',
    'paragraphGap is a number of points, 0 or more',
    'paragraphGap is a number of points, 0 or more',
  ]);
  assert.deepEqual(refused.manifest, chartManifest);
  // What is a colour and a size is taken, a colour written with # or in small letters as six hex digits.
  const made = changeManifest(template, [chartSlot({ text: { color: '#1e6b3a', size: 11 }, lineColor: '#d9d9d9', colors: ['#0e7c7b', 'F4A259'] }), tableSlot({ bodyText: { color: '1b2a41', size: 10 } }), textSlot({ paragraphGap: 6 })]);
  assert.deepEqual(made.rejected, []);
  const chart = made.manifest.patterns.find((pattern) => pattern.id === 'chart')?.slots.find((slot) => slot.kind === 'chart');
  assert.deepEqual(chart?.kind === 'chart' && [chart.text, chart.lineColor, chart.colors], [{ size: 11, color: '1E6B3A' }, 'D9D9D9', ['0E7C7B', 'F4A259']]);
  const body = made.manifest.patterns.find((pattern) => pattern.id === 'table')?.slots.find((slot) => slot.kind === 'table');
  assert.deepEqual(body?.kind === 'table' && [body.bodyText?.color, body.bodyText?.size], ['1B2A41', 10]);
  // The file made with that manifest holds no colour or size that is none.
  const pptx = renderDeck(loadTemplate(fixture, made.manifest), { version: 1, slides: [{ pattern: 'chart', slots: { title: 'T', chart: series(2) } }] }).pptx;
  assert.ok(pptx.length > 0);
});

test('checkManifest says when the colours of a manifest are given in a shape that is none', () => {
  const check = (colors: unknown) => checkManifest(loadTemplate(demoPptx, { ...demoManifest, colors } as unknown as TemplateManifest)).map((issue) => `${issue.severity}: ${issue.message}`);
  assert.deepEqual(check('F4A259'), ['error: colors is "F4A259"; it is an object such as { "accent": "0E7C7B" }, or false for a template that keeps its own colours']);
  assert.deepEqual(check(true), ['error: colors is true; it is an object such as { "accent": "0E7C7B" }, or false for a template that keeps its own colours']);
  assert.deepEqual(check(['F4A259']), ['error: colors is ["F4A259"]; it is an object such as { "accent": "0E7C7B" }, or false for a template that keeps its own colours']);
  assert.deepEqual(check({ main: 'F4A259' }), ['warning: colors has no setting called "main"; it is ignored']);
  assert.deepEqual(check({ accent: 'amber' }), ['error: colors.accent is "amber"; a colour is six hex digits such as "0E7C7B"']);
  for (const fine of [undefined, null, false, {}, { accent: 'F4A259' }, { accent: '0E7C7B', accent2: 'F4A259' }]) assert.deepEqual(check(fine), [], JSON.stringify(fine));
});

test('a sample deck holds no more than a slide shows, whatever the budgets of the manifest say', () => {
  const huge = structuredClone(tableManifest);
  for (const pattern of huge.patterns) {
    for (const slot of pattern.slots) {
      if (slot.kind === 'text' && !slot.deckField) Object.assign(slot, { maxChars: 1e9, ...(slot.maxParagraphs !== undefined ? { maxParagraphs: 1e9 } : {}) });
      if (slot.kind === 'table') Object.assign(slot, { maxColumns: 1e9, maxRows: 1e9, maxCellChars: 1e9 });
    }
  }
  let sample: ReturnType<typeof sampleDeck> | undefined;
  const took = timed(() => (sample = sampleDeck(loadTemplate(fixture, huge))));
  assert.ok(took < 5000, `the sample is made in ${Math.round(took)} ms`);
  const values = (sample as ReturnType<typeof sampleDeck>).deck.slides.flatMap((slide) => Object.values(slide.slots));
  const texts = values.filter((value): value is string | string[] => typeof value === 'string' || Array.isArray(value));
  assert.ok(texts.length > 0 && texts.every((value) => (Array.isArray(value) ? value.length <= 100 && value.join('').length <= 6000 : value.length <= 6000)));
  const tables = values.filter((value): value is { header: string[]; rows: string[][] } => typeof value === 'object' && value !== null && 'header' in value);
  assert.ok(tables.length > 0 && tables.every((value) => value.header.length === 50 && value.rows.length === 200 && (value.header[0] as string).length <= 1000));
  // A manifest with budgets of its own gives the sample it always gave.
  const usual = sampleDeck(loadTemplate(fixture, tableManifest)).deck.slides.flatMap((slide) => Object.values(slide.slots));
  const usualTable = usual.find((value): value is { header: string[]; rows: string[][] } => typeof value === 'object' && value !== null && 'header' in value);
  const slot = tableManifest.patterns.flatMap((pattern) => pattern.slots).find((entry) => entry.kind === 'table');
  assert.deepEqual([usualTable?.header.length, usualTable?.rows.length], [slot?.kind === 'table' ? (slot.maxColumns ?? 3) : 0, slot?.kind === 'table' ? (slot.maxRows ?? 3) : 0]);
});

test('a design of more parts than its grid has cells is refused in one sentence', () => {
  const design: PatternDesign = {
    id: 'too-many',
    description: 'A thousand parts.',
    keep: ['title'],
    elements: Array.from({ length: 1000 }, (_, index) => ({ name: `part${index}Text`, style: 'T1', role: 'text' as const, optional: false, sample: 'x', column: 1, row: 4, width: 4, height: 4 })),
  };
  let message = '';
  const took = timed(() => {
    try {
      buildPattern(demo, design);
    } catch (error) {
      message = (error as Error).message;
    }
  });
  assert.match(message, /The layout has 1000 parts; a layout takes at most 144\./);
  assert.ok(!/lie on each other/.test(message), 'no pair of parts is compared');
  assert.ok(took < 2000, `refused in ${Math.round(took)} ms`);
});

test('a word of thousands of characters is cut into lines as before, without measuring what is left of it after every line', () => {
  /** The way lines were cut before: what is left of the word is measured whole after every line. */
  const simple = (word: string, first: number, rest: number, measure: (text: string) => number): string[] => {
    const lines: string[] = [];
    let chars = [...word];
    while (chars.length > 0 && measure(chars.join('')) > (lines.length === 0 ? first : rest)) {
      let count = 1;
      while (count < chars.length && measure(chars.slice(0, count + 1).join('')) <= (lines.length === 0 ? first : rest)) count++;
      lines.push(chars.slice(0, count).join(''));
      chars = chars.slice(count);
    }
    lines.push(chars.join(''));
    return lines;
  };
  // Letters of very different widths, so that a line of narrow ones holds four times as many as a line of wide ones.
  const widths: Record<string, number> = { W: 1, m: 0.9, x: 0.5, i: 0.25, '.': 0.2 };
  const width = (text: string) => [...text].reduce((sum, char) => sum + (widths[char] ?? 0.5), 0);
  let seed = 7;
  const random = () => (seed = (seed * 16807) % 2147483647) / 2147483647;
  for (let round = 0; round < 60; round++) {
    let word = '';
    const length = 1 + Math.floor(random() * 400);
    while (word.length < length) word += (['W', 'm', 'x', 'i', '.'][Math.floor(random() * 5)] as string).repeat(1 + Math.floor(random() * 30));
    const first = 3 + random() * 20;
    const rest = 3 + random() * 20;
    assert.deepEqual(breakLines(word, first, rest, width), simple(word, first, rest, width), `round ${round}`);
  }
  // Twenty thousand characters: what is handed to the measure grows with the length of the word, not with its square.
  let measured = 0;
  const counting = (text: string) => {
    measured += text.length;
    return text.length;
  };
  const lines = breakLines('x'.repeat(20000), 30, 30, counting);
  assert.equal(lines.length, Math.ceil(20000 / 30));
  assert.ok(lines.every((line, index) => line.length === (index < lines.length - 1 ? 30 : 20000 - 30 * (lines.length - 1))));
  assert.ok(measured < 20000 * 60, `${measured} characters were measured for a word of 20,000`);
});
