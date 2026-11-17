// Layout freedom that stays inside the template: a row with fewer or more items, items that
// widen to fill their row, larger body text, and patterns derived from a slide by moving and
// resizing a text box.
import assert from 'node:assert/strict';
import { test } from 'node:test';
import { placeholderPhoto } from '../examples/placeholder.ts';
import {
  PptxPackage,
  TOOL_NAMES,
  changeManifest,
  checkManifest,
  describeManifest,
  editManifest,
  extractAssets,
  loadTemplate,
  memoryStore,
  editPlan as changePlan,
  planDeck,
  planEdit,
  planOutline,
  presentationTools,
  renderDeck,
  sampleDeck,
  validateDeck,
} from '../src/index.ts';
import type { DeckSpec, JsonRequest, LlmProvider, SlideSpec, TemplateManifest } from '../src/index.ts';
import { readRels, resolveTarget } from '../src/pptx/package.ts';
import { NS, childElement, childElements, descendants, parseXml, path, rootElement } from '../src/pptx/xml.ts';
import { indexShapes } from '../src/render/render.ts';
import { applyRepeat } from '../src/render/repeat.ts';
import type { Element } from '../src/pptx/xml.ts';
import { effectiveSlots, rowLayout, slotSources, textBudget } from '../src/template/manifest.ts';
import { buildFixturePptx, cardsManifest, fixtureManifest, gridManifest } from './fixture.ts';

const pptx = buildFixturePptx();
const template = loadTemplate(pptx, cardsManifest);
const cards = cardsManifest.patterns.at(-1) as TemplateManifest['patterns'][number];
const png = placeholderPhoto(40, 40, [0, 0, 0], [255, 255, 255]);

const cardSlide = (items: number, extra: SlideSpec['slots'] = {}): SlideSpec => ({
  pattern: 'cards',
  items,
  slots: { title: 'T', ...Object.fromEntries(Array.from({ length: items }, (_, index) => [`card${index + 1}Text`, `Card ${index + 1}`])), ...extra },
});

function slideRoot(deck: Partial<DeckSpec> & { slides: SlideSpec[] }, assets: Record<string, Uint8Array> = {}): { root: Element; pkg: PptxPackage; pptx: Uint8Array } {
  const result = renderDeck(template, { version: 1, ...deck }, { assets });
  const pkg = PptxPackage.load(result.pptx);
  return { root: rootElement(parseXml(pkg.readText('ppt/slides/slide1.xml'))), pkg, pptx: result.pptx };
}

function shapeById(root: Element, id: string): Element | undefined {
  const props = descendants(root, NS.p, 'cNvPr').find((el) => el.getAttribute('id') === id);
  return props?.parentNode?.parentNode as Element | undefined;
}

/** x of a shape, whatever kind it is. */
function xOf(root: Element, id: string): number | undefined {
  const shape = shapeById(root, id);
  if (!shape) return undefined;
  const xfrm = shape.localName === 'grpSp' ? path(shape, [NS.p, 'grpSpPr'], [NS.a, 'xfrm']) : path(shape, [NS.p, 'spPr'], [NS.a, 'xfrm']);
  return Number(childElement(xfrm as Element, NS.a, 'off')?.getAttribute('x'));
}

const textOf = (root: Element, id: string) =>
  descendants(shapeById(root, id) as Element, NS.a, 't')
    .map((node) => node.textContent ?? '')
    .join('');

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

// --- a row with a different number of items ----------------------------------

test('fewer items: the surplus is removed and the row re-centred at the same spacing', () => {
  const { root } = slideRoot({ slides: [cardSlide(2)] });
  assert.equal(shapeById(root, '30'), undefined, 'the third card is gone');
  // The row was 2000000..7000000 with a pitch of 2000000; two cards sit either side of its centre.
  assert.equal(xOf(root, '10'), 5000000);
  assert.equal(xOf(root, '20'), 3000000);
  assert.equal(xOf(root, '13'), 5100000, "a shape outside the card's group moves with it");
  assert.equal(textOf(root, '11'), 'Card 1');
});

test('more items: the last card is copied, with ids the manifest can predict', () => {
  const { root } = slideRoot({ slides: [cardSlide(4)] });
  assert.deepEqual(
    ['10', '20', '30', '10030'].map((id) => xOf(root, id)),
    [7000000, 5000000, 3000000, 1000000],
    'first card rightmost, the copy at the far end',
  );
  assert.equal(textOf(root, '10031'), 'Card 4', 'the copy has its own text');
  assert.equal(textOf(root, '31'), 'Card 3');
  const ids = descendants(root, NS.p, 'cNvPr').map((el) => el.getAttribute('id'));
  assert.equal(new Set(ids).size, ids.length, 'shape ids stay unique');
});

test('a full row tightens its spacing to stay inside its span', () => {
  const { root } = slideRoot({ slides: [cardSlide(5)] });
  // Span is 5% to 95% of 9144000; five cards of 1000000 leave a pitch of 1807400.
  assert.deepEqual(
    ['10', '20', '30', '10030', '20030'].map((id) => xOf(root, id)),
    [7686800, 5879400, 4072000, 2264600, 457200],
  );
});

test('a copied item takes its own picture, and gives it back', () => {
  const slide = cardSlide(4, { card4Icon: { asset: 'icon' } });
  const { root, pkg, pptx: rendered } = slideRoot({ slides: [slide] }, { icon: png });
  const pic = shapeById(root, '10032') as Element;
  const relId = path(pic, [NS.p, 'blipFill'], [NS.a, 'blip'])?.getAttributeNS(NS.r, 'embed');
  const rel = readRels(pkg, 'ppt/slides/slide1.xml').find((candidate) => candidate.id === relId);
  assert.deepEqual(pkg.read(resolveTarget('ppt/slides/slide1.xml', rel?.target ?? '')), png);
  // The original third card still shows the template's own picture.
  assert.notEqual(path(shapeById(root, '32') as Element, [NS.p, 'blipFill'], [NS.a, 'blip'])?.getAttributeNS(NS.r, 'embed'), relId);

  assert.deepEqual(extractAssets(template, { version: 1, slides: [slide] }, rendered), { icon: png });
});

test('the slots of a pattern follow the number of items', () => {
  const names = (items?: number) => effectiveSlots(cards, items).map((slot) => slot.name);
  assert.deepEqual(names(), ['title', 'card1Text', 'card1Icon', 'card2Text', 'card2Icon', 'card3Text', 'card3Icon']);
  assert.deepEqual(names(1), ['title', 'card1Text', 'card1Icon']);
  assert.deepEqual(names(5).slice(-4), ['card4Text', 'card4Icon', 'card5Text', 'card5Icon']);
  const fifth = effectiveSlots(cards, 5).find((slot) => slot.name === 'card5Icon');
  assert.deepEqual([fifth?.shapeId, fifth?.kind === 'image' ? fifth.copyOf : undefined], ['20032', '32']);
});

test('a deck that reads the other way gets its first item where its readers start', () => {
  assert.deepEqual(Object.fromEntries(slotSources(cards, true, 2)), {
    title: 'title',
    card1Text: 'card2Text',
    card1Icon: 'card2Icon',
    card2Text: 'card1Text',
    card2Icon: 'card1Icon',
  });
  const { root } = slideRoot({ direction: 'ltr', slides: [cardSlide(4)] });
  // Card 1 of the deck lands in the leftmost card, which is the copy.
  assert.equal(textOf(root, '10031'), 'Card 1');
  assert.equal(textOf(root, '11'), 'Card 4');
});

test('item counts are validated', () => {
  const check = (slide: SlideSpec) => validateDeck(cardsManifest, { version: 1, slides: [slide] });
  assert.deepEqual(check(cardSlide(3)), []);
  assert.match(check(cardSlide(6))[0]?.message ?? '', /takes 1 to 5 cards, not 6/);
  assert.match(check({ ...cardSlide(2), slots: { ...cardSlide(3).slots } })[0]?.message ?? '', /no such slot; ignored/);
  assert.match(check({ pattern: 'closing', items: 2, slots: { title: 'x' } })[0]?.message ?? '', /fixed number of items/);
  assert.match(check({ pattern: 'cards', items: 3, slots: { title: 'T', card1Text: 'a', card2Text: 'b' } })[0]?.message ?? '', /required/);
});

// --- larger body text ----------------------------------------------------------

test('a larger text size enlarges body text only, and shrinks its budgets', () => {
  const sizeOf = (root: Element, id: string) => descendants(shapeById(root, id) as Element, NS.a, 'rPr')[0]?.getAttribute('sz');
  const normal = slideRoot({ slides: [cardSlide(3)] }).root;
  const large = slideRoot({ textScale: 1.25, slides: [cardSlide(3)] }).root;
  assert.equal(sizeOf(normal, '11'), '900');
  assert.equal(sizeOf(large, '11'), '1150', '9 pt becomes 11.5 pt, rounded to half a point');
  assert.equal(sizeOf(large, '2'), '2600', 'the 26 pt title is not body text');

  const text = cards.slots.find((slot) => slot.name === 'card1Text');
  assert.ok(text?.kind === 'text');
  assert.deepEqual(textBudget(text, 1), { maxChars: 20 });
  assert.deepEqual(textBudget(text, 1.25), { maxChars: 12 });
  assert.deepEqual(textBudget(text, 1.5), { maxChars: 8 });

  const long = cardSlide(3, { card1Text: 'Fifteen letters' });
  assert.deepEqual(validateDeck(cardsManifest, { version: 1, slides: [long] }), []);
  assert.match(validateDeck(cardsManifest, { version: 1, textScale: 1.25, slides: [long] })[0]?.message ?? '', /text is 15 characters; the slot fits about 12/);
});

// --- derived patterns ------------------------------------------------------------

test('a derived pattern can move a text box, set its size and space its paragraphs', () => {
  const derived: TemplateManifest = {
    ...fixtureManifest,
    patterns: [
      {
        id: 'text-only',
        sourceSlide: 1,
        description: 'Text only.',
        slots: [
          { name: 'title', kind: 'text', shapeId: '2', styleRun: 'last' },
          { name: 'body', kind: 'text', shapeId: '3', fontSize: 16, scales: true, paragraphGap: 12, box: { x: 0.1, y: 0.2, w: 0.5, h: 0.3 } },
          { name: 'picture', kind: 'image', imageKind: 'graphic', shapeId: '4', empty: 'drop' },
          { name: 'note', kind: 'text', shapeId: '6', empty: 'drop' },
        ],
      },
    ],
  };
  const render = (textScale?: number) => {
    const deck: DeckSpec = { version: 1, slides: [{ pattern: 'text-only', slots: { title: 'T', body: ['One', 'Two'] } }] };
    if (textScale) deck.textScale = textScale;
    return rootElement(parseXml(PptxPackage.load(renderDeck(loadTemplate(pptx, derived), deck).pptx).readText('ppt/slides/slide1.xml')));
  };

  const root = render();
  const body = shapeById(root, '3') as Element;
  const xfrm = path(body, [NS.p, 'spPr'], [NS.a, 'xfrm']) as Element;
  assert.deepEqual(
    [childElement(xfrm, NS.a, 'off')?.getAttribute('x'), childElement(xfrm, NS.a, 'off')?.getAttribute('y'), childElement(xfrm, NS.a, 'ext')?.getAttribute('cx')],
    ['914400', '1028700', '4572000'],
  );
  const paragraphs = childElements(childElement(body, NS.p, 'txBody') as Element, NS.a, 'p');
  assert.deepEqual(
    paragraphs.map((p) => descendants(p, NS.a, 'rPr')[0]?.getAttribute('sz')),
    ['1600', '1600'],
    'both paragraphs take the set size, whatever the template had',
  );
  assert.equal(path(paragraphs[0] as Element, [NS.a, 'pPr'], [NS.a, 'spcAft'], [NS.a, 'spcPts'])?.getAttribute('val'), '1200');
  assert.equal(shapeById(root, '4'), undefined, 'the picture the pattern does not use is dropped');
  assert.equal(shapeById(root, '6'), undefined);

  assert.equal(descendants(shapeById(render(1.25), '3') as Element, NS.a, 'rPr')[0]?.getAttribute('sz'), '2000', 'the set size scales with the deck');
});

// --- planning and editing ----------------------------------------------------------

test('the planner chooses the number of items and writes exactly that many', async () => {
  const { llm, requests } = scripted(
    { title: 'T', slides: [{ pattern: 'cards', purpose: 'Two things', items: 2 }] },
    { s1: { title: 'Two', card1Text: 'One', card2Text: 'Two' } },
  );
  const { deck } = await planDeck(template, { prompt: 'Two things', language: 'en', imageDensity: 'none' }, { llm });

  assert.match(requests[0]?.prompt ?? '', /- cards: A row of cards\.\n {2}slots: title \(text\); cardNText \(at most 20 characters\)\n {2}items: 1 to 5 cards side by side in one row, numbered/);
  const outline = (requests[0]?.schema as { properties: { slides: { items: { properties: { items: { minimum: number; maximum: number } } } } } }).properties.slides.items
    .properties.items;
  assert.deepEqual([outline.minimum, outline.maximum], [1, 5]);
  const written = (requests[1]?.schema as { properties: { s1: { properties: Record<string, unknown>; required: string[] } } }).properties.s1;
  assert.deepEqual(Object.keys(written.properties), ['title', 'card1Text', 'card2Text']);
  assert.deepEqual(deck.slides[0], { pattern: 'cards', items: 2, slots: { title: 'Two', card1Text: 'One', card2Text: 'Two' } });

  const tooMany = scripted({ title: 'T', slides: [{ pattern: 'cards', purpose: 'p', items: 9 }] }, { title: 'T', slides: [{ pattern: 'cards', purpose: 'p', items: 5 }] }, { s1: {} });
  await planDeck(template, { prompt: 'x', language: 'en', imageDensity: 'none' }, { llm: tooMany.llm, maxRepairs: 0 }).catch(() => undefined);
  assert.match(tooMany.requests[1]?.prompt ?? '', /takes 1 to 5 cards, not 9/);
});

test('the planner is told which patterns set their items in one row and which break them over lines', async () => {
  // Both hold cards. Only one of them answers a wish for three columns side by side.
  const both: TemplateManifest = { ...cardsManifest, patterns: [...cardsManifest.patterns, ...gridManifest.patterns.filter((pattern) => pattern.id === 'grid')] };
  const { llm, requests } = scripted({ title: 'T', slides: [{ pattern: 'cards', purpose: 'Three things', items: 3 }] }, { s1: { title: 'T', card1Text: 'One', card2Text: 'Two', card3Text: 'Three' } });
  await planDeck(loadTemplate(pptx, both), { prompt: 'Three columns side by side', language: 'en', imageDensity: 'none' }, { llm });
  const prompt = requests[0]?.prompt ?? '';
  assert.match(prompt, /\n {2}items: 1 to 5 cards side by side in one row, numbered N = 1, 2, \.\.\./);
  assert.match(prompt, /\n {2}items: 1 to 4 cards in lines of up to 2, numbered N = 1, 2, \.\.\./);
  // The rule asks for a pattern that takes the number and sets it on one line: a row, or a grid
  // whose line holds it. It says what to do when the list has none.
  assert.match(prompt, /When the brief asks for a number of columns, or for items side by side, use a pattern that takes that number of items and whose items are "side by side in one row" or "in lines of" that number or more: either sets them on one line\. .* When the list has no such pattern, use the one that comes closest\./);
  // The number in the schema covers every row on offer.
  const allowed = (requests[0]?.schema as { properties: { slides: { items: { properties: { items: { minimum: number; maximum: number } } } } } }).properties.slides.items.properties.items;
  assert.deepEqual([allowed.minimum, allowed.maximum], [1, 5]);
});

test('a row that takes one number of items says so, and a grid that never breaks a line reads as a row', async () => {
  const cards = cardsManifest.patterns.find((pattern) => pattern.id === 'cards') as TemplateManifest['patterns'][number];
  const grid = gridManifest.patterns.find((pattern) => pattern.id === 'grid') as TemplateManifest['patterns'][number];
  const fixed: TemplateManifest = {
    ...cardsManifest,
    patterns: [
      ...fixtureManifest.patterns,
      // Three cards are drawn; this row always takes two.
      { ...cards, repeat: { ...(cards.repeat as NonNullable<typeof cards.repeat>), min: 2, max: 2 } },
      // Two of the grid's four cards are one line of two.
      { ...grid, repeat: { ...(grid.repeat as NonNullable<typeof grid.repeat>), min: 1, max: 2 } },
    ],
  };
  // The outline names no number for the row: its description gives none, and the line in the list does.
  const { llm, requests } = scripted({ title: 'T', slides: [{ pattern: 'cards', purpose: 'Two things' }] }, { s1: { title: 'T', card1Text: 'One', card2Text: 'Two' } });
  const { deck } = await planDeck(loadTemplate(pptx, fixed), { prompt: 'Two things', language: 'en', imageDensity: 'none' }, { llm });
  const prompt = requests[0]?.prompt ?? '';
  assert.match(prompt, /\n {2}items: always 2 cards side by side in one row, numbered N = 1, \.\.\., 2/);
  assert.match(prompt, /\n {2}items: 1 to 2 cards side by side in one row, numbered N = 1, 2, \.\.\./);
  // A grid that always takes its four is two lines of two, and reads so.
  const four: TemplateManifest = { ...fixed, patterns: [...fixed.patterns.slice(0, -1), { ...grid, repeat: { ...(grid.repeat as NonNullable<typeof grid.repeat>), min: 4, max: 4 } }] };
  const asFour = scripted({ title: 'T', slides: [{ pattern: 'cards', purpose: 'p' }] }, { s1: { title: 'T', card1Text: 'One', card2Text: 'Two' } });
  await planDeck(loadTemplate(pptx, four), { prompt: 'Two things', language: 'en', imageDensity: 'none' }, { llm: asFour.llm });
  assert.match(asFour.requests[0]?.prompt ?? '', /\n {2}items: always 4 cards in lines of 2, numbered N = 1, \.\.\., 4/);
  // The row that takes two is in the list beside rows that vary, so the schema lets a model name two.
  const allowed = (requests[0]?.schema as { properties: { slides: { items: { properties: { items: { minimum: number; maximum: number } } } } } }).properties.slides.items.properties.items;
  assert.deepEqual([allowed.minimum, allowed.maximum], [1, 2]);
  // With no number given, the row is planned with the number it takes, not the three it draws.
  assert.deepEqual(Object.keys(deck.slides[0]?.slots ?? {}), ['title', 'card1Text', 'card2Text']);

  const wrong = scripted({ title: 'T', slides: [{ pattern: 'cards', purpose: 'p', items: 3 }] }, { title: 'T', slides: [{ pattern: 'cards', purpose: 'p' }] }, { s1: {} });
  await planDeck(loadTemplate(pptx, fixed), { prompt: 'x', language: 'en', imageDensity: 'none' }, { llm: wrong.llm, maxRepairs: 0 }).catch(() => undefined);
  assert.match(wrong.requests[1]?.prompt ?? '', /takes exactly 2 cards, not 3/);
});

test('the number a row always takes is allowed by the schema of a plan and of an edit', async () => {
  // A row that always takes its three cards, beside a grid that takes one or two.
  const cards = cardsManifest.patterns.find((pattern) => pattern.id === 'cards') as TemplateManifest['patterns'][number];
  const grid = gridManifest.patterns.find((pattern) => pattern.id === 'grid') as TemplateManifest['patterns'][number];
  const mixed: TemplateManifest = {
    ...cardsManifest,
    patterns: [...fixtureManifest.patterns, { ...cards, repeat: { ...(cards.repeat as NonNullable<typeof cards.repeat>), min: 3, max: 3 } }, { ...grid, repeat: { ...(grid.repeat as NonNullable<typeof grid.repeat>), min: 1, max: 2 } }],
  };
  const outline = scripted({ title: 'T', slides: [{ pattern: 'cards', purpose: 'Three things', items: 3 }] }, { s1: { title: 'T', card1Text: 'One', card2Text: 'Two', card3Text: 'Three' } });
  const { deck } = await planDeck(loadTemplate(pptx, mixed), { prompt: 'Three things', language: 'en', imageDensity: 'none' }, { llm: outline.llm });
  assert.match(outline.requests[0]?.prompt ?? '', /\n {2}items: always 3 cards side by side in one row, numbered N = 1, \.\.\., 3/);
  const planned = (outline.requests[0]?.schema as { properties: { slides: { items: { properties: { items: { minimum: number; maximum: number } } } } } }).properties.slides.items.properties.items;
  assert.deepEqual([planned.minimum, planned.maximum], [1, 3]);

  const change = scripted(editPlan([{ from: 1, pattern: 'cards', change: '', images: false }]));
  await planEdit(loadTemplate(pptx, mixed), deck, 'Nothing to change', { llm: change.llm });
  const edited = (change.requests[0]?.schema as { properties: { slides: { items: { properties: { items: { maximum: number } } } } } }).properties.slides.items.properties.items;
  assert.equal(edited.maximum, 3);
});

test('a request to change a plan or a deck carries the rule about columns too', async () => {
  const outline = scripted({ title: 'T', slides: [{ pattern: 'cards', purpose: 'Three things', items: 3 }] });
  const plan = await planOutline(template, { prompt: 'Three things', language: 'en', imageDensity: 'none' }, { llm: outline.llm });
  const change = scripted({ summary: 'Done.', title: 'T', slides: [{ pattern: 'cards', purpose: 'Five things', items: 5 }] });
  await changePlan(template, plan, 'Show slide 1 as five columns', { llm: change.llm }).catch(() => undefined);
  assert.match(change.requests[0]?.prompt ?? '', /- When the request asks for a number of columns, or for items side by side, use a pattern that takes that number of items and whose items are "side by side in one row"/);
});

test('an edit leaves alone a slide that states no number of items, whatever its row takes', async () => {
  // The row draws three cards and takes exactly two. This slide was stored without `items`, so it shows three.
  const cards = cardsManifest.patterns.find((pattern) => pattern.id === 'cards') as TemplateManifest['patterns'][number];
  const exact: TemplateManifest = { ...cardsManifest, patterns: [...fixtureManifest.patterns, { ...cards, repeat: { ...(cards.repeat as NonNullable<typeof cards.repeat>), min: 2, max: 2 } }] };
  const stored: DeckSpec = { version: 1, language: 'en', direction: 'ltr', slides: [{ pattern: 'cards', slots: { title: 'T', card1Text: 'A', card2Text: 'B', card3Text: 'C' } }, cardSlide(2)] };
  const { llm } = scripted(
    editPlan([
      { from: 1, pattern: 'cards', change: '', images: false },
      { from: 2, pattern: 'cards', change: 'New title', images: false, items: 2 },
    ]),
    { s2: { title: 'New', card1Text: 'A', card2Text: 'B' } },
  );
  const result = await planEdit(loadTemplate(pptx, exact), stored, 'Give slide 2 a new title', { llm });
  assert.deepEqual(result.rewritten, [2], 'the first slide is copied, not rewritten to the number its row takes');
  assert.deepEqual(result.deck.slides[0], stored.slides[0]);
});

test('a larger text size reaches the model as smaller limits', async () => {
  const { llm, requests } = scripted({ title: 'T', slides: [{ pattern: 'cards', purpose: 'p', items: 1 }] }, { s1: { title: 'T', card1Text: 'Short' } });
  const { deck } = await planDeck(template, { prompt: 'x', language: 'en', imageDensity: 'none', textSize: 'large' }, { llm });
  assert.match(requests[0]?.prompt ?? '', /cardNText \(at most 12 characters\)/);
  assert.match(JSON.stringify(requests[1]?.schema), /at most 12 characters/);
  assert.equal(deck.textScale, 1.25);
});

const three: DeckSpec = { version: 1, language: 'en', direction: 'ltr', slides: [cardSlide(3), cardSlide(3, { card1Text: 'Fifteen letters' })] };
const editPlan = (slides: Array<Record<string, unknown>>, extra: Record<string, unknown> = {}) => ({ summary: 'Done.', language: 'en', title: '', text_size: 'normal', slides, ...extra });

test('an edit can change the number of items on a slide', async () => {
  const { llm, requests } = scripted(
    editPlan([
      { from: 1, pattern: 'cards', change: '', images: false, items: 2 },
      { from: 2, pattern: 'cards', change: '', images: false },
    ]),
    { s1: { title: 'T', card1Text: 'A', card2Text: 'B' } },
  );
  const result = await planEdit(template, three, 'Make slide 1 two cards', { llm });

  assert.match(requests[0]?.prompt ?? '', /1\. \[cards, 3 items\]/);
  // The list of patterns is for a slide that changes its layout as well, and comes with the rule about columns.
  assert.match(requests[0]?.prompt ?? '', /When the request asks for a number of columns, or for items side by side/);
  assert.match(requests[0]?.prompt ?? '', /Patterns, for new slides and for a slide that changes its layout/);
  assert.deepEqual(result.rewritten, [1], 'a new count rewrites the slide even with no other change');
  assert.match(requests[1]?.prompt ?? '', /Change: Use 2 cards\./);
  assert.equal(result.deck.slides[0]?.items, 2);
  assert.deepEqual(Object.keys(result.deck.slides[0]?.slots ?? {}), ['title', 'card1Text', 'card2Text']);
  assert.deepEqual(result.deck.slides[1], three.slides[1], 'a slide the plan leaves out of "items" keeps its own count');
});

test('an edit to a larger text size rewrites only the slides that no longer fit', async () => {
  const keep = [
    { from: 1, pattern: 'cards', change: '', images: false },
    { from: 2, pattern: 'cards', change: '', images: false },
  ];
  const { llm, requests } = scripted(editPlan(keep, { text_size: 'large' }), { s2: { title: 'T', card1Text: 'Short', card2Text: 'Card 2', card3Text: 'Card 3' } });
  const result = await planEdit(template, three, 'Make the text bigger', { llm });

  assert.match(requests[0]?.prompt ?? '', /body text size "normal"/);
  assert.equal(result.deck.textScale, 1.25);
  assert.deepEqual(result.rewritten, [2], 'slide 1 still fits; slide 2 has a 15-character card and 12 now fit');
  assert.match(requests[1]?.prompt ?? '', /Shorten the text to fit the new text size/);
  assert.match(JSON.stringify(requests[1]?.schema), /at most 12 characters/);
  assert.equal(result.deck.slides[1]?.slots.card1Text, 'Short');
  assert.deepEqual(result.deck.slides[0], three.slides[0]);

  // Back to normal: nothing needs rewriting, and the setting is cleared.
  const back = scripted(editPlan(keep, { text_size: 'normal' }));
  const smaller = await planEdit(template, result.deck, 'Normal text size again', { llm: back.llm });
  assert.equal(smaller.deck.textScale, undefined);
  assert.equal(back.requests.length, 1);
});

test('a text size that is not one of the list keeps the size there is, also one every object has', async () => {
  // "constructor" and "__proto__" are in every object: they were once read as sizes, and the deck could not be rendered.
  const larger: DeckSpec = { ...three, slides: [cardSlide(3)], textScale: 1.25 };
  for (const size of ['constructor', 'toString', '__proto__', 'hasOwnProperty', 'huge']) {
    const { llm, requests } = scripted(editPlan([{ from: 1, pattern: 'cards', change: '', images: false }], { text_size: size }));
    const result = await planEdit(template, larger, 'Make the text larger', { llm });
    assert.equal(result.deck.textScale, 1.25, size);
    assert.deepEqual(result.rewritten, [], size);
    assert.equal(requests.length, 1, size);
  }
  // A brief's size is read the same way: the deck is written in the usual size.
  for (const size of ['__proto__', 'constructor', 'huge']) {
    const { llm } = scripted({ title: 'T', slides: [{ pattern: 'cards', purpose: 'p', items: 1 }] }, { s1: { title: 'T', card1Text: 'Short' } });
    const { deck } = await planDeck(template, { prompt: 'x', language: 'en', imageDensity: 'none', textSize: size as 'large' }, { llm });
    assert.equal(deck.textScale, undefined, size);
    assert.equal(renderDeck(template, deck).warnings.filter((issue) => issue.severity === 'error').length, 0, size);
  }
});

test('the plan tool passes the text size on', async () => {
  const model = scripted({ title: 'T', slides: [{ pattern: 'cards', purpose: 'p', items: 2 }] });
  const tools = presentationTools({ template, llm: model.llm, store: memoryStore() });
  const planned = await tools.call(TOOL_NAMES.plan, { topic: 'x', text_size: 'x-large' });
  assert.equal(planned.plan?.brief.textSize, 'x-large');
  assert.equal(planned.plan?.slides[0]?.items, 2);
  assert.match(model.requests[0]?.prompt ?? '', /cardNText \(at most 8 characters\)/);
});

// --- grids ---------------------------------------------------------------------------

test('a grid spreads its items evenly over as few rows as they need', () => {
  const grid = { item: 'card', slots: ['c{n}'], shapes: [['1'], ['2'], ['3'], ['4'], ['5'], ['6']], columns: 3 };
  assert.deepEqual([1, 2, 3, 4, 5, 6].map((count) => rowLayout(grid, count)), [[1], [2], [3], [2, 2], [3, 2], [3, 3]]);
  assert.deepEqual(rowLayout({ ...grid, columns: 2 }, 3), [2, 1]);
  const { columns: _columns, ...row } = grid;
  assert.deepEqual(rowLayout(row, 5), [5], 'without columns everything is one row');
});

test('a grid with fewer items re-forms around the centre of the area it had', () => {
  const gridTemplate = loadTemplate(pptx, gridManifest);
  const render = (items: number, direction?: 'ltr') => {
    const slots = Object.fromEntries(Array.from({ length: items }, (_, index) => [`card${index + 1}Text`, `Card ${index + 1}`]));
    const deck: DeckSpec = { version: 1, slides: [{ pattern: 'grid', items, slots: { title: 'T', ...slots } }] };
    if (direction) deck.direction = direction;
    return rootElement(parseXml(PptxPackage.load(renderDeck(gridTemplate, deck).pptx).readText('ppt/slides/slide1.xml')));
  };
  const place = (root: Element, id: string) => {
    const shape = shapeById(root, id);
    if (!shape) return undefined;
    const off = childElement(path(shape, [NS.p, 'grpSpPr'], [NS.a, 'xfrm']) as Element, NS.a, 'off');
    return [Number(off?.getAttribute('x')), Number(off?.getAttribute('y'))];
  };

  // The grid is 4000000..7000000 wide and 1000000..4500000 tall, with a pitch of 2000000 both ways.
  const full = render(4);
  assert.deepEqual(['10', '20', '30', '40'].map((id) => place(full, id)), [[6000000, 1000000], [4000000, 1000000], [6000000, 3000000], [4000000, 3000000]]);

  const three = render(3);
  assert.deepEqual(['10', '20', '30'].map((id) => place(three, id)), [[6000000, 1000000], [4000000, 1000000], [5000000, 3000000]], 'the odd card is centred under the row');
  assert.equal(place(three, '40'), undefined);

  const two = render(2);
  assert.deepEqual(['10', '20'].map((id) => place(two, id)), [[6000000, 2000000], [4000000, 2000000]], 'a single row moves to the middle');

  // Read left to right, each row is mirrored on its own.
  assert.deepEqual(Object.fromEntries(slotSources(gridManifest.patterns.at(-1) as never, true, 3)), {
    title: 'title',
    card1Text: 'card2Text',
    card2Text: 'card1Text',
    card3Text: 'card3Text',
  });
  const english = render(3, 'ltr');
  assert.deepEqual(['21', '11', '31'].map((id) => textOf(english, id)), ['Card 1', 'Card 2', 'Card 3'], 'first card top-left, third below');
});

// --- rows that widen their items ----------------------------------------------------------

const wideManifest: TemplateManifest = {
  ...cardsManifest,
  patterns: cardsManifest.patterns.map((pattern) => (pattern.repeat ? { ...pattern, repeat: { ...pattern.repeat, widen: 2 } } : pattern)),
};
const wideTemplate = loadTemplate(pptx, wideManifest);
const wideCards = wideManifest.patterns.at(-1) as TemplateManifest['patterns'][number];

function slideWith(used: typeof template, deck: Partial<DeckSpec> & { slides: SlideSpec[] }): { root: Element; pptx: Uint8Array } {
  const rendered = renderDeck(used, { version: 1, ...deck });
  return { root: rootElement(parseXml(PptxPackage.load(rendered.pptx).readText('ppt/slides/slide1.xml'))), pptx: rendered.pptx };
}

/** x and width of a shape, in the units of whatever it sits in; for a group also the width of its own units. */
function across(root: Element, id: string): number[] | undefined {
  const shape = shapeById(root, id);
  if (!shape) return undefined;
  const xfrm = (shape.localName === 'grpSp' ? path(shape, [NS.p, 'grpSpPr'], [NS.a, 'xfrm']) : path(shape, [NS.p, 'spPr'], [NS.a, 'xfrm'])) as Element;
  const inner = childElement(xfrm, NS.a, 'chExt');
  return [childElement(xfrm, NS.a, 'off')?.getAttribute('x'), childElement(xfrm, NS.a, 'ext')?.getAttribute('cx'), ...(inner ? [inner.getAttribute('cx')] : [])].map(Number);
}

test('with widen, a row of fewer items grows to fill its line, and pictures keep their size', () => {
  const { root } = slideWith(wideTemplate, { slides: [cardSlide(2)] });
  // Three cards of 1000000 were drawn 1000000 apart. Two may each be twice as wide; the gap stays.
  assert.deepEqual(across(root, '10'), [5000000, 2000000, 2000000], 'the first card, at the right, is twice as wide and so are its own units');
  assert.deepEqual(across(root, '20'), [2000000, 2000000, 2000000]);
  assert.equal(shapeById(root, '30'), undefined);
  // Inside a card, the text box grows with it. The picture is not stretched: it stays 400000 wide, in the middle.
  assert.deepEqual(across(root, '11'), [6000000, 2000000]);
  assert.deepEqual(across(root, '12'), [6800000, 400000]);
  assert.equal(path(shapeById(root, '12') as Element, [NS.p, 'spPr'], [NS.a, 'xfrm'], [NS.a, 'ext'])?.getAttribute('cy'), '400000');
  // A text box that belongs to the card but sits outside its group grows and moves with it.
  assert.deepEqual(across(root, '13'), [5100000, 1800000]);
  assert.equal(textOf(root, '11'), 'Card 1');

  // One card: as wide as the limit allows, in the middle of the row.
  assert.deepEqual(across(slideWith(wideTemplate, { slides: [cardSlide(1)] }).root, '10'), [3500000, 2000000, 2000000]);

  // A deck that reads the other way gets the same shapes, with its first card at the left.
  const english = slideWith(wideTemplate, { direction: 'ltr', slides: [cardSlide(2)] }).root;
  assert.deepEqual([across(english, '10'), across(english, '20')], [across(root, '10'), across(root, '20')]);
  assert.deepEqual([textOf(english, '21'), textOf(english, '11')], ['Card 1', 'Card 2']);
});

test('the number of items alone decides: at the drawn number and above, and without the setting, nothing changes', () => {
  for (const items of [3, 4, 5]) {
    assert.deepEqual(slideWith(wideTemplate, { slides: [cardSlide(items)] }).pptx, slideWith(template, { slides: [cardSlide(items)] }).pptx, `${items} cards`);
  }
  // Long text does not widen anything, and short text does not narrow it.
  const long = slideWith(wideTemplate, { slides: [cardSlide(2, { card1Text: 'A much longer text than the others have' })] }).root;
  assert.deepEqual(across(long, '10'), [5000000, 2000000, 2000000]);
  // A factor of 1 is the same as no setting.
  const one = loadTemplate(pptx, { ...wideManifest, patterns: wideManifest.patterns.map((pattern) => (pattern.repeat ? { ...pattern, repeat: { ...pattern.repeat, widen: 1 } } : pattern)) });
  assert.deepEqual(slideWith(one, { slides: [cardSlide(2)] }).pptx, slideWith(template, { slides: [cardSlide(2)] }).pptx);
});

test('budgets of widened items follow their width, in validation and in what the model is shown', async () => {
  const budget = (items?: number) => {
    const slot = effectiveSlots(wideCards, items).find((candidate) => candidate.name === 'card1Text');
    return slot?.kind === 'text' ? slot.maxChars : undefined;
  };
  // Three drawn: two are at least 1.5 times as wide, one is twice (the limit), three or more are as drawn.
  assert.deepEqual([budget(1), budget(2), budget(3), budget(5), budget()], [40, 30, 20, 20, 20]);
  assert.equal(effectiveSlots(wideCards, 3)[1], wideCards.slots[1], 'a row that does not widen hands out the slots of the manifest themselves');
  assert.deepEqual(wideCards.slots.find((slot) => slot.name === 'card1Text'), cards.slots.find((slot) => slot.name === 'card1Text'), 'the manifest is not changed');

  const slide = cardSlide(2, { card1Text: 'Twenty-eight characters long' });
  assert.deepEqual(validateDeck(wideManifest, { version: 1, slides: [slide] }), []);
  assert.match(validateDeck(cardsManifest, { version: 1, slides: [slide] })[0]?.message ?? '', /text is 28 characters; the slot fits about 20/);
  assert.match(validateDeck(wideManifest, { version: 1, slides: [cardSlide(3, { card1Text: 'Twenty-eight characters long' })] })[0]?.message ?? '', /fits about 20/);
  // Larger text still shrinks the budget, on top.
  assert.match(validateDeck(wideManifest, { version: 1, textScale: 1.25, slides: [slide] })[0]?.message ?? '', /text is 28 characters; the slot fits about 19/);

  const { llm, requests } = scripted({ title: 'T', slides: [{ pattern: 'cards', purpose: 'Two things', items: 2 }] }, { s1: { title: 'Two', card1Text: 'One', card2Text: 'Two' } });
  await planDeck(wideTemplate, { prompt: 'Two things', language: 'en', imageDensity: 'none' }, { llm });
  assert.match(requests[0]?.prompt ?? '', /cardNText \(at most 20 characters\)\n {2}items: 1 to 5 cards side by side in one row, numbered N = 1, 2, \.\.\. With fewer on a line than 3, each card is wider and holds more text\./);
  assert.match(JSON.stringify(requests[1]?.schema), /at most 30 characters/);
  // The plain row says nothing of the kind.
  const plain = scripted({ title: 'T', slides: [{ pattern: 'cards', purpose: 'p', items: 2 }] }, { s1: { title: 'Two', card1Text: 'One', card2Text: 'Two' } });
  await planDeck(template, { prompt: 'x', language: 'en', imageDensity: 'none' }, { llm: plain.llm });
  assert.doesNotMatch(plain.requests[0]?.prompt ?? '', /is wider and holds more text/);
});

test('a grid widens when its fullest line is shorter than drawn, and only then', () => {
  const widen = (manifest: TemplateManifest): TemplateManifest => ({ ...manifest, patterns: manifest.patterns.map((pattern) => (pattern.repeat ? { ...pattern, repeat: { ...pattern.repeat, widen: 1.5 } } : pattern)) });
  const plainGrid = loadTemplate(pptx, gridManifest);
  const wideGrid = loadTemplate(pptx, widen(gridManifest));
  const grid = (items: number): SlideSpec => ({ pattern: 'grid', items, slots: { title: 'T', ...Object.fromEntries(Array.from({ length: items }, (_, index) => [`card${index + 1}Text`, `Card ${index + 1}`])) } });
  // Two to a line were drawn. One card is a shorter line: it grows by half and sits in the middle of the grid's area.
  assert.deepEqual(across(slideWith(wideGrid, { slides: [grid(1)] }).root, '10'), [4750000, 1500000, 1500000]);
  // Three make a line of two and a line of one: the fullest line is as drawn, so every card keeps its width.
  for (const items of [2, 3, 4]) assert.deepEqual(slideWith(wideGrid, { slides: [grid(items)] }).pptx, slideWith(plainGrid, { slides: [grid(items)] }).pptx, `${items} cards`);
});

test('widening a card keeps round corners, circles and pictures as drawn, through a scaled group', () => {
  const sp = (id: number, x: number, y: number, cx: number, cy: number, geometry: string) =>
    `<p:sp><p:nvSpPr><p:cNvPr id="${id}" name="s"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x="${x}" y="${y}"/><a:ext cx="${cx}" cy="${cy}"/></a:xfrm>${geometry}</p:spPr></p:sp>`;
  // A card whose own units are half the slide's: 1000000 inside is 2000000 on the slide.
  const card = (id: number, x: number) =>
    `<p:grpSp><p:nvGrpSpPr><p:cNvPr id="${id}" name="Card"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>` +
    `<p:grpSpPr><a:xfrm><a:off x="${x}" y="0"/><a:ext cx="2000000" cy="4000000"/><a:chOff x="0" y="0"/><a:chExt cx="1000000" cy="2000000"/></a:xfrm></p:grpSpPr>` +
    sp(id + 1, 0, 0, 1000000, 2000000, '<a:prstGeom prst="roundRect"><a:avLst><a:gd name="adj" fmla="val 10000"/></a:avLst></a:prstGeom>') +
    sp(id + 2, 100000, 100000, 200000, 200000, '<a:prstGeom prst="ellipse"><a:avLst/></a:prstGeom>') +
    `<p:pic><p:nvPicPr><p:cNvPr id="${id + 3}" name="Icon"/><p:cNvPicPr/><p:nvPr/></p:nvPicPr><p:blipFill/><p:spPr><a:xfrm><a:off x="700000" y="100000"/><a:ext cx="200000" cy="200000"/></a:xfrm></p:spPr></p:pic>` +
    `<p:cxnSp><p:nvCxnSpPr><p:cNvPr id="${id + 4}" name="Rule"/><p:cNvCxnSpPr/><p:nvPr/></p:nvCxnSpPr><p:spPr><a:xfrm><a:off x="100000" y="400000"/><a:ext cx="800000" cy="0"/></a:xfrm><a:prstGeom prst="line"><a:avLst/></a:prstGeom></p:spPr></p:cxnSp>` +
    sp(id + 5, 100000, 500000, 300000, 200000, '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom>') +
    sp(id + 6, 600000, 1500000, 300000, 200000, '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom>') +
    `</p:grpSp>`;
  const root = rootElement(
    parseXml(
      `<p:sld xmlns:a="${NS.a}" xmlns:p="${NS.p}"><p:cSld><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/>${card(10, 1000000)}${card(20, 4000000)}</p:spTree></p:cSld></p:sld>`,
    ),
  );
  // One card of two: the limit of 1.5 makes it 3000000 wide on the slide, which is 500000 more in its own units.
  applyRepeat(indexShapes(root), { item: 'card', slots: ['card{n}Text'], shapes: [['10'], ['20']], widen: 1.5 }, 1, 12192000, new Set(['15']));
  assert.deepEqual(across(root, '10'), [2000000, 3000000, 1500000], 'the card grows and keeps the scale of its own units');
  assert.deepEqual(across(root, '11'), [0, 1500000], 'the background fills the card');
  // The corner was a tenth of the shorter side, the width; the same corner is a fifteenth of the new width.
  assert.equal(path(shapeById(root, '11') as Element, [NS.p, 'spPr'], [NS.a, 'prstGeom'], [NS.a, 'avLst'], [NS.a, 'gd'])?.getAttribute('fmla'), 'val 6667');
  assert.deepEqual(across(root, '12'), [100000, 200000], 'a circle near the left edge stays a circle, near the left edge');
  assert.deepEqual(across(root, '13'), [1200000, 200000], 'a picture near the right edge keeps its size and its distance to that edge');
  assert.deepEqual(across(root, '14'), [100000, 1300000], 'a rule that spans the card grows with it');
  assert.deepEqual(across(root, '15'), [100000, 800000], 'a text the deck writes grows even though it is narrow');
  assert.deepEqual(across(root, '16'), [1100000, 300000], 'a narrow box that is not such a text keeps its size, by the nearer edge');
  assert.equal(shapeById(root, '20'), undefined);
});

test('the studio sets widen, a model can ask for it, and the checker knows it', async () => {
  const changed = changeManifest(template, [
    { op: 'setRepeat', pattern: 'cards', widen: 1.5 },
    { op: 'setRepeat', pattern: 'cards', widen: 9 },
  ]);
  assert.match(changed.applied[0]?.message ?? '', /^"cards": the row of cards takes 1 to 5\. With fewer cards on a line than drawn, each grows up to 1\.5 times as wide to fill the row, and holds more text\.$/);
  assert.match(changed.rejected[0]?.reason ?? '', /^widen is a factor between 1 and 4/);
  assert.equal(changed.manifest.patterns.at(-1)?.repeat?.widen, 1.5);
  const made = { pkg: template.pkg, manifest: changed.manifest };
  assert.match(describeManifest(made), /row of 3 cards \(card\{n\}Text, card\{n\}Icon\); takes 1 to 5; fewer on a line grow up to 1\.5 times as wide/);
  const cleared = changeManifest(made, [{ op: 'setRepeat', pattern: 'cards', widen: null }]);
  assert.match(cleared.applied[0]?.message ?? '', /The cards keep their drawn width whatever their number\.$/);
  assert.deepEqual(cleared.manifest, cardsManifest);

  const { llm, requests } = scripted({ changes: [{ op: 'setRepeat', pattern: 'cards', shape: '', slot: '', values: [{ property: 'widen', value: '1.6' }] }], note: '' });
  const edited = await editManifest(template, 'With two cards, let them fill the row.', { llm });
  assert.match(requests[0]?.prompt ?? '', /- setRepeat \(pattern\): .* Values: min, max, and widen \(a factor such as "1\.5"/);
  assert.deepEqual(edited.applied.map((entry) => entry.change), [{ op: 'setRepeat', pattern: 'cards', widen: 1.6 }]);

  const broken = { ...cardsManifest, patterns: cardsManifest.patterns.map((pattern) => (pattern.repeat ? { ...pattern, repeat: { ...pattern.repeat, widen: 0.5 } } : pattern)) };
  assert.deepEqual(checkManifest(loadTemplate(pptx, broken)).filter((issue) => issue.severity === 'error').map((issue) => issue.message), ['repeat.widen is a factor of 1 or more, such as 1.5']);
  // The sample deck shows the widened row at its own budgets when asked for the fewest items.
  const sample = sampleDeck(made, { items: 'min' }).deck.slides.at(-1);
  assert.deepEqual([sample?.items, String(sample?.slots.card1Text).length > 20], [1, true]);
});
