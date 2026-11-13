// Layout freedom that stays inside the template: a row with fewer or more items, larger body
// text, and patterns derived from a slide by moving and resizing a text box.
import assert from 'node:assert/strict';
import { test } from 'node:test';
import { placeholderPhoto } from '../examples/placeholder.ts';
import { PptxPackage, TOOL_NAMES, extractAssets, loadTemplate, memoryStore, planDeck, planEdit, presentationTools, renderDeck, validateDeck } from '../src/index.ts';
import type { DeckSpec, JsonRequest, LlmProvider, SlideSpec, TemplateManifest } from '../src/index.ts';
import { readRels, resolveTarget } from '../src/pptx/package.ts';
import { NS, childElement, childElements, descendants, parseXml, path, rootElement } from '../src/pptx/xml.ts';
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

  assert.match(requests[0]?.prompt ?? '', /- cards: A row of cards\.\n {2}slots: title \(text\); cardNText \(at most 20 characters\)\n {2}items: 1 to 5 cards/);
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
