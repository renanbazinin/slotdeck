// Behaviour added after the first live runs: item order in mirrored decks, decorations that
// must not outlive the sample content, and number patterns that tempt a model to invent data.
import assert from 'node:assert/strict';
import { test } from 'node:test';
import { placeholderPhoto } from '../examples/placeholder.ts';
import { PptxPackage, loadTemplate, planDeck, renderDeck, validateDeck } from '../src/index.ts';
import type { DeckSpec, JsonRequest, LlmProvider, PatternDef, SlideSpec, TemplateManifest } from '../src/index.ts';
import { slotSources } from '../src/template/manifest.ts';
import { NS, descendants, parseXml, rootElement } from '../src/pptx/xml.ts';
import { buildFixturePptx, fixtureManifest } from './fixture.ts';

const pptx = buildFixturePptx();
const photo = placeholderPhoto(100, 100, [0, 0, 0], [255, 255, 255]);

/** The fixture manifest with changes applied to its `content` pattern. */
function manifestWith(change: (pattern: PatternDef) => PatternDef): TemplateManifest {
  return { ...fixtureManifest, patterns: fixtureManifest.patterns.map((pattern) => (pattern.id === 'content' ? change(pattern) : pattern)) };
}

function shapeText(deckPptx: Uint8Array, shapeId: string): string | undefined {
  const root = rootElement(parseXml(PptxPackage.load(deckPptx).readText('ppt/slides/slide1.xml')));
  const props = descendants(root, NS.p, 'cNvPr').find((el) => el.getAttribute('id') === shapeId);
  const shape = props?.parentNode?.parentNode;
  if (!shape) return undefined;
  return descendants(shape as typeof root, NS.a, 't')
    .map((node) => node.textContent ?? '')
    .join('');
}

const slide = (slots: SlideSpec['slots'] = {}): SlideSpec => ({
  pattern: 'content',
  slots: { title: 'First', body: 'Body', photo: { asset: 'photo' }, note: 'Second', ...slots },
});

function scripted(...answers: unknown[]): { llm: LlmProvider; requests: JsonRequest[] } {
  const requests: JsonRequest[] = [];
  return {
    requests,
    llm: {
      async generateJson(request) {
        requests.push(request);
        return answers.shift();
      },
    },
  };
}

const offered = (request: JsonRequest | undefined): string[] =>
  (request?.schema as { properties: { slides: { items: { properties: { pattern: { enum: string[] } } } } } }).properties.slides.items.properties.pattern
    .enum;

// --- item order -------------------------------------------------------------

// Title (shape 2) and note (shape 6) stand in for the first and last card of a row.
const row = manifestWith((pattern) => ({ ...pattern, sequences: [[['title'], ['note']]] }));

test('a sequence keeps its order for readers of a deck that runs the other way', () => {
  const template = loadTemplate(pptx, row);
  const render = (direction: DeckSpec['direction']) =>
    renderDeck(template, { version: 1, direction, slides: [slide()] }, { assets: { photo } }).pptx;

  // Same direction as the template: every value goes to its own shape.
  assert.equal(shapeText(render('rtl'), '2'), 'First');
  assert.equal(shapeText(render('rtl'), '6'), 'Second');

  // Other direction: the first item takes the place of the last, and so on inwards.
  assert.equal(shapeText(render('ltr'), '2'), 'Second');
  assert.equal(shapeText(render('ltr'), '6'), 'First');
});

test('slotSources maps items mirror-wise and leaves other slots alone', () => {
  const pattern: PatternDef = {
    id: 'cards',
    sourceSlide: 1,
    description: '',
    slots: ['title', 'a1', 'b1', 'a2', 'b2', 'a3', 'b3'].map((name, index) => ({ name, kind: 'text', shapeId: String(index) })),
    sequences: [
      [
        ['a1', 'b1'],
        ['a2', 'b2'],
        ['a3', 'b3'],
      ],
    ],
  };
  assert.deepEqual(Object.fromEntries(slotSources(pattern, true)), { title: 'title', a1: 'a3', b1: 'b3', a2: 'a2', b2: 'b2', a3: 'a1', b3: 'b1' });
  assert.deepEqual(Object.fromEntries(slotSources(pattern, false)), { title: 'title', a1: 'a1', b1: 'b1', a2: 'a2', b2: 'b2', a3: 'a3', b3: 'b3' });
});

test('problems in a mirrored sequence are reported under the name the deck used', () => {
  // The title shape fits 20 characters. In an LTR deck it receives the value of `note`.
  const deck: DeckSpec = { version: 1, direction: 'ltr', slides: [slide({ note: 'A note that is longer than twenty characters' })] };
  const issues = validateDeck(row, deck, { assets: ['photo'] });
  assert.deepEqual(
    issues.map((issue) => [issue.severity, issue.slot]),
    [['warning', 'note']],
  );
});

// --- decorations ------------------------------------------------------------

test("a 'drop' slot is removed in strict layout and never offered to the model", async () => {
  const stickers = manifestWith((pattern) => ({
    ...pattern,
    slots: pattern.slots.map((slot) => (slot.name === 'icon' ? { ...slot, empty: 'drop' as const } : slot)),
  }));
  const template = loadTemplate(pptx, stickers);

  const rendered = renderDeck(template, { version: 1, slides: [slide()] }, { assets: { photo }, layout: 'strict' }).pptx;
  const xml = PptxPackage.load(rendered).readText('ppt/slides/slide1.xml');
  assert.ok(!xml.includes('name="Icon"'), 'the decoration is gone');
  assert.ok(xml.includes('name="Group"'), 'its group stays');

  const withValue = renderDeck(template, { version: 1, slides: [slide({ icon: { asset: 'photo' } })] }, { assets: { photo } }).pptx;
  assert.ok(PptxPackage.load(withValue).readText('ppt/slides/slide1.xml').includes('name="Icon"'), 'a deck may still fill it');

  const { llm, requests } = scripted({ title: 'T', slides: [{ pattern: 'content', purpose: 'p' }] }, { s1: {} });
  await planDeck(template, { prompt: 'x', language: 'he' }, { llm, icons: [{ id: 'i', description: 'i', data: photo }], maxRepairs: 0 }).catch(() => undefined);
  const s1 = (requests[1]?.schema as { properties: { s1: { properties: Record<string, unknown> } } }).properties.s1;
  assert.ok(!('icon' in s1.properties));
});

// --- number patterns --------------------------------------------------------

test('patterns built around numbers are offered only when the brief has figures', async () => {
  const numbers = manifestWith((pattern) => ({ ...pattern, figures: true }));
  const template = loadTemplate(pptx, numbers);
  const fields = { closing: 'Thanks' };
  const outline = { title: 'T', slides: [{ pattern: 'closing', purpose: 'p' }] };
  const ask = async (prompt: string, options: { numberPatterns?: 'from-brief' | 'always' | 'never'; instructions?: string } = {}) => {
    const { llm, requests } = scripted(outline, { s1: {} });
    const brief = options.instructions ? { prompt, fields, instructions: options.instructions } : { prompt, fields };
    await planDeck(template, brief, options.numberPatterns ? { llm, numberPatterns: options.numberPatterns } : { llm });
    return offered(requests[0]);
  };

  assert.deepEqual(await ask('A deck about dogs'), ['closing']);
  assert.deepEqual(await ask('We served 2341 customers this year'), ['content', 'closing']);
  assert.deepEqual(await ask('A deck about dogs', { instructions: 'Mention the 15% growth' }), ['content', 'closing']);
  assert.deepEqual(await ask('A deck about dogs', { numberPatterns: 'always' }), ['content', 'closing']);
  assert.deepEqual(await ask('We served 2341 customers', { numberPatterns: 'never' }), ['closing']);
});
