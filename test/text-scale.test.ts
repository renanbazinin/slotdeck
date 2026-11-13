// A template whose body text is too small to read can have it enlarged for every deck, in the
// manifest. Budgets shrink with it, tables and charts follow, and a deck's own size goes on top.
import assert from 'node:assert/strict';
import { test } from 'node:test';
import { strFromU8, unzipSync } from 'fflate';
import { changeManifest, loadTemplate, planOutline, renderDeck, sampleDeck, validateDeck } from '../src/index.ts';
import type { DeckSpec, JsonRequest, TemplateManifest } from '../src/index.ts';
import { buildFixturePptx, cardsManifest, tableManifest } from './fixture.ts';

const pptx = buildFixturePptx();
const picture = new Uint8Array(sampleDeck(loadTemplate(pptx, cardsManifest)).assets['sample-icon'] as Uint8Array);
const slideXml = (bytes: Uint8Array, number = 1) => strFromU8(unzipSync(bytes)[`ppt/slides/slide${number}.xml`] as Uint8Array);
const sizesOf = (xml: string, text: string) => {
  const end = xml.indexOf(`<a:t>${text}</a:t>`);
  return /sz="(\d+)"/.exec(xml.slice(xml.lastIndexOf('<a:r>', end), end))?.[1];
};

const cards = (texts: string[]): DeckSpec => ({
  version: 1,
  slides: [
    {
      pattern: 'cards',
      items: texts.length,
      slots: { title: 'A title', ...Object.fromEntries(texts.flatMap((text, index) => [[`card${index + 1}Text`, text], [`card${index + 1}Icon`, { asset: 'icon' }]])) },
    },
  ],
});

test('a text scale on the manifest enlarges body text in every deck and leaves titles alone', () => {
  const enlarged: TemplateManifest = { ...cardsManifest, textScale: 1.5 };
  const deck = cards(['Short', 'Words', 'Here']);
  const assets = { icon: picture };

  const plain = slideXml(renderDeck(loadTemplate(pptx, cardsManifest), deck, { assets }).pptx);
  const larger = slideXml(renderDeck(loadTemplate(pptx, enlarged), deck, { assets }).pptx);
  // The cards' text is 9 pt in the template: 13.5 pt at 1.5. The 26 pt title is not body text.
  assert.deepEqual([sizesOf(plain, 'Short'), sizesOf(larger, 'Short')], ['900', '1350']);
  assert.deepEqual([sizesOf(plain, 'A title'), sizesOf(larger, 'A title')], ['2600', '2600']);

  // A deck's own text size is applied on top of the template's: 9 pt x 1.5 x 1.25, to the half point.
  const both = slideXml(renderDeck(loadTemplate(pptx, enlarged), { ...deck, textScale: 1.25 }, { assets }).pptx);
  assert.equal(sizesOf(both, 'Short'), '1700');
});

test('budgets shrink with the template scale, for validation, the planner and the sample deck', async () => {
  const enlarged: TemplateManifest = { ...cardsManifest, textScale: 1.5 };
  const deck = cards(['Twelve chars', 'Words', 'Here']);
  // A card holds 20 characters as designed and 8 at 1.5: twelve fit the first and not the second.
  assert.deepEqual(validateDeck(cardsManifest, deck, { assets: ['icon'] }), []);
  assert.deepEqual(
    validateDeck(enlarged, deck, { assets: ['icon'] }).map((issue) => `${issue.slot}: ${issue.message}`),
    ['card1Text: text is 12 characters; the slot fits about 8'],
  );

  // The planning model is shown the smaller budget.
  const prompts: string[] = [];
  const llm = { generateJson: async (request: JsonRequest) => (prompts.push(request.prompt), { title: 'T', slides: [{ pattern: 'closing', purpose: 'x' }] }) };
  const icons = [{ id: 'icon', description: 'an icon', data: picture }];
  for (const manifest of [cardsManifest, enlarged]) {
    await planOutline(loadTemplate(pptx, manifest), { prompt: 'Dogs', slideCount: 1, fields: { closing: 'Bye' }, imageDensity: 'none' }, { llm, icons, layout: 'adaptive' });
  }
  assert.match(prompts[0] ?? '', /cardNText \(at most 20 characters\)/);
  assert.match(prompts[1] ?? '', /cardNText \(at most 8 characters\)/);

  // The sample deck is filled to the budgets real decks get.
  const template = loadTemplate(pptx, enlarged);
  const sample = sampleDeck(template);
  assert.deepEqual(validateDeck(enlarged, sample.deck, { assets: Object.keys(sample.assets) }), []);
  const cardText = sample.deck.slides.find((slide) => slide.pattern === 'cards')?.slots.card1Text as string;
  assert.ok(cardText.length <= 8, `"${cardText}" is within the smaller budget`);
});

test('table text that the manifest sizes follows the scale, and a cell holds less', () => {
  const enlarged: TemplateManifest = { ...tableManifest, textScale: 1.5 };
  const deck: DeckSpec = { version: 1, slides: [{ pattern: 'table', slots: { title: 'T', table: { header: ['Head', 'B', 'C'], rows: [['Seven c', 'y', 'z']] } } }] };
  const xml = slideXml(renderDeck(loadTemplate(pptx, enlarged), deck).pptx);
  // Header 10 pt and body 9 pt in the manifest: 15 and 13.5 at 1.5.
  assert.deepEqual([sizesOf(xml, 'Head'), sizesOf(xml, 'Seven c')], ['1500', '1350']);
  // A cell holds 10 characters as designed and 4 at 1.5.
  assert.deepEqual(validateDeck(tableManifest, deck), []);
  assert.deepEqual(validateDeck(enlarged, deck).map((issue) => issue.message), ['a cell is 7 characters; cells fit about 4']);
});

test('the text scale is set, checked and cleared through changeManifest', () => {
  const template = loadTemplate(pptx, cardsManifest);
  const set = changeManifest(template, [{ op: 'setManifest', textScale: 1.3 }]);
  assert.equal(set.manifest.textScale, 1.3);
  assert.match(set.applied[0]?.message ?? '', /shows body text at 1\.3 times its designed size in every deck, with budgets to match/);
  assert.equal(cardsManifest.textScale, undefined, 'the manifest that was passed in is not touched');

  const cleared = changeManifest(loadTemplate(pptx, set.manifest), [{ op: 'setManifest', textScale: null }]);
  assert.equal('textScale' in cleared.manifest, false);

  assert.match(changeManifest(template, [{ op: 'setManifest', textScale: 8 }]).rejected[0]?.reason ?? '', /a text scale is a factor between 0\.5 and 3/);
});
