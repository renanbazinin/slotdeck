// The demo template and manifest that ship with the repository: they must agree with each other,
// and the example decks must fit them, because they are the first thing a newcomer runs.
import assert from 'node:assert/strict';
import { readdirSync, readFileSync } from 'node:fs';
import { dirname, join } from 'node:path';
import { test } from 'node:test';
import { fileURLToPath } from 'node:url';
import { unzipSync, strFromU8 } from 'fflate';
import { checkManifest, loadTemplate, renderDeck, sampleDeck, studioView, validateDeck } from '../src/index.ts';
import type { DeckSpec, TemplateManifest } from '../src/index.ts';

const root = join(dirname(fileURLToPath(import.meta.url)), '..');
const manifest = JSON.parse(readFileSync(join(root, 'templates', 'demo.manifest.json'), 'utf8')) as TemplateManifest;
const template = loadTemplate(readFileSync(join(root, 'templates', 'demo.pptx')), manifest);
const icons = (JSON.parse(readFileSync(join(root, 'examples', 'icons', 'icons.json'), 'utf8')) as Array<{ id: string }>).map((icon) => icon.id);
const picture = new Uint8Array(readFileSync(join(root, 'examples', 'icons', 'star.png')));

const slideTexts = (pptx: Uint8Array, number: number): string[] =>
  [...strFromU8(unzipSync(pptx)[`ppt/slides/slide${number}.xml`] as Uint8Array).matchAll(/<a:t>([^<]*)<\/a:t>/g)].map((match) => match[1] as string);

test('the demo manifest matches the demo template', () => {
  assert.deepEqual(checkManifest(template), []);
  assert.deepEqual([manifest.language, manifest.direction, manifest.alignOnDirectionChange], ['en', 'ltr', 'flip']);
  assert.equal(manifest.patterns.length, 18);
  // Every slide of the template is in use. Five serve derived patterns too: a text-only slide, a
  // native chart, and the mirror image of each of the three slides with a photo.
  assert.deepEqual(studioView(template).slides.map((slide) => slide.patterns.length), [1, 3, 2, 2, 1, 1, 1, 1, 2, 1, 1, 1, 1]);
  assert.deepEqual(manifest.patterns.filter((pattern) => pattern.id.endsWith('-mirrored')).map((pattern) => pattern.id), ['text-and-photo-mirrored', 'photo-stat-mirrored', 'list-with-photo-mirrored']);
});

test('a sample deck with full rows renders from the demo template', () => {
  const { deck, assets } = sampleDeck(template, { items: 'max' });
  assert.deepEqual(validateDeck(manifest, deck, { assets: Object.keys(assets), layout: 'adaptive' }), []);
  const { pptx, warnings } = renderDeck(template, deck, { assets, layout: 'adaptive' });
  assert.deepEqual(warnings, []);
  assert.equal(Object.keys(unzipSync(pptx)).filter((name) => /^ppt\/slides\/slide\d+\.xml$/.test(name)).length, 18);
});

test('the example decks fit the demo template without a warning', () => {
  const dir = join(root, 'examples', 'decks');
  const assets = Object.fromEntries(['photo-tall', 'photo-square', 'photo-wide', 'chart-months', 'chart-wide', ...icons].map((id) => [id, picture]));
  const files = readdirSync(dir).filter((name) => name.endsWith('.deck.json'));
  assert.ok(files.length >= 2);
  for (const file of files) {
    const deck = JSON.parse(readFileSync(join(dir, file), 'utf8')) as DeckSpec;
    assert.deepEqual(validateDeck(manifest, deck, { assets: Object.keys(assets), layout: 'adaptive' }), [], file);
    const { pptx, warnings } = renderDeck(template, deck, { assets, layout: 'adaptive' });
    assert.deepEqual(warnings, [], file);
    assert.ok(!slideTexts(pptx, 1).some((text) => /Presentation title/.test(text)), `${file}: the template's filler is replaced`);
  }
});

test('a right-to-left deck on the left-to-right template mirrors its text and keeps the frame', () => {
  const deck = JSON.parse(readFileSync(join(root, 'examples', 'decks', 'patterns.he.deck.json'), 'utf8')) as DeckSpec;
  const assets = Object.fromEntries(['photo-tall', 'photo-square', ...icons].map((id) => [id, picture]));
  const xml = strFromU8(unzipSync(renderDeck(template, deck, { assets, layout: 'adaptive' }).pptx)['ppt/slides/slide2.xml'] as Uint8Array);
  const paragraphOf = (text: string): string => {
    const end = xml.indexOf(text);
    return xml.slice(xml.lastIndexOf('<a:p>', end), end);
  };
  // The title was left-aligned English; in Hebrew it is right-aligned, right to left, and its run
  // is marked Hebrew so that PowerPoint places punctuation on the correct side.
  const title = paragraphOf('האופניים הפכו');
  assert.match(title, /<a:pPr algn="r" rtl="1"/);
  assert.match(title, /<a:rPr [^>]*lang="he-IL"/);
  // The line beside the logo is a deck field that keeps its alignment.
  assert.match(paragraphOf('אופני הנהר'), /<a:pPr algn="l" rtl="1"/);
});

test('a chart format the renderer does not know is reported, not silently dropped', () => {
  const deck: DeckSpec = {
    version: 1,
    fields: { organizationDepartment: 'x' },
    slides: [
      {
        pattern: 'chart-native',
        slots: {
          title: 'T',
          chartDescription: 'D',
          chart: { type: 'column', categories: ['a'], series: [{ name: 's', values: [1] }], format: '0.0' as 'number' },
        },
      },
    ],
  };
  assert.deepEqual(
    validateDeck(manifest, deck, { layout: 'adaptive' }).map((issue) => `${issue.severity} ${issue.slot}: ${issue.message}`),
    ['warning chart: chart format "0.0" is not known and is ignored; use "number" or "percent"'],
  );
});
