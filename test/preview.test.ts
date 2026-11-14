// Pictures of slides through a preview provider: the hook itself with a made-up provider, where
// the chat tools keep the pictures, the thumbnails of a template's patterns, and the adapter for
// the Pagus library on real decks.
import assert from 'node:assert/strict';
import { readFileSync } from 'node:fs';
import { dirname, join } from 'node:path';
import { test } from 'node:test';
import { fileURLToPath } from 'node:url';
import * as pagusCore from '@pagus-kit/core';
import * as pagusRenderer from '@pagus-kit/renderer';
import { placeholderPhoto } from '../examples/placeholder.ts';
import { PptxPackage, TOOL_NAMES, loadTemplate, memoryStore, pagusPreviews, patternPreviews, presentationTools, previewDeck, previewExtension, sampleDeck, storeKeys } from '../src/index.ts';
import type { DeckSpec, PreviewProvider, PreviewRequest, TemplateManifest } from '../src/index.ts';
import { readPresentation } from '../src/pptx/presentation.ts';
import { buildFixturePptx, fixtureManifest } from './fixture.ts';

const root = join(dirname(fileURLToPath(import.meta.url)), '..');
const demoManifest = JSON.parse(readFileSync(join(root, 'templates', 'demo.manifest.json'), 'utf8')) as TemplateManifest;
const demo = loadTemplate(readFileSync(join(root, 'templates', 'demo.pptx')), demoManifest);
const text = (data: Uint8Array) => new TextDecoder().decode(data);

/** A provider that draws nothing: one small SVG for each slide, naming the slide. It records what it was asked. */
function counting(): PreviewProvider & { requests: PreviewRequest[] } {
  const requests: PreviewRequest[] = [];
  return {
    requests,
    async renderSlides(request) {
      requests.push(request);
      const count = readPresentation(PptxPackage.load(request.pptx)).slides.length;
      return Array.from({ length: count }, (_, index) => index + 1)
        .filter((slide) => !request.slides || request.slides.includes(slide))
        .map((slide) => ({ slide, mimeType: 'image/svg+xml', data: new TextEncoder().encode(`<svg data-slide="${slide}"/>`), width: 960, height: 540 }));
    },
  };
}

test('a preview is named in the store by its type', () => {
  assert.equal(previewExtension({ mimeType: 'image/svg+xml' }), 'svg');
  assert.equal(previewExtension({ mimeType: 'image/png' }), 'png');
  assert.equal(previewExtension({ mimeType: 'image/jpeg' }), 'jpg');
  assert.equal(previewExtension({ mimeType: 'IMAGE/WEBP; q=1' }), 'webp');
  assert.equal(previewExtension({ mimeType: 'application/pdf' }), 'bin');
  assert.equal(storeKeys.preview('deck_7', 3, 'svg'), 'preview/deck_7/slide-03.svg');
});

// --- the chat tools ---------------------------------------------------------------------------

const outline = { title: 'Dogs', slides: [{ pattern: 'content', purpose: 'Why dogs' }, { pattern: 'closing', purpose: 'Thanks' }] };
const content = { s1: { title: 'Dogs are great', body: ['Loyal'], note: 'A note', photo: { prompt: 'A dog', alt: 'Dog' } }, s2: {} };

function tools(preview: PreviewProvider | undefined, ...answers: unknown[]) {
  const store = memoryStore();
  let next = 0;
  const options: Parameters<typeof presentationTools>[0] = {
    template: loadTemplate(buildFixturePptx(), fixtureManifest),
    llm: { generateJson: async () => answers.shift() },
    store,
    fields: { closing: 'Thank you' },
    images: { generateImage: async () => placeholderPhoto(20, 10, [0, 0, 0], [255, 255, 255]) },
    newId: (kind) => `${kind}_${++next}`,
  };
  if (preview) options.preview = preview;
  return { store, tools: presentationTools(options) };
}

test('the tools draw every presentation they create or change, and keep the pictures in the store', async () => {
  const preview = counting();
  const edit = {
    summary: 'Shortened slide 1.',
    language: 'en',
    title: 'Dogs',
    fields: { closing: 'Thank you' },
    slides: [
      { from: 1, pattern: 'content', change: 'Shorter title', images: false },
      { from: 2, pattern: 'closing', change: '', images: false },
    ],
  };
  const { store, tools: set } = tools(preview, outline, content, edit, { s1: { title: 'Dogs', body: ['Loyal'], note: 'A note' } });

  await set.call(TOOL_NAMES.plan, { topic: 'A deck about dogs', language: 'en', slide_count: 2 });
  assert.equal(preview.requests.length, 0, 'a plan has nothing to draw');

  const created = await set.call(TOOL_NAMES.create, { plan_id: 'plan_1' });
  assert.equal(created.output.presentation_id, 'deck_2');
  // The chat model gets the keys; the app gets the pictures themselves.
  assert.deepEqual(created.output.previews, ['preview/deck_2/slide-01.svg', 'preview/deck_2/slide-02.svg']);
  assert.deepEqual(created.previews?.map((picture) => [picture.slide, picture.mimeType, text(picture.data)]), [
    [1, 'image/svg+xml', '<svg data-slide="1"/>'],
    [2, 'image/svg+xml', '<svg data-slide="2"/>'],
  ]);
  assert.equal(text((await store.get('preview/deck_2/slide-02.svg')) as Uint8Array), '<svg data-slide="2"/>');
  // The provider was handed the file that was stored.
  assert.deepEqual(preview.requests[0]?.pptx, await store.get(storeKeys.pptx('deck_2')));

  const edited = await set.call(TOOL_NAMES.edit, { presentation_id: 'deck_2', instruction: 'Shorten the first title' });
  assert.equal(edited.output.presentation_id, 'deck_3');
  assert.deepEqual(edited.output.previews, ['preview/deck_3/slide-01.svg', 'preview/deck_3/slide-02.svg']);
  assert.ok(store.keys().includes('preview/deck_2/slide-01.svg'), 'the earlier version keeps its pictures');
});

test('a presentation stands when its pictures cannot be made, and nothing is drawn without a provider', async () => {
  const failing: PreviewProvider = {
    async renderSlides() {
      throw new Error('the renderer is not reachable');
    },
  };
  const broken = tools(failing, outline, content);
  await broken.tools.call(TOOL_NAMES.plan, { topic: 'A deck about dogs', language: 'en', slide_count: 2 });
  const created = await broken.tools.call(TOOL_NAMES.create, { plan_id: 'plan_1' });
  assert.equal(created.output.presentation_id, 'deck_2');
  assert.equal(created.output.preview_error, 'the renderer is not reachable');
  assert.equal(created.output.previews, undefined);
  assert.equal(created.previews, undefined);
  assert.ok(created.pptx && (await broken.store.get(storeKeys.pptx('deck_2'))));

  const plain = tools(undefined, outline, content);
  await plain.tools.call(TOOL_NAMES.plan, { topic: 'A deck about dogs', language: 'en', slide_count: 2 });
  const without = await plain.tools.call(TOOL_NAMES.create, { plan_id: 'plan_1' });
  assert.deepEqual([without.output.previews, without.output.preview_error, without.previews], [undefined, undefined, undefined]);
  assert.deepEqual(plain.store.keys().filter((key) => key.startsWith('preview/')), []);
});

// --- patterns and decks -----------------------------------------------------------------------

test('every pattern gets a thumbnail: its slide of the sample deck', async () => {
  const preview = counting();
  const pictures = await patternPreviews(demo, preview);
  assert.deepEqual(Object.keys(pictures), demoManifest.patterns.map((pattern) => pattern.id));
  // The sample deck has one slide for each pattern, in the manifest's order. A pattern derived
  // from another pattern's template slide therefore has a picture of its own.
  demoManifest.patterns.forEach((pattern, index) => assert.equal(text(pictures[pattern.id]?.data as Uint8Array), `<svg data-slide="${index + 1}"/>`));
  const derived = demoManifest.patterns.filter((pattern) => demoManifest.patterns.some((other) => other !== pattern && other.sourceSlide === pattern.sourceSlide));
  assert.ok(derived.length >= 2);
  assert.equal(readPresentation(PptxPackage.load(preview.requests[0]?.pptx as Uint8Array)).slides.length, demoManifest.patterns.length);
});

test('previewDeck renders a deck and draws it in one call', async () => {
  const { deck, assets } = sampleDeck(demo);
  const { pptx, previews } = await previewDeck(demo, deck, counting(), { assets, layout: 'adaptive' });
  assert.equal(readPresentation(PptxPackage.load(pptx)).slides.length, deck.slides.length);
  assert.deepEqual(previews.map((picture) => picture.slide), deck.slides.map((_, index) => index + 1));
});

// --- the Pagus adapter ------------------------------------------------------------------------

const pagus = pagusPreviews(pagusCore, pagusRenderer);

test('Pagus draws each slide as an SVG whose text is real text', async () => {
  const { deck, assets } = sampleDeck(demo);
  const { previews } = await previewDeck(demo, deck, pagus, { assets, layout: 'adaptive' });
  assert.equal(previews.length, demoManifest.patterns.length);
  for (const picture of previews) {
    assert.equal(picture.mimeType, 'image/svg+xml');
    assert.match(text(picture.data), /^<svg[^>]* viewBox="0 0 1279\.9\d* 720"/);
    assert.deepEqual([Math.round(picture.width ?? 0), picture.height], [1280, 720]);
  }
  // The title of the cover is in the picture as text, in the template's typeface and weight.
  const title = deck.slides[0]?.slots.title as string;
  assert.match(text(previews[0]?.data as Uint8Array), new RegExp(`font-family:'Cambria'[^"]*font-weight:bold[^"]*">${title}<`));
  // English paragraphs are left to run left to right.
  assert.doesNotMatch(text(previews[0]?.data as Uint8Array), /dir="rtl"/);

  // Only the slides asked for are drawn.
  const { pptx } = await previewDeck(demo, deck, counting(), { assets, layout: 'adaptive' });
  const some = await pagus.renderSlides({ pptx, slides: [2, 5] });
  assert.deepEqual(some.map((picture) => picture.slide), [2, 5]);
  assert.equal(text(some[0]?.data as Uint8Array), text(previews[1]?.data as Uint8Array));
});

test('Hebrew paragraphs are marked right to left, so a browser puts their punctuation and mixed text in order', async () => {
  const deck: DeckSpec = {
    version: 1,
    language: 'he',
    direction: 'rtl',
    fields: { organizationDepartment: 'אופני הנהר' },
    slides: [
      {
        pattern: 'text-only',
        slots: {
          title: 'כותרת בעברית',
          subtitle: 'An English line',
          body: ['שאלות: hello@example.org.', '2025 היה שנה טובה.'],
        },
      },
    ],
  };
  const { previews } = await previewDeck(demo, deck, pagus, { layout: 'adaptive' });
  const svg = text(previews[0]?.data as Uint8Array);
  const direction = (words: string) => new RegExp(`<p dir="rtl" [^>]*>(?:<[^>]+>)*${words}`).test(svg);
  assert.ok(direction('כותרת'), 'the Hebrew title');
  assert.ok(direction('שאלות: hello@example\\.org\\.'), 'Hebrew followed by an address');
  // A paragraph that starts with a number takes its direction from its first letter.
  assert.ok(direction('2025 '), 'a paragraph that starts with digits');
  assert.match(svg, /<p (?!dir=)[^>]*>(?:<[^>]+>)*An English line/);
});

test('the adapter passes its options on and stops when the caller gives up', async () => {
  const seen: unknown[] = [];
  const renderer = { renderSlide: (slide: unknown, size: unknown, options: unknown) => (seen.push(options), pagusRenderer.renderSlide(slide as never, size as never, options as never)) };
  const scaled = pagusPreviews(pagusCore, renderer, { scale: 0.5, fontSubstitutes: { Calibri: 'Carlito' } });
  const { deck, assets } = sampleDeck(demo);
  const { previews } = await previewDeck(demo, { ...deck, slides: deck.slides.slice(0, 1) }, scaled, { assets, layout: 'adaptive' });
  assert.deepEqual(seen, [{ scale: 0.5, fontSubstitutes: { Calibri: 'Carlito' } }]);
  assert.deepEqual([Math.round(previews[0]?.width ?? 0), previews[0]?.height], [640, 360]);
  // The substitute is named after the template's typeface, for a viewer that does not have that one.
  assert.match(text(previews[0]?.data as Uint8Array), /font-family:'Calibri', 'Carlito'/);

  const stopped = new AbortController();
  stopped.abort();
  await assert.rejects(pagus.renderSlides({ pptx: readFileSync(join(root, 'templates', 'demo.pptx')), signal: stopped.signal }), /abort/i);
});
