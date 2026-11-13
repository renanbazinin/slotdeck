// The pieces a template studio needs: draft a manifest with a model, check it against the
// pptx, and render a sample deck that pushes every slot to its limit.
import assert from 'node:assert/strict';
import { test } from 'node:test';
import { checkManifest, draftManifest, gemini, loadTemplate, renderDeck, sampleDeck, validateDeck } from '../src/index.ts';
import type { JsonRequest, LlmProvider, TemplateManifest } from '../src/index.ts';
import { buildFixturePptx, cardsManifest, fixtureManifest, gridManifest, tableManifest } from './fixture.ts';

const pptx = buildFixturePptx();

function scripted(answers: (request: JsonRequest) => unknown): { llm: LlmProvider; requests: JsonRequest[] } {
  const requests: JsonRequest[] = [];
  return {
    requests,
    llm: {
      async generateJson(request) {
        requests.push(request);
        return answers(request);
      },
    },
  };
}

test('the checker passes the manifests used in these tests', () => {
  for (const manifest of [fixtureManifest, cardsManifest, gridManifest, tableManifest]) {
    assert.deepEqual(checkManifest(loadTemplate(pptx, manifest)).filter((issue) => issue.severity === 'error'), []);
  }
});

test('the checker finds a manifest that no longer matches its pptx', () => {
  const broken: TemplateManifest = {
    ...fixtureManifest,
    patterns: [
      {
        id: 'content',
        sourceSlide: 1,
        description: '',
        slots: [
          { name: 'title', kind: 'text', shapeId: '2', maxChars: 10 },
          { name: 'title', kind: 'text', shapeId: '3', maxChars: 10 },
          { name: 'gone', kind: 'text', shapeId: '99', maxChars: 10 },
          { name: 'photo', kind: 'image', imageKind: 'photo', shapeId: '6' },
          { name: 'moved', kind: 'image', imageKind: 'icon', shapeId: '8', box: { x: 0, y: 0, w: 0.1, h: 0.1 } },
          { name: 'open', kind: 'text', shapeId: '6' },
        ],
        repeat: { item: 'card', slots: ['card{n}Text'], shapes: [['2'], ['77']] },
      },
      { id: 'content', sourceSlide: 9, description: '', slots: [] },
    ],
  };
  const messages = checkManifest(loadTemplate(pptx, broken)).map((issue) => `${issue.severity} ${issue.slot ?? issue.pattern}: ${issue.message}`);
  for (const expected of [
    /error title: another slot of the pattern has the same name/,
    /error gone: points at shape #99, which is not on slide 1/,
    /error photo: is an image slot, but shape #6 is a text box/,
    /error moved: has a box, but its shape is inside a group/,
    /warning open: has no character budget/,
    /warning open: points at the same shape as "photo"/,
    /error content: repeat expects a slot named "card1Text"/,
    /error content: repeat item 2 points at shape #77/,
    /error content: another pattern has the same id/,
    /error content: points at slide 9, but the template has 5 slides/,
  ]) {
    assert.ok(messages.some((message) => expected.test(message)), `${expected} not in:\n${messages.join('\n')}`);
  }
});

test('the sample deck uses every pattern and fills each slot close to its limit', () => {
  const template = loadTemplate(pptx, cardsManifest);
  const { deck, assets } = sampleDeck(template);
  assert.deepEqual(deck.slides.map((slide) => slide.pattern), ['content', 'closing', 'cards']);
  assert.equal(deck.direction, 'rtl');

  const title = deck.slides[0]?.slots.title as string;
  assert.ok(title.length > 10 && title.length <= 20, `title is ${title.length} characters of a budget of 20`);
  assert.match(title, /^SAMPLE/, "the filler is made of the template's own words");
  assert.equal(deck.slides[2]?.items, 3);
  assert.deepEqual(deck.fields, { closing: 'SAMPLE CLOSING' });
  assert.equal((deck.slides[0]?.slots.body as string[]).length, 3, 'a list slot gets as many paragraphs as it allows');

  assert.deepEqual(validateDeck(cardsManifest, deck, { assets: Object.keys(assets) }), [], 'the sample stays inside every limit');
  assert.ok(renderDeck(template, deck, { assets }).pptx.length > 1000);
  assert.equal(sampleDeck(template, { items: 'max' }).deck.slides[2]?.items, 5);
});

test('draftManifest turns model labels and measured facts into a manifest', async () => {
  const { llm, requests } = scripted((request) => {
    if (request.prompt.startsWith('Sample texts')) return { language: 'he', fields: [] };
    if (request.prompt.startsWith('Slide 4')) {
      return {
        usable: true,
        reason: 'a row of cards',
        id: 'Card Row',
        description: 'A row of cards with an icon each.',
        figures: false,
        shapes: [
          { shape: '2', role: 'title', name: 'title', optional: false },
          { shape: '11', role: 'text', name: 'card1Text', optional: false },
          { shape: '12', role: 'icon', name: 'card1Icon', optional: false },
          { shape: '13', role: 'decoration', name: 'firstLabel', optional: false },
          { shape: '21', role: 'text', name: 'card2Text', optional: false },
          { shape: '22', role: 'icon', name: 'card2Icon', optional: false },
          { shape: '31', role: 'text', name: 'card3Text', optional: false },
          { shape: '32', role: 'icon', name: 'card3Icon', optional: false },
        ],
      };
    }
    if (request.prompt.startsWith('Slide 5')) {
      return {
        usable: true,
        reason: '',
        id: 'card-grid',
        description: 'A grid.',
        figures: true,
        shapes: [
          { shape: '2', role: 'title', name: 'title', optional: true },
          ...[11, 21, 31, 41].map((id, index) => ({ shape: String(id), role: 'number', name: `stat${index + 1}`, optional: false })),
          { shape: '12', role: 'fixed', name: 'x', optional: false },
        ],
      };
    }
    if (request.prompt.startsWith('Slide 3')) {
      return { usable: true, reason: '', id: 'card-grid', description: 'A table.', figures: false, shapes: [{ shape: '2', role: 'title', name: 'title', optional: false }, { shape: '3', role: 'table', name: 'table', optional: false }] };
    }
    return { usable: false, reason: 'A sheet of brand assets.', id: '', description: '', figures: false, shapes: [] };
  });
  const picture = { mimeType: 'image/png', data: new Uint8Array([1, 2, 3]) };
  const progress: number[] = [];
  const { manifest, skipped, issues } = await draftManifest(pptx, {
    llm,
    name: 'fixture',
    slides: [2, 3, 4, 5],
    // A picture may arrive later, as it does when the app renders or fetches it on request.
    slideImage: async (slide) => (slide === 4 ? picture : undefined),
    onProgress: (event) => progress.push(event.done),
  });

  // The reading direction is taken from the letters of the sample text, which in the fixture is English.
  assert.deepEqual([manifest.name, manifest.language, manifest.direction, manifest.alignOnDirectionChange], ['fixture', 'he', 'ltr', 'keep']);
  assert.deepEqual(manifest.patterns.map((pattern) => [pattern.id, pattern.sourceSlide]), [['card-grid', 3], ['card-row', 4], ['card-grid-5', 5]], 'ids are made unique');
  assert.deepEqual(skipped, [{ slide: 2, reason: 'A sheet of brand assets.' }]);
  assert.deepEqual(issues.filter((issue) => issue.severity === 'error'), [], 'the draft matches the pptx it was drafted from');
  assert.deepEqual(progress.sort(), [1, 2, 3, 4]);

  // The model saw measured facts, and the picture where one was offered.
  const slide4 = requests.find((request) => request.prompt.startsWith('Slide 4'));
  assert.match(slide4?.prompt ?? '', /#11 text at x66 y54, 11 wide 8 high, in group #10 \| 9 pt \| 1 paragraph, 11 characters: "SAMPLE CARD"/);
  assert.match(slide4?.prompt ?? '', /#12 picture at x69 y41, 4 wide 8 high, in group #10 \| 8x8 px png, same picture on 3 slides/);
  assert.match(slide4?.prompt ?? '', /The template reads left to right\./);
  assert.deepEqual(slide4?.images, [picture]);
  assert.equal(requests.find((request) => request.prompt.startsWith('Slide 5'))?.images, undefined);

  // Budgets come from the box and font; roles decide the kind of slot.
  const row = manifest.patterns[1];
  assert.deepEqual(row?.slots.find((slot) => slot.name === 'card1Text'), { name: 'card1Text', kind: 'text', shapeId: '11', size: 9, maxChars: 40 });
  assert.deepEqual(row?.slots.find((slot) => slot.name === 'card1Icon'), { name: 'card1Icon', kind: 'image', imageKind: 'icon', shapeId: '12' });
  assert.equal(row?.slots.find((slot) => slot.name === 'firstLabel')?.empty, 'drop', 'a decoration is dropped when the content changes');
  // Numbered names become a row, with each item's outermost shape.
  assert.deepEqual(row?.repeat, { item: 'card', slots: ['card{n}Text', 'card{n}Icon'], shapes: [['10'], ['20'], ['30']], min: 2, max: 3 });

  const grid = manifest.patterns[2];
  assert.equal(grid?.figures, true);
  assert.deepEqual(grid?.repeat, { item: 'stat', slots: ['stat{n}'], shapes: [['10'], ['20'], ['30'], ['40']], min: 2, max: 4, columns: 2 });
  assert.equal(grid?.slots.find((slot) => slot.name === 'title')?.empty, 'remove');
  assert.equal(grid?.slots.find((slot) => slot.name === 'stat1' && slot.kind === 'text')?.kind === 'text' && (grid.slots[1] as { maxChars: number }).maxChars, 11, 'a big figure holds what its sample holds');

  // A table cell holds what fits before the table leaves the slide: here three lines of 18 pt
  // text in each of three rows, in columns 157 pt wide.
  assert.deepEqual(manifest.patterns[0]?.slots.find((slot) => slot.kind === 'table'), { name: 'table', kind: 'table', shapeId: '3', maxColumns: 3, maxRows: 2, maxCellChars: 40 });

  // A drafted manifest is usable straight away: its sample deck renders.
  const template = loadTemplate(pptx, manifest);
  const { deck, assets } = sampleDeck(template);
  assert.ok(renderDeck(template, deck, { assets }).pptx.length > 1000);
});

test('gemini() gives up on a request that does not answer', async () => {
  const never = ((_url: string | URL | Request, init?: RequestInit) =>
    new Promise<Response>((_resolve, reject) => init?.signal?.addEventListener('abort', () => reject(init.signal?.reason)))) as typeof fetch;
  await assert.rejects(
    gemini({ model: 'm', apiKey: 'k', fetch: never, timeoutMs: 20 }).generateJson({ system: 's', prompt: 'p', schema: {} }),
    /Gemini did not answer within 0 seconds/,
  );
  // The caller's own signal still ends the request, and is not reported as a timeout.
  const stop = new AbortController();
  const pending = gemini({ model: 'm', apiKey: 'k', fetch: never }).generateJson({ system: 's', prompt: 'p', schema: {}, signal: stop.signal });
  stop.abort(new Error('stopped by the caller'));
  await assert.rejects(pending, /stopped by the caller/);
});

test('gemini() sends pictures as image blocks next to the prompt', async () => {
  let body: { input?: unknown } = {};
  const fetchMock = (async (_url: string | URL | Request, init?: RequestInit) => {
    body = JSON.parse(String(init?.body));
    return new Response(JSON.stringify({ status: 'completed', steps: [{ type: 'model_output', content: [{ type: 'text', text: '{}' }] }] }));
  }) as typeof fetch;
  await gemini({ model: 'm', apiKey: 'k', fetch: fetchMock }).generateJson({
    system: 's',
    prompt: 'Look at this slide',
    schema: {},
    images: [{ mimeType: 'image/png', data: new Uint8Array([1, 2, 3]) }],
  });
  assert.deepEqual(body.input, [
    { type: 'user_input', content: [{ type: 'text', text: 'Look at this slide' }, { type: 'image', mime_type: 'image/png', data: 'AQID' }] },
  ]);
});
