import assert from 'node:assert/strict';
import { test } from 'node:test';
import { placeholderPhoto } from '../examples/placeholder.ts';
import {
  DeckValidationError,
  GeminiError,
  PlanError,
  PptxPackage,
  directionOf,
  gemini,
  geminiImages,
  generateDeck,
  loadTemplate,
  planDeck,
  planOutline,
  resolveAssets,
  vertexEndpoint,
} from '../src/index.ts';
import type { GenerateProgress, IconDef, ImageRef, ImageRequest, JsonRequest, LlmProvider, TemplateManifest } from '../src/index.ts';
import { nearestAspectRatio } from '../src/llm/gemini.ts';
import { readPresentation } from '../src/pptx/presentation.ts';
import { buildFixturePptx, cardsManifest, fixtureManifest } from './fixture.ts';

const pptx = buildFixturePptx();
const template = loadTemplate(pptx, fixtureManifest);
const png = placeholderPhoto(20, 10, [0, 0, 0], [255, 255, 255]);

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

type Schema = { properties: Record<string, Schema>; required?: string[]; enum?: string[]; items?: Schema; maxItems?: number; type?: string };
const schemaOf = (request: JsonRequest | undefined) => request?.schema as unknown as Schema;

const outline = { title: 'Dogs', slides: [{ pattern: 'content', purpose: 'Why dogs' }, { pattern: 'closing', purpose: 'Thanks' }] };
const content = {
  s1: { title: 'Dogs are great', body: ['Loyal', 'Playful'], photo: { prompt: 'A dog in a park', alt: 'Dog' }, note: 'A note' },
  s2: {},
};
const brief = { prompt: 'A deck about dogs', language: 'en', fields: { closing: 'Thank you' } };

test('plans a deck in two calls: outline, then content', async () => {
  const { llm, requests } = scripted(outline, content);
  const { deck, warnings } = await planDeck(template, brief, { llm });

  assert.equal(requests.length, 2);
  assert.deepEqual(warnings, []);
  assert.equal(deck.title, 'Dogs');
  assert.equal(deck.language, 'en');
  assert.equal(deck.direction, 'ltr');
  assert.deepEqual(deck.fields, { closing: 'Thank you' });
  assert.deepEqual(
    deck.slides.map((slide) => slide.pattern),
    ['content', 'closing'],
  );
  const first = deck.slides[0]?.slots ?? {};
  assert.equal(first.title, 'Dogs are great');
  assert.deepEqual(first.body, ['Loyal', 'Playful']);
  const photo = first.photo as ImageRef;
  assert.match(photo.asset, /^gen-[0-9a-f]{16}$/);
  assert.equal(photo.prompt, 'A dog in a park');
  assert.equal(photo.alt, 'Dog');

  assert.match(requests[0]?.system ?? '', /BCP 47 tag "en"/);
  assert.match(requests[0]?.prompt ?? '', /A deck about dogs/);
  assert.match(requests[1]?.prompt ?? '', /s1: content\. Why dogs/);
});

test('the content schema holds exactly the chosen slides and the slots the model may write', async () => {
  const { llm, requests } = scripted(outline, content);
  await planDeck(template, brief, { llm });

  const outlineSchema = schemaOf(requests[0]);
  assert.deepEqual(outlineSchema.properties.slides?.items?.properties.pattern?.enum, ['content', 'closing']);

  const schema = schemaOf(requests[1]);
  assert.deepEqual(Object.keys(schema.properties), ['s1', 's2']);
  const s1 = schema.properties.s1 as Schema;
  // `icon` keeps the template's own picture, so the model is not asked for it.
  assert.deepEqual(Object.keys(s1.properties), ['title', 'body', 'photo', 'note']);
  assert.deepEqual(s1.required, ['title', 'body', 'photo', 'note'], 'strict layout: removable slots must be filled');
  assert.equal(s1.properties.body?.type, 'array');
  assert.equal(s1.properties.body?.maxItems, 3);
  assert.deepEqual(s1.properties.photo?.required, ['prompt', 'alt']);
  // The closing title comes from a deck field, so that slide has nothing left to write.
  assert.deepEqual(Object.keys((schema.properties.s2 as Schema).properties), []);
  assert.ok(!JSON.stringify(schema).includes('anyOf'));
});

test('adaptive layout makes removable slots optional', async () => {
  const { llm, requests } = scripted(outline, { s1: { ...content.s1, note: undefined }, s2: {} });
  const { deck } = await planDeck(template, brief, { llm, layout: 'adaptive' });
  assert.deepEqual((schemaOf(requests[1]).properties.s1 as Schema).required, ['title', 'body', 'photo']);
  assert.ok(!('note' in (deck.slides[0]?.slots ?? {})));
});

test('patterns the caller cannot fill are not offered', async () => {
  // No `closing` deck field: the closing pattern is unusable.
  const noField = scripted({ title: 'T', slides: [{ pattern: 'content', purpose: 'p' }] }, { s1: content.s1 });
  await planDeck(template, { prompt: 'x', language: 'en' }, { llm: noField.llm });
  assert.deepEqual(schemaOf(noField.requests[0]).properties.slides?.items?.properties.pattern?.enum, ['content']);

  // No photos: the content pattern needs one, so only closing is left.
  const noPhotos = scripted({ title: 'T', slides: [{ pattern: 'closing', purpose: 'p' }] }, { s1: {} });
  await planDeck(template, { ...brief, imageDensity: 'none' }, { llm: noPhotos.llm });
  assert.deepEqual(schemaOf(noPhotos.requests[0]).properties.slides?.items?.properties.pattern?.enum, ['closing']);

  await assert.rejects(planDeck(template, { prompt: 'x', imageDensity: 'none' }, { llm: scripted().llm }), (error: unknown) => {
    assert.ok(error instanceof PlanError);
    assert.match(error.message, /content needs a photo/);
    assert.match(error.message, /closing needs the deck field "closing"/);
    return true;
  });
});

test('icon slots are offered only with an icon set, as a choice of ids', async () => {
  const withIcon: TemplateManifest = {
    ...fixtureManifest,
    patterns: fixtureManifest.patterns.map((pattern) => ({
      ...pattern,
      slots: pattern.slots.map((slot) => (slot.name === 'icon' ? { ...slot, empty: 'error' as const } : slot)),
    })),
  };
  const iconTemplate = loadTemplate(pptx, withIcon);
  const icons: IconDef[] = [
    { id: 'clock', description: 'a clock', data: png },
    { id: 'paw', description: 'a paw print', data: png },
  ];

  const { llm, requests } = scripted(outline, { s1: { ...content.s1, icon: 'paw' }, s2: {} });
  const { deck } = await planDeck(iconTemplate, brief, { llm, icons });
  assert.deepEqual((schemaOf(requests[1]).properties.s1 as Schema).properties.icon?.enum, ['clock', 'paw']);
  assert.match(requests[1]?.prompt ?? '', /- paw: a paw print/);
  assert.deepEqual(deck.slides[0]?.slots.icon, { asset: 'paw' });

  const without = scripted({ title: 'T', slides: [{ pattern: 'closing', purpose: 'p' }] }, { s1: {} });
  await planDeck(iconTemplate, brief, { llm: without.llm });
  assert.deepEqual(schemaOf(without.requests[0]).properties.slides?.items?.properties.pattern?.enum, ['closing']);
});

test('a draft over a slot budget is sent back with the list of problems', async () => {
  const tooLong = { s1: { ...content.s1, title: 'This title is far too long for it' }, s2: {} };
  const { llm, requests } = scripted(outline, tooLong, content);
  const { deck, warnings } = await planDeck(template, brief, { llm });

  assert.equal(requests.length, 3);
  assert.match(requests[2]?.prompt ?? '', /s1\.title: text is 33 characters; the slot fits about 20/);
  assert.match(requests[2]?.prompt ?? '', /This title is far too long for it/, 'the previous answer is included');
  assert.equal(deck.slides[0]?.slots.title, 'Dogs are great');
  assert.deepEqual(warnings, []);
});

test('remaining warnings are returned; remaining errors throw', async () => {
  const tooLong = { s1: { ...content.s1, title: 'This title is far too long for it' }, s2: {} };
  const stillLong = scripted(outline, tooLong, tooLong);
  const result = await planDeck(template, brief, { llm: stillLong.llm, maxRepairs: 1 });
  assert.equal(result.warnings.length, 1);
  assert.equal(stillLong.requests.length, 3);

  const missing = { s1: { body: ['x'], photo: content.s1.photo, note: 'n' }, s2: {} };
  await assert.rejects(planDeck(template, brief, { llm: scripted(outline, missing, missing, missing).llm }), DeckValidationError);
});

test('an outline with unknown patterns is retried once, then rejected', async () => {
  const bad = { title: 'T', slides: [{ pattern: 'made-up', purpose: 'p' }] };
  const recovered = scripted(bad, outline, content);
  const { deck } = await planDeck(template, brief, { llm: recovered.llm });
  assert.equal(deck.slides.length, 2);
  assert.match(recovered.requests[1]?.prompt ?? '', /not in the list: made-up/);

  await assert.rejects(planDeck(template, brief, { llm: scripted(bad, bad).llm }), PlanError);
});

test('text from the model is tidied: whitespace collapsed, bullet marks dropped', async () => {
  const messy = { s1: { ...content.s1, title: '  Dogs\nare   great ', body: ['• Loyal', '- Playful', ''] }, s2: {} };
  const { deck } = await planDeck(template, brief, { llm: scripted(outline, messy).llm });
  assert.equal(deck.slides[0]?.slots.title, 'Dogs are great');
  assert.deepEqual(deck.slides[0]?.slots.body, ['Loyal', 'Playful']);
});

test('pictures are generated for the frame they go into, and cached by description', async () => {
  const { deck } = await planDeck(template, brief, { llm: scripted(outline, content).llm });
  const calls: ImageRequest[] = [];
  const images = {
    async generateImage(request: ImageRequest) {
      calls.push(request);
      return png;
    },
  };
  const cache = new Map<string, Uint8Array>();

  const assets = await resolveAssets(template, deck, { images, cache });
  assert.equal(calls.length, 1);
  assert.equal(calls[0]?.prompt, 'A dog in a park');
  assert.equal(calls[0]?.kind, 'photo');
  // The fixture photo frame is 4000000 x 2000000; the ratio is derived from rounded slide fractions.
  assert.ok(Math.abs((calls[0]?.aspectRatio ?? 0) - 2) < 0.01, `aspect ratio ${calls[0]?.aspectRatio}`);
  assert.deepEqual(Object.values(assets), [png]);

  await resolveAssets(template, deck, { images, cache });
  assert.equal(calls.length, 1, 'second run reuses the cache');

  await assert.rejects(resolveAssets(template, deck, {}), /no image provider was given/);
});

test('generateDeck goes from brief to a pptx and reports progress', async () => {
  const events: GenerateProgress[] = [];
  const result = await generateDeck(template, brief, {
    llm: scripted(outline, content).llm,
    images: { generateImage: async () => png },
    onProgress: (event) => events.push(event),
  });

  assert.deepEqual(
    events.map((event) => event.stage),
    ['outline', 'write', 'images', 'render'],
  );
  const pkg = PptxPackage.load(result.pptx);
  assert.equal(readPresentation(pkg).slides.length, 2);
  assert.ok(pkg.readText('ppt/slides/slide1.xml').includes('Dogs are great'));
  assert.ok(pkg.readText('ppt/slides/slide2.xml').includes('Thank you'));
  assert.equal(Object.keys(result.assets).length, 1);
});

test('without an image provider the deck is planned with no photos', async () => {
  const { llm, requests } = scripted({ title: 'T', slides: [{ pattern: 'closing', purpose: 'p' }] }, { s1: {} });
  const result = await generateDeck(template, brief, { llm });
  assert.match(requests[0]?.prompt ?? '', /Do not plan any photos/);
  assert.equal(readPresentation(PptxPackage.load(result.pptx)).slides.length, 1);
});

test('language decides the default direction', () => {
  assert.equal(directionOf('he'), 'rtl');
  assert.equal(directionOf('ar-EG'), 'rtl');
  assert.equal(directionOf('en-GB'), 'ltr');
});

// ---------------------------------------------------------------------------
// Gemini adapter, against a fake fetch.

function fakeFetch(status: number, body: unknown): { fetch: typeof fetch; calls: Array<{ url: string; init: RequestInit }> } {
  const calls: Array<{ url: string; init: RequestInit }> = [];
  const fake = (async (url: string | URL | Request, init?: RequestInit) => {
    calls.push({ url: String(url), init: init ?? {} });
    return new Response(JSON.stringify(body), { status });
  }) as typeof fetch;
  return { fetch: fake, calls };
}

test('gemini() asks the Interactions API for JSON and reads the model output', async () => {
  const { fetch, calls } = fakeFetch(200, {
    status: 'completed',
    steps: [
      { type: 'thought' },
      { type: 'model_output', content: [{ type: 'text', text: '{"answer":' }] },
      { type: 'model_output', content: [{ type: 'text', text: ' 42}' }] },
    ],
  });
  const llm = gemini({ model: 'test-model', apiKey: 'test-key', thinkingLevel: 'low', fetch });
  const value = await llm.generateJson({ system: 'SYSTEM', prompt: 'PROMPT', schema: { type: 'object' } });

  assert.deepEqual(value, { answer: 42 });
  assert.equal(calls[0]?.url, 'https://generativelanguage.googleapis.com/v1/interactions');
  const headers = calls[0]?.init.headers as Record<string, string>;
  assert.equal(headers['x-goog-api-key'], 'test-key');
  assert.deepEqual(JSON.parse(String(calls[0]?.init.body)), {
    model: 'test-model',
    input: 'PROMPT',
    system_instruction: 'SYSTEM',
    response_format: { type: 'text', mime_type: 'application/json', schema: { type: 'object' } },
    store: false,
    // No estimate came with the request, so the limit is the default room for an answer plus room for reasoning.
    generation_config: { max_output_tokens: 24576, thinking_level: 'low' },
  });
});

test('gemini() can target Vertex AI with a bearer token', async () => {
  const { fetch, calls } = fakeFetch(200, { status: 'completed', steps: [{ type: 'model_output', content: [{ type: 'text', text: '{}' }] }] });
  const llm = gemini({
    model: 'test-model',
    endpoint: vertexEndpoint('my-project'),
    headers: async () => ({ Authorization: 'Bearer token' }),
    fetch,
  });
  await llm.generateJson({ system: '', prompt: 'p', schema: {} });
  assert.equal(calls[0]?.url, 'https://aiplatform.googleapis.com/v1beta1/projects/my-project/locations/global/interactions');
  const headers = calls[0]?.init.headers as Record<string, string>;
  assert.equal(headers.Authorization, 'Bearer token');
  assert.equal(headers['x-goog-api-key'], undefined);
});

test('gemini() turns API failures into errors that carry the status', async () => {
  const denied = fakeFetch(400, { error: { message: 'Schema is too complex' } });
  await assert.rejects(gemini({ model: 'm', apiKey: 'k', fetch: denied.fetch }).generateJson({ system: '', prompt: '', schema: {} }), (error: unknown) => {
    assert.ok(error instanceof GeminiError);
    assert.equal(error.status, 400);
    assert.match(error.message, /Schema is too complex/);
    assert.ok(!error.message.includes('k'.repeat(2)));
    return true;
  });

  const failed = fakeFetch(200, { status: 'failed', steps: [], errors: [{ message: 'blocked' }] });
  await assert.rejects(gemini({ model: 'm', fetch: failed.fetch }).generateJson({ system: '', prompt: '', schema: {} }), /status "failed": blocked/);

  const notJson = fakeFetch(200, { status: 'completed', steps: [{ type: 'model_output', content: [{ type: 'text', text: 'sorry' }] }] });
  await assert.rejects(gemini({ model: 'm', fetch: notJson.fetch }).generateJson({ system: '', prompt: '', schema: {} }), /did not return valid JSON/);
});

test('geminiImages() requests the closest supported ratio and decodes the picture', async () => {
  const base64 = Buffer.from(png).toString('base64');
  const { fetch, calls } = fakeFetch(200, {
    status: 'completed',
    steps: [{ type: 'model_output', content: [{ type: 'text', text: 'Here you go' }, { type: 'image', mime_type: 'image/png', data: base64 }] }],
  });
  const images = geminiImages({ model: 'image-model', apiKey: 'k', style: 'Natural light.', fetch });
  const data = await images.generateImage({ prompt: 'A dog', kind: 'photo', aspectRatio: 2 });

  assert.deepEqual(data, png);
  assert.deepEqual(JSON.parse(String(calls[0]?.init.body)), {
    model: 'image-model',
    input: [{ type: 'user_input', content: [{ type: 'text', text: 'A dog\n\nNatural light.' }] }],
    response_format: [{ type: 'text' }, { type: 'image', aspect_ratio: '16:9', image_size: '1K' }],
    store: false,
  });

  const none = fakeFetch(200, { status: 'completed', steps: [{ type: 'model_output', content: [{ type: 'text', text: 'no' }] }] });
  await assert.rejects(geminiImages({ model: 'm', fetch: none.fetch }).generateImage({ prompt: 'x', kind: 'photo', aspectRatio: 1 }), /returned no image/);
});

test('nearest supported aspect ratio', () => {
  assert.equal(nearestAspectRatio(1), '1:1');
  assert.equal(nearestAspectRatio(2), '16:9');
  assert.equal(nearestAspectRatio(0.5), '9:16');
  assert.equal(nearestAspectRatio(1.3), '4:3');
  assert.equal(nearestAspectRatio(7), '8:1');
});

test('gemini() limits every answer, and asks once more when one is cut off', async () => {
  const good = { status: 'completed', steps: [{ type: 'model_output', content: [{ type: 'text', text: '{"ok":true}' }] }] };
  const cutOff = { status: 'incomplete', steps: [{ type: 'model_output', content: [{ type: 'text', text: '{"purpose":"and so on. and so on. and' }] }] };
  const sequence = (answers: unknown[]) => {
    const limits: unknown[] = [];
    const fetch = (async (_url: string | URL | Request, init?: RequestInit) => {
      limits.push((JSON.parse(String(init?.body)) as { generation_config?: { max_output_tokens?: number } }).generation_config?.max_output_tokens);
      return new Response(JSON.stringify(answers[Math.min(limits.length, answers.length) - 1]));
    }) as typeof globalThis.fetch;
    return { fetch, limits };
  };
  const request = { system: 's', prompt: 'p', schema: {} };

  // The caller's estimate of the answer, plus room for reasoning.
  const estimated = sequence([good]);
  await gemini({ model: 'm', fetch: estimated.fetch }).generateJson({ ...request, maxOutputTokens: 1360 });
  assert.deepEqual(estimated.limits, [1360 + 8192]);

  // A limit set on the adapter wins.
  const fixed = sequence([good]);
  await gemini({ model: 'm', fetch: fixed.fetch, maxOutputTokens: 5000 }).generateJson({ ...request, maxOutputTokens: 1360 });
  assert.deepEqual(fixed.limits, [5000]);

  // A model that ran away is stopped at the limit; the same request usually works the second time.
  const once = sequence([cutOff, good]);
  assert.deepEqual(await gemini({ model: 'm', fetch: once.fetch }).generateJson(request), { ok: true });
  assert.equal(once.limits.length, 2);

  // Twice in a row is reported, with what to do about it.
  const twice = sequence([cutOff, cutOff, good]);
  await assert.rejects(gemini({ model: 'm', fetch: twice.fetch }).generateJson(request), /cut off at the limit of 24576 tokens, twice/);
  assert.equal(twice.limits.length, 2, 'it is not asked a third time');
});

test('a text-only deck is offered the plain text patterns and nothing else', async () => {
  const template = loadTemplate(pptx, cardsManifest);
  const requests: JsonRequest[] = [];
  const llm: LlmProvider = {
    async generateJson(request) {
      requests.push(request);
      return { title: 'T', slides: [{ pattern: 'closing', purpose: 'x' }] };
    },
  };
  const patternsOffered = (request: JsonRequest | undefined) =>
    (request?.schema as { properties: { slides: { items: { properties: { pattern: { enum: string[] } } } } } }).properties.slides.items.properties.pattern.enum;
  const icons = [{ id: 'star', description: 'a star', data: new Uint8Array([1]) }];
  const fields = { closing: 'Thanks' };

  // Without the wish every usable pattern is on offer, and the plain ones are named as such.
  await planOutline(template, { prompt: 'Dogs', slideCount: 1, fields, imageDensity: 'none' }, { llm, icons, layout: 'adaptive' });
  assert.deepEqual(patternsOffered(requests[0]), ['closing', 'cards']);
  assert.match(requests[0]?.prompt ?? '', /- closing: Closing words\. Plain text\.\n/);
  assert.match(requests[0]?.prompt ?? '', /- cards: A row of cards\.\n/);
  assert.match(requests[0]?.prompt ?? '', /A wish about the kind of slides \(text only, no tables, only cards\) decides which patterns may be used, and comes before variety/);
  assert.match(requests[0]?.prompt ?? '', /- Otherwise vary the patterns across the deck\./);

  // With it, cards are not offered at all, so the wish cannot be missed; and repeating a pattern is said to be fine.
  const plan = await planOutline(template, { prompt: 'Dogs', slideCount: 1, fields, textOnly: true }, { llm: { generateJson: async (request: JsonRequest) => (requests.push(request), { title: 'T', slides: [{ pattern: 'closing', purpose: 'x' }] }) }, icons, layout: 'adaptive' });
  assert.deepEqual(patternsOffered(requests[1]), ['closing']);
  assert.match(requests[1]?.prompt ?? '', /The deck is text only, and every pattern in the list is plain text\. Using the same pattern for several slides in a row is expected\./);
  assert.doesNotMatch(requests[1]?.prompt ?? '', /vary the patterns/i);
  assert.equal(plan.brief.textOnly, true, 'the wish stays with the plan');
  // The answer is given room for a plan of this size and no more.
  assert.equal(requests[1]?.maxOutputTokens, 400 + 160);

  // A template without a plain text pattern says so.
  const cardsOnly = loadTemplate(pptx, { ...cardsManifest, patterns: cardsManifest.patterns.filter((pattern) => pattern.id === 'cards') });
  await assert.rejects(planOutline(cardsOnly, { prompt: 'Dogs', textOnly: true }, { llm, icons }), /has no plain text pattern, so a text-only deck cannot be planned/);
});

test('gemini() asks once more after a failure that passes by itself, and not after one that will not', async () => {
  const good = { status: 'completed', steps: [{ type: 'model_output', content: [{ type: 'text', text: '{"ok":true}' }] }] };
  const sequence = (answers: Array<[status: number, body: unknown]>) => {
    let calls = 0;
    const fetch = (async () => {
      const [status, body] = answers[Math.min(calls, answers.length - 1)] as [number, unknown];
      calls++;
      return new Response(JSON.stringify(body), { status });
    }) as typeof globalThis.fetch;
    return { fetch, calls: () => calls };
  };
  const request = { system: 's', prompt: 'p', schema: {} };

  // The service's recitation filter trips at random on ordinary answers.
  const recitation = sequence([[400, { error: { message: 'Request blocked due to copyright/recitation content. Please modify your input and retry.' } }], [200, good]]);
  assert.deepEqual(await gemini({ model: 'm', fetch: recitation.fetch }).generateJson(request), { ok: true });
  assert.equal(recitation.calls(), 2);

  const overloaded = sequence([[503, { error: { message: 'The model is overloaded' } }], [200, good]]);
  assert.deepEqual(await gemini({ model: 'm', fetch: overloaded.fetch }).generateJson(request), { ok: true });

  // A second failure in a row is reported.
  const down = sequence([[503, { error: { message: 'The model is overloaded' } }]]);
  await assert.rejects(gemini({ model: 'm', fetch: down.fetch }).generateJson(request), /overloaded/);
  assert.equal(down.calls(), 2);

  // A bad key, a spending cap or a bad request will fail the same way again, so it is not repeated.
  for (const [status, message] of [[401, 'API key not valid'], [429, 'Your project has exceeded its monthly spending cap'], [400, 'Schema is too complex']] as const) {
    const failing = sequence([[status, { error: { message } }], [200, good]]);
    await assert.rejects(gemini({ model: 'm', fetch: failing.fetch }).generateJson(request), new RegExp(message));
    assert.equal(failing.calls(), 1, message);
  }
});
