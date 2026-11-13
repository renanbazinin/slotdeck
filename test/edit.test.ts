import assert from 'node:assert/strict';
import { test } from 'node:test';
import { placeholderPhoto } from '../examples/placeholder.ts';
import { PlanError, PptxPackage, editDeck, extractAssets, loadTemplate, planEdit, renderDeck } from '../src/index.ts';
import type { DeckSpec, ImageRef, JsonRequest, LlmProvider, TemplateManifest } from '../src/index.ts';
import { readPresentation } from '../src/pptx/presentation.ts';
import { buildFixturePptx, fixtureManifest } from './fixture.ts';

const pptx = buildFixturePptx();
const template = loadTemplate(pptx, fixtureManifest);
const png = placeholderPhoto(20, 10, [0, 0, 0], [255, 255, 255]);

const photo: ImageRef = { asset: 'gen-0000000000000001', prompt: 'A dog in a park', alt: 'Dog' };
const base: DeckSpec = {
  version: 1,
  title: 'Dogs',
  brief: 'A deck about dogs',
  language: 'en',
  direction: 'ltr',
  fields: { closing: 'Thank you' },
  slides: [
    { pattern: 'content', slots: { title: 'Dogs are great', body: ['Loyal', 'Playful'], photo, note: 'A note' } },
    { pattern: 'closing', slots: {} },
  ],
};

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

type Schema = { properties: Record<string, Schema>; required?: string[]; enum?: string[]; items?: Schema };
const schemaOf = (request: JsonRequest | undefined) => request?.schema as unknown as Schema;

const plan = (slides: Array<[from: number, pattern: string, change?: string, images?: boolean]>, extra: Record<string, unknown> = {}) => ({
  summary: 'Done.',
  language: 'en',
  title: 'Dogs',
  slides: slides.map(([from, pattern, change = '', images = false]) => ({ from, pattern, change, images })),
  ...extra,
});

test('rewrites only the slide the instruction is about, and keeps its picture', async () => {
  const { llm, requests } = scripted(
    plan([
      [1, 'content', 'Shorten the body to one point'],
      [2, 'closing'],
    ]),
    { s1: { title: 'Dogs are great', body: ['Loyal'], note: 'A note' } },
  );
  const result = await planEdit(template, base, 'Shorten slide 1', { llm });

  assert.equal(requests.length, 2);
  assert.deepEqual(result.rewritten, [1]);
  assert.equal(result.summary, 'Done.');
  assert.deepEqual(result.deck.slides[0]?.slots.body, ['Loyal']);
  assert.deepEqual(result.deck.slides[0]?.slots.photo, photo, 'the photo reference is carried over untouched');
  assert.deepEqual(result.deck.slides[1], base.slides[1], 'the other slide is copied as it was');
  assert.equal(result.deck.brief, 'A deck about dogs\nShorten slide 1', 'the request joins the brief');
  assert.deepEqual(result.deck.fields, base.fields);

  // The model sees the whole deck when planning, and only the changing slide when writing.
  assert.match(requests[0]?.prompt ?? '', /1\. \[content\] title: "Dogs are great"/);
  assert.match(requests[0]?.prompt ?? '', /picture of "A dog in a park"/);
  assert.match(requests[0]?.prompt ?? '', /Requested change:\nShorten slide 1/);
  const schema = schemaOf(requests[1]);
  assert.deepEqual(Object.keys(schema.properties), ['s1']);
  assert.deepEqual(Object.keys((schema.properties.s1 as Schema).properties), ['title', 'body', 'note'], 'no photo: it is not being changed');
  assert.match(requests[1]?.prompt ?? '', /s1 \(content\)\. Change: Shorten the body to one point/);
  assert.match(requests[1]?.prompt ?? '', /Current content: \{"title":"Dogs are great","body":\["Loyal","Playful"\],"note":"A note"\}/);
  assert.match(requests[1]?.prompt ?? '', /2\. closing, unchanged/);
});

test('offers the picture to the model only when the instruction is about it', async () => {
  const { llm, requests } = scripted(
    plan([
      [1, 'content', 'Show a puppy instead', true],
      [2, 'closing'],
    ]),
    { s1: { title: 'Dogs are great', body: ['Loyal', 'Playful'], note: 'A note', photo: { prompt: 'A puppy on a sofa', alt: 'Puppy' } } },
  );
  const result = await planEdit(template, base, 'Change the photo on slide 1 to a puppy', { llm });

  assert.ok('photo' in (schemaOf(requests[1]).properties.s1 as Schema).properties);
  const next = result.deck.slides[0]?.slots.photo as ImageRef;
  assert.equal(next.prompt, 'A puppy on a sofa');
  assert.notEqual(next.asset, photo.asset, 'a new description is a new asset');
});

test('adds, removes and reorders slides', async () => {
  const added = scripted(
    plan([
      [2, 'closing'],
      [0, 'content', 'A new slide about cats', true],
    ]),
    { s2: { title: 'Cats too', body: ['Quiet'], note: 'n', photo: { prompt: 'A cat', alt: 'Cat' } } },
  );
  const result = await planEdit(template, base, 'Drop slide 1, and add a slide about cats at the end', { llm: added.llm });
  assert.deepEqual(
    result.deck.slides.map((slide) => slide.pattern),
    ['closing', 'content'],
  );
  assert.deepEqual(result.rewritten, [2]);
  assert.equal(result.deck.slides[1]?.slots.title, 'Cats too');
  assert.deepEqual(Object.keys(schemaOf(added.requests[1]).properties), ['s2']);
  assert.match(added.requests[1]?.prompt ?? '', /s2 \(content\)\. New slide: A new slide about cats/);

  // Moving and deleting need no second call at all.
  const moved = scripted(
    plan([
      [2, 'closing'],
      [1, 'content'],
    ]),
  );
  const reordered = await planEdit(template, base, 'Put the closing slide first', { llm: moved.llm });
  assert.equal(moved.requests.length, 1);
  assert.deepEqual(reordered.rewritten, []);
  assert.deepEqual(reordered.deck.slides, [base.slides[1], base.slides[0]]);
});

test('a change of language rewrites every slide and sets the direction', async () => {
  const { llm, requests } = scripted(
    plan(
      [
        [1, 'content'],
        [2, 'closing'],
      ],
      { language: 'he' },
    ),
    { s1: { title: 'כלבים', body: ['נאמנים'], note: 'הערה' }, s2: {} },
  );
  const result = await planEdit(template, base, 'Translate the deck to Hebrew', { llm });

  assert.equal(result.deck.language, 'he');
  assert.equal(result.deck.direction, 'rtl');
  assert.deepEqual(result.rewritten, [1, 2]);
  assert.match(requests[1]?.system ?? '', /BCP 47 tag "he"/);
  assert.match(requests[1]?.prompt ?? '', /Translate into the new deck language/);
  assert.deepEqual(result.deck.slides[0]?.slots.photo, photo, 'pictures survive a translation');
});

test('deck-level texts can be changed by an edit, unless the caller fixes them', async () => {
  const translated = plan(
    [
      [1, 'content'],
      [2, 'closing'],
    ],
    { language: 'he', fields: { closing: 'תודה' } },
  );
  const content = { s1: { title: 'כלבים', body: ['נאמנים'], note: 'הערה' }, s2: {} };

  const { llm, requests } = scripted(translated, content);
  const result = await planEdit(template, base, 'Translate the deck to Hebrew', { llm });
  assert.deepEqual(result.deck.fields, { closing: 'תודה' });
  assert.match(requests[0]?.prompt ?? '', /Deck-level texts, shown on several slides:\nclosing: "Thank you"/);
  assert.deepEqual((schemaOf(requests[0]).properties.fields as Schema).required, ['closing']);

  // A field the model returns empty keeps its value.
  const blank = await planEdit(template, base, 'Translate', { llm: scripted({ ...translated, fields: { closing: '' } }, content).llm });
  assert.deepEqual(blank.deck.fields, { closing: 'Thank you' });

  // Values the caller passes are not offered to the model at all.
  const fixed = scripted(translated, content);
  const pinned = await planEdit(template, base, 'Translate', { llm: fixed.llm, fields: { closing: 'Merci' } });
  assert.deepEqual(pinned.deck.fields, { closing: 'Merci' });
  assert.ok(!('fields' in schemaOf(fixed.requests[0]).properties));
});

test('a slide given a new layout keeps its photo', async () => {
  const twoLayouts: TemplateManifest = {
    ...fixtureManifest,
    patterns: [
      ...fixtureManifest.patterns,
      {
        id: 'photo-only',
        sourceSlide: 1,
        description: 'A title over a photo.',
        slots: [
          { name: 'heading', kind: 'text', shapeId: '2' },
          { name: 'picture', kind: 'image', imageKind: 'photo', shapeId: '4' },
        ],
      },
    ],
  };
  const { llm, requests } = scripted(
    plan([
      [1, 'photo-only', 'Turn this into a photo slide'],
      [2, 'closing'],
    ]),
    { s1: { heading: 'Dogs' } },
  );
  const result = await planEdit(loadTemplate(pptx, twoLayouts), base, 'Make slide 1 a big photo slide', { llm });

  assert.deepEqual(Object.keys((schemaOf(requests[1]).properties.s1 as Schema).properties), ['heading']);
  assert.deepEqual(result.deck.slides[0], { pattern: 'photo-only', slots: { picture: photo, heading: 'Dogs' } });
  assert.match(requests[1]?.prompt ?? '', /Content of the slide it replaces: \{"title":"Dogs are great"/);
});

test('number patterns: an existing slide may be revised, a new one needs figures', async () => {
  const numbers: TemplateManifest = {
    ...fixtureManifest,
    patterns: fixtureManifest.patterns.map((pattern) => (pattern.id === 'content' ? { ...pattern, figures: true } : pattern)),
  };
  const numberTemplate = loadTemplate(pptx, numbers);
  const revise = plan([
    [1, 'content', 'Shorten'],
    [2, 'closing'],
  ]);
  const add = plan([
    [1, 'content'],
    [0, 'content', 'Another number slide'],
    [2, 'closing'],
  ]);
  const written = { title: 'Short', body: ['One'], note: 'n' };

  const revised = await planEdit(numberTemplate, base, 'Shorten slide 1', { llm: scripted(revise, { s1: written }).llm });
  assert.equal(revised.deck.slides[0]?.slots.title, 'Short');

  await assert.rejects(planEdit(numberTemplate, base, 'Add another slide like the first', { llm: scripted(add, add).llm }), (error: unknown) => {
    assert.ok(error instanceof PlanError);
    assert.match(error.message, /cannot be used for a new or re-laid-out slide/);
    return true;
  });

  const withFigures = scripted(add, { s2: { ...written, photo: { prompt: 'p', alt: 'a' } } });
  const grown = await planEdit(numberTemplate, base, 'Add a slide saying we walk 3 times a day', { llm: withFigures.llm });
  assert.equal(grown.deck.slides.length, 3);
});

test('an unusable edit plan is retried once, then rejected', async () => {
  const bad = plan([[7, 'content', 'x']]);
  const good = plan([
    [1, 'content'],
    [2, 'closing'],
  ]);
  const recovered = scripted(bad, good);
  await planEdit(template, base, 'Do nothing', { llm: recovered.llm });
  assert.match(recovered.requests[1]?.prompt ?? '', /"from" must be 0 or the number of an existing slide \(1 to 2\)/);

  await assert.rejects(planEdit(template, base, 'x', { llm: scripted(bad, bad).llm }), PlanError);
  await assert.rejects(planEdit(template, base, 'x', { llm: scripted(plan([[0, 'content', '']]), plan([[0, 'content', '']])).llm }), /needs "change"/);
});

test('only the rewritten slide is sent back for repair', async () => {
  const tooLong = { s1: { title: 'A title far longer than twenty characters', body: ['Loyal'], note: 'A note' } };
  const fixed = { s1: { title: 'Dogs', body: ['Loyal'], note: 'A note' } };
  const { llm, requests } = scripted(
    plan([
      [1, 'content', 'New title'],
      [2, 'closing'],
    ]),
    tooLong,
    fixed,
  );
  const result = await planEdit(template, base, 'Retitle slide 1', { llm });
  assert.equal(requests.length, 3);
  assert.match(requests[2]?.prompt ?? '', /s1\.title: text is 41 characters; the slot fits about 20/);
  assert.equal(result.deck.slides[0]?.slots.title, 'Dogs');
  assert.deepEqual(result.warnings, []);
});

test('extractAssets reads the pictures of a rendered deck back by asset id', () => {
  const rendered = renderDeck(template, base, { assets: { [photo.asset]: png } }).pptx;
  assert.deepEqual(extractAssets(template, base, rendered), { [photo.asset]: png });

  // A deck that no longer matches the file yields nothing rather than the wrong picture.
  const retitled: DeckSpec = { ...base, slides: [{ pattern: 'content', slots: { ...base.slides[0]?.slots, title: 'Something else' } }, base.slides[1] as DeckSpec['slides'][number]] };
  assert.deepEqual(extractAssets(template, retitled, rendered), {});
  assert.deepEqual(extractAssets(template, { ...base, slides: base.slides.slice(0, 1) }, rendered), {});
});

test('editDeck renders the edited deck without generating pictures it already has', async () => {
  const before = renderDeck(template, base, { assets: { [photo.asset]: png } }).pptx;
  const cache = new Map(Object.entries(extractAssets(template, base, before)));
  const { llm } = scripted(
    plan([
      [1, 'content', 'Shorten the body'],
      [2, 'closing'],
    ]),
    { s1: { title: 'Dogs are great', body: ['Loyal'], note: 'A note' } },
  );

  const result = await editDeck(template, base, 'Shorten slide 1', {
    llm,
    cache,
    images: {
      async generateImage() {
        throw new Error('no picture should be generated for a text edit');
      },
    },
  });

  const pkg = PptxPackage.load(result.pptx);
  assert.equal(readPresentation(pkg).slides.length, 2);
  const slide = pkg.readText('ppt/slides/slide1.xml');
  assert.ok(slide.includes('Loyal') && !slide.includes('Playful'));
  assert.deepEqual(result.assets, { [photo.asset]: png });
  assert.deepEqual(result.rewritten, [1]);
});
