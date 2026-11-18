// Design problems found by measurement alone, and the second look at a draft that sends them
// (or what a model sees in pictures of the slides) back to the writing model.
import assert from 'node:assert/strict';
import { readFileSync } from 'node:fs';
import { dirname, join } from 'node:path';
import { test } from 'node:test';
import { fileURLToPath } from 'node:url';
import * as fontkit from 'fontkit';
import { placeholderPhoto } from '../examples/placeholder.ts';
import {
  PptxPackage,
  TOOL_NAMES,
  checkDesign,
  designReview,
  fontkitMetrics,
  loadTemplate,
  memoryStore,
  planDeck,
  presentationTools,
  textMeasurer,
  validateDeck,
  visionReview,
} from '../src/index.ts';
import type { DeckSpec, FontFiles, FontMetrics, Issue, JsonRequest, PreviewProvider, TemplateManifest } from '../src/index.ts';
import { readPresentation } from '../src/pptx/presentation.ts';
import { officeFonts, openFonts } from './fonts.ts';
import { withoutCodes } from './issues.ts';

const root = join(dirname(fileURLToPath(import.meta.url)), '..');
const pptx = readFileSync(join(root, 'templates', 'demo.pptx'));
const manifest = JSON.parse(readFileSync(join(root, 'templates', 'demo.manifest.json'), 'utf8')) as TemplateManifest;

/** Every character half an em wide, so that lines can be counted by hand. */
const fixedWidth: FontMetrics = { face: () => ({ has: () => true, advances: (text) => [...text].map(() => 0.5) }) };
const template = loadTemplate(pptx, manifest, { fonts: fixedWidth });
const measure = textMeasurer(template);
const words = (count: number) => Array.from({ length: count }, () => 'word').join(' ');
const png = placeholderPhoto(20, 10, [0, 0, 0], [255, 255, 255]);

const deckOf = (...slides: DeckSpec['slides']): DeckSpec => ({ version: 1, fields: { organizationDepartment: 'Name' }, slides });
const textOnly = (slots: Record<string, string | string[]>): DeckSpec['slides'][number] => ({ pattern: 'text-only', slots: { title: 'A title', subtitle: 'A subtitle', body: words(60), ...slots } });
const photoSlide = (slots: Record<string, string | string[]>): DeckSpec['slides'][number] => ({
  pattern: 'text-and-photo',
  slots: { title: 'A title', subtitle: 'A subtitle', body: [words(30), words(14)], photo: { asset: 'photo' }, ...slots },
});
const kindsOf = (issues: Array<{ slot?: string; kind: string }>) => issues.map((issue) => `${issue.slot}:${issue.kind}`);

// --- the checks -------------------------------------------------------------------------------

test('a deck whose text sits as the template intends has no design problems', () => {
  assert.deepEqual(checkDesign(template, deckOf(textOnly({}), photoSlide({}))), []);
  // The checks measure, so they need the template's fonts.
  assert.throws(() => checkDesign(loadTemplate(pptx, manifest), deckOf(textOnly({}))), /needs font metrics/);
});

test('text that runs over its box is reported as validation reports it', () => {
  const { capacity = 0 } = measure.fit('text-only', 'body', 'a') ?? {};
  const perLine = measure.fit('text-only', 'body', 'a'.repeat(400))?.paragraphs[0]?.[0]?.length ?? 0;
  const deck = deckOf(textOnly({ body: 'a'.repeat(perLine * (capacity + 2)) }));
  const [issue, ...others] = checkDesign(template, deck);
  assert.deepEqual(others, []);
  assert.deepEqual([issue?.severity, issue?.slide, issue?.slot, issue?.kind], ['warning', 1, 'body', 'overflow']);
  assert.equal(issue?.message, validateDeck(manifest, deck, { measure })[0]?.message);
  assert.match(issue?.message ?? '', /^text takes \d+ lines at \d+ pt; the box holds \d+ \(2 lines too many\)/);
});

test('a box the template fills with several lines is reported when it is left almost empty', () => {
  // The body of "text-and-photo" holds five lines of the template's own text in the made-up font.
  const sample = measure.fit('text-and-photo', 'body', 'x')?.sampleLines ?? 0;
  assert.equal(sample, 5);
  const [issue, ...others] = checkDesign(template, deckOf(photoSlide({ body: 'Two words' })));
  assert.deepEqual(others, []);
  assert.deepEqual([issue?.slide, issue?.slot, issue?.kind], [1, 'body', 'empty']);
  assert.equal(issue?.message, "text takes 1 line where the template's own text takes 5, so its box will look empty. Text of about 45 characters fills it as designed.");

  // Half as full as the template draws it is not empty; and the share can be set.
  const half = deckOf(photoSlide({ body: words(30) }));
  assert.deepEqual(checkDesign(template, half), []);
  assert.deepEqual(kindsOf(checkDesign(template, half, { emptyBelow: 0.75 })), ['body:empty']);

  // "text-only" is derived from the same slide with its body moved and widened, so the template has
  // no text of its own to compare with: the box is taken to be meant half full.
  const { capacity = 0, sampleLines } = measure.fit('text-only', 'body', 'x') ?? {};
  assert.deepEqual([sampleLines, capacity >= 6], [0, true]);
  const derived = checkDesign(template, deckOf(textOnly({ body: 'Two words' })));
  assert.deepEqual(kindsOf(derived), ['body:empty']);
  assert.equal(derived[0]?.message, `text takes 1 line in a box that holds ${capacity}, so its box will look empty. Text of about ${9 * Math.ceil(capacity / 2)} characters fills it as designed.`);
  // A one-line box is never empty, and the header that every slide shares is the app's text.
  assert.deepEqual(checkDesign(template, deckOf(textOnly({ title: 'A', subtitle: 'B' }))), []);
});

test('a card with a fraction of its neighbour\'s text makes the row uneven', () => {
  const cards = (texts: string[]): DeckSpec['slides'][number] => ({
    pattern: 'icon-cards',
    items: texts.length,
    slots: {
      title: 'A title',
      ...Object.fromEntries(texts.flatMap((text, index) => [[`card${index + 1}Heading`, 'Heading'], [`card${index + 1}Text`, text], [`card${index + 1}Icon`, { asset: 'icon' }]])),
    },
  });
  // The cards' text is 15 pt in a box 162 pt wide when the row holds the four cards it draws:
  // four words to a line in the made-up font. The demo's row widens its cards when it holds
  // fewer, so with three the box is 237.6 pt wide and holds six words to a line.
  const issues = checkDesign(template, deckOf(cards([words(20), 'word', words(16)])), { kinds: ['uneven'] });
  assert.deepEqual(kindsOf(issues), ['card2Text:uneven']);
  assert.equal(issues[0]?.message, 'text takes 1 line where "card1Text" beside it takes 4, so the row will look uneven. Text of about 66 characters or more evens it out.');
  // Cards that are all short, or all long, are even.
  assert.deepEqual(checkDesign(template, deckOf(cards(['word', 'two words', 'word'])), { kinds: ['uneven'] }), []);
  assert.deepEqual(checkDesign(template, deckOf(cards([words(20), words(12), words(16)])), { kinds: ['uneven'] }), []);

  // A deck that reads the other way fills the row mirror-wise; the problem is named as the deck names the slot.
  const mirrored = checkDesign(template, { ...deckOf(cards([words(20), 'word', words(16)])), direction: 'rtl' }, { kinds: ['uneven'] });
  assert.deepEqual(kindsOf(mirrored), ['card2Text:uneven']);
});

test('a heading that ends with one word on a line of its own is reported, with the length that avoids it', () => {
  // The title of "text-and-photo" is 34 pt in a box 460.8 pt wide: 27 characters to a line.
  const slide = (title: string): DeckSpec['slides'][number] => ({ pattern: 'text-and-photo', slots: { title, subtitle: 'A subtitle', body: [words(30), words(14)], photo: { asset: 'photo' } } });
  const [issue, ...others] = checkDesign(template, deckOf(slide('aaaaaaaaaaaa bbbbbbbbbbbb cc')));
  assert.deepEqual(others, []);
  assert.deepEqual([issue?.slot, issue?.kind], ['title', 'orphan']);
  assert.equal(issue?.message, 'the last of its 2 lines holds one word, "cc". Text of about 25 characters ends a line sooner; a few words more fill the line.');

  // Two words on the last line, a last line that is long, and a title on one line are all fine.
  assert.deepEqual(checkDesign(template, deckOf(slide('aaaaaaaaaaaa bbbbbbbbbbbb cc dd'))), []);
  assert.deepEqual(checkDesign(template, deckOf(slide('aaaaaaaaaaaa bbbbbbbbbbbb cccccccccccccc'))), []);
  assert.deepEqual(checkDesign(template, deckOf(slide('aaaaaaaaaaaa bbbbbbbbbbbb'))), []);
  // An address is meant to stand on its own line.
  assert.deepEqual(checkDesign(template, deckOf(slide('aaaaaaaaaaaa bbbbbbbbbbbb a@b.org'))), []);
  // Body text is not a heading: its last line may hold one word.
  assert.deepEqual(checkDesign(template, deckOf(textOnly({ body: `${words(40)} a` }))), []);
  // From which size a paragraph is a heading can be set.
  assert.deepEqual(checkDesign(template, deckOf(slide('aaaaaaaaaaaa bbbbbbbbbbbb cc')), { headingSize: 40 }), []);
});

test('the checks agree with PowerPoint on the example deck, measured with real fonts', () => {
  // In PowerPoint's own picture of this deck the cover title ends with "Share" on a line of its
  // own, the title of slide 5 with "trip", and the title of slide 6 (the mirrored figure slide)
  // with "week". Nothing else in the deck is a problem.
  const deck = JSON.parse(readFileSync(join(root, 'examples', 'decks', 'patterns.en.deck.json'), 'utf8')) as DeckSpec;
  const found = (files: FontFiles) =>
    checkDesign(loadTemplate(pptx, manifest, { fonts: fontkitMetrics(fontkit, files) }), deck, { layout: 'adaptive' }).map((issue) => `${issue.slide} ${issue.slot} ${issue.kind} ${/"(\w+)"/.exec(issue.message)?.[1]}`);
  const office = officeFonts();
  const seen = ['1 title orphan Share', '5 title orphan trip', '6 title orphan week'];
  if (office) assert.deepEqual(found(office), seen);
  // The open stand-ins are not the template's fonts to the last unit: they find the same three,
  // and may break one more title a word differently.
  const open = found(openFonts());
  assert.ok(seen.every((entry) => open.includes(entry)), open.join('; '));
  assert.ok(open.length <= seen.length + 1 && open.every((entry) => / orphan /.test(entry)), open.join('; '));
});

// --- a second look in the planner -------------------------------------------------------------

function scripted(...answers: unknown[]) {
  const requests: JsonRequest[] = [];
  return {
    requests,
    llm: {
      async generateJson(request: JsonRequest) {
        requests.push(request);
        if (answers.length === 0) throw new Error('the scripted model has no answer left');
        return answers.shift();
      },
    },
  };
}
const outline = { title: 'Deck', slides: [{ pattern: 'text-only', purpose: 'One point' }] };
const brief = { prompt: 'One point', slideCount: 1, fields: { organizationDepartment: 'Name' } };

test('nothing reviews a draft unless asked; with the design review, its findings go back to the model', async () => {
  const thin = { s1: { title: 'A title', subtitle: 'A subtitle', body: ['Two words'] } };
  const full = { s1: { title: 'A title', subtitle: 'A subtitle', body: [words(60)] } };

  // By default a draft within its limits is accepted as it is.
  const plain = scripted(outline, thin);
  const accepted = await planDeck(template, brief, { llm: plain.llm, patterns: ['text-only'] });
  assert.equal(plain.requests.length, 2);
  assert.deepEqual(accepted.warnings, []);

  const reviewed = scripted(outline, thin, full);
  const progress: string[] = [];
  const result = await planDeck(template, brief, { llm: reviewed.llm, patterns: ['text-only'], review: designReview(template), onProgress: (event) => progress.push(event.stage) });
  assert.equal(reviewed.requests.length, 3);
  assert.deepEqual(progress, ['outline', 'write', 'repair']);
  assert.match(reviewed.requests[2]?.prompt ?? '', /It has these problems:\n- s1\.body: text takes 1 line in a box that holds \d+, so its box will look empty\. Text of about \d+ characters fills it as designed\./);
  assert.deepEqual(result.warnings, []);
  assert.deepEqual(result.deck.slides[0]?.slots.body, [words(60)]);

  // What the model does not fix is returned with the warnings.
  const stubborn = scripted(outline, thin, thin, thin);
  const left = await planDeck(template, brief, { llm: stubborn.llm, patterns: ['text-only'], review: designReview(template) });
  assert.equal(stubborn.requests.length, 4, 'the outline, the draft and two repairs');
  assert.deepEqual(left.warnings.map((issue) => [issue.slide, issue.slot, (issue as { kind?: string }).kind]), [[1, 'body', 'empty']]);
});

test('a review of one\'s own is a function from a deck to problems, and gets the caller\'s signal', async () => {
  const seen: Array<{ slides: number; signal: boolean }> = [];
  const model = scripted(outline, { s1: { title: 'A title', subtitle: 'A subtitle', body: [words(60)] } }, { s1: { title: 'Better', subtitle: 'A subtitle', body: [words(60)] } });
  const stop = new AbortController();
  const result = await planDeck(template, brief, {
    llm: model.llm,
    patterns: ['text-only'],
    signal: stop.signal,
    review: (deck, { signal }) => {
      seen.push({ slides: deck.slides.length, signal: signal === stop.signal });
      return deck.slides[0]?.slots.title === 'A title' ? [{ severity: 'warning', slide: 1, slot: 'title', message: 'the title says nothing' }] : [];
    },
  });
  assert.deepEqual(seen, [{ slides: 1, signal: true }, { slides: 1, signal: true }]);
  assert.match(model.requests[2]?.prompt ?? '', /- s1\.title: the title says nothing/);
  assert.equal(result.deck.slides[0]?.slots.title, 'Better');
});

// --- a model that looks at pictures -----------------------------------------------------------

/** A provider that gives a small PNG for each slide and remembers the files it was given. */
function pictures(mimeType = 'image/png'): PreviewProvider & { decks: number[] } {
  const decks: number[] = [];
  return {
    decks,
    async renderSlides({ pptx: file }) {
      const count = readPresentation(PptxPackage.load(file)).slides.length;
      decks.push(count);
      return Array.from({ length: count }, (_, index) => ({ slide: index + 1, mimeType, data: png }));
    },
  };
}

const threeSlides = deckOf(
  { pattern: 'text-and-photo', slots: { title: 'First', subtitle: 'A subtitle', body: [words(30)], photo: { asset: 'gen-photo', prompt: 'A photo not made yet' } } },
  textOnly({ title: 'Second' }),
  textOnly({ title: 'Third' }),
);

test('the vision review shows the model a picture of each slide with its text, and turns what it reports into problems', async () => {
  const preview = pictures();
  const model = scripted({
    problems: [
      { slide: 2, slot: 'body', problem: 'The text runs under the bottom edge; shorten it by a third.' },
      { slide: 1, slot: 'no-such-slot', problem: '  The slide is  mostly empty. ' },
      { slide: 9, slot: 'title', problem: 'A slide that was not shown.' },
      { slide: 3, slot: 'title', problem: '' },
    ],
  });
  const review = visionReview({ template, llm: model.llm, preview });
  const issues = await review(threeSlides, {});

  // The deck was rendered with a stand-in for the photo that does not exist yet, and drawn once.
  assert.deepEqual(preview.decks, [3]);
  assert.equal(model.requests.length, 1);
  const [request] = model.requests;
  assert.deepEqual(request?.images?.map((image) => [image.mimeType, image.data]), [['image/png', png], ['image/png', png], ['image/png', png]]);
  assert.match(request?.prompt ?? '', /The pictures show slides s1, s2, s3 of a deck, in that order/);
  assert.match(request?.prompt ?? '', /s1 \(text-and-photo\):\n- title: "First"\n- subtitle: "A subtitle"\n- body: "word word/);
  assert.match(request?.prompt ?? '', /Photos may be plain coloured rectangles/);
  assert.equal(request?.maxOutputTokens, 900, 'the answer is capped: a sentence for each problem');
  assert.deepEqual((request?.schema as { properties: { problems: { items: { properties: { slide: { enum: number[] } } } } } }).properties.problems.items.properties.slide.enum, [1, 2, 3]);

  // A slot the slide does not have is dropped from the problem; a slide not shown, and an empty report, are dropped whole.
  assert.deepEqual(withoutCodes(issues), [
    { severity: 'warning', slide: 2, slot: 'body', message: 'seen in the picture of the slide: The text runs under the bottom edge; shorten it by a third.' },
    { severity: 'warning', slide: 1, message: 'seen in the picture of the slide: The slide is mostly empty.' },
  ] satisfies Issue[]);
});

test('the vision review asks in batches, needs bitmap pictures, and leaves a draft with errors alone', async () => {
  const batches = scripted({ problems: [] }, { problems: [{ slide: 3, slot: 'title', problem: 'One word on the last line.' }] });
  const issues = await visionReview({ template, llm: batches.llm, preview: pictures(), slidesPerRequest: 2 })(threeSlides, {});
  assert.deepEqual(batches.requests.map((request) => [request.images?.length, request.maxOutputTokens]), [[2, 700], [1, 500]]);
  assert.match(batches.requests[1]?.prompt ?? '', /The pictures show slides s3 of a deck/);
  assert.deepEqual(issues.map((issue) => [issue.slide, issue.slot]), [[3, 'title']]);

  // A model cannot read SVG, which is what the Pagus adapter gives.
  const unread = scripted();
  await assert.rejects(async () => visionReview({ template, llm: unread.llm, preview: pictures('image/svg+xml') })(threeSlides, {}), /needs bitmap pictures such as PNG, and the preview provider gave image\/svg\+xml/);
  assert.equal(unread.requests.length, 0);

  // A draft that cannot be rendered is validation's business: nothing is drawn and nobody is asked.
  const preview = pictures();
  const broken = deckOf({ pattern: 'text-only', slots: { title: 'No body' } });
  assert.deepEqual(await visionReview({ template, llm: unread.llm, preview, layout: 'strict' })(broken, {}), []);
  assert.deepEqual([preview.decks.length, unread.requests.length], [0, 0]);
});

test('the vision review takes part in the repair round, and the chat tools pass a review on', async () => {
  const model = scripted(
    outline,
    { s1: { title: 'A title', subtitle: 'A subtitle', body: [words(60)] } },
    { problems: [{ slide: 1, slot: 'title', problem: 'The title is cut off at the right; make it shorter.' }] },
    { s1: { title: 'Short', subtitle: 'A subtitle', body: [words(60)] } },
    { problems: [] },
  );
  const result = await planDeck(template, brief, { llm: model.llm, patterns: ['text-only'], review: visionReview({ template, llm: model.llm, preview: pictures() }) });
  assert.equal(model.requests.length, 5, 'outline, draft, a look, repair, another look');
  assert.match(model.requests[3]?.prompt ?? '', /- s1\.title: seen in the picture of the slide: The title is cut off at the right; make it shorter\./);
  assert.deepEqual([result.deck.slides[0]?.slots.title, result.warnings], ['Short', []]);

  let looked = 0;
  const tools = presentationTools({
    template,
    llm: scripted(outline, { s1: { title: 'A title', subtitle: 'A subtitle', body: [words(60)] } }).llm,
    store: memoryStore(),
    fields: { organizationDepartment: 'Name' },
    newId: (kind) => `${kind}_1`,
    review: () => (looked++, []),
  });
  await tools.call(TOOL_NAMES.plan, { topic: 'One point', slide_count: 1 });
  const created = await tools.call(TOOL_NAMES.create, { plan_id: 'plan_1' });
  assert.equal(created.output.error, undefined);
  assert.equal(looked, 1);
});
