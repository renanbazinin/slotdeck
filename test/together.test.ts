// Layout freedom and measurement together: text is measured in the box the renderer really gives
// it, whether a row widened its cards, a pattern was derived from another, or a slide was written
// again for another pattern. Each test here needs both halves to pass.
import assert from 'node:assert/strict';
import { readFileSync } from 'node:fs';
import { dirname, join } from 'node:path';
import { test } from 'node:test';
import { fileURLToPath } from 'node:url';
import * as fontkit from 'fontkit';
import * as pagusCore from '@pagus-kit/core';
import * as pagusRenderer from '@pagus-kit/renderer';
import {
  PptxPackage,
  changeManifest,
  checkDesign,
  checkManifest,
  designReview,
  fontkitMetrics,
  loadTemplate,
  measuredBudgets,
  pagusPreviews,
  patternPreviews,
  planDeck,
  previewDeck,
  renderDeck,
  sampleDeck,
  slideAlternatives,
  textMeasurer,
  validateDeck,
} from '../src/index.ts';
import type { DeckSpec, FontMetrics, JsonRequest, LlmProvider, PatternDef, PreviewProvider, SlideSpec, SlotValue, Template, TemplateManifest } from '../src/index.ts';
import { readPresentation } from '../src/pptx/presentation.ts';
import { effectiveSlots, itemRange, textBudget, widenFactor } from '../src/template/manifest.ts';
import { TextStyles } from '../src/text/styles.ts';
import { strFromU8, strToU8, unzipSync, zipSync } from 'fflate';
import { buildFixturePptx, cardsManifest } from './fixture.ts';
import { openFonts } from './fonts.ts';

const root = join(dirname(fileURLToPath(import.meta.url)), '..');
const pptx = readFileSync(join(root, 'templates', 'demo.pptx'));
const manifest = JSON.parse(readFileSync(join(root, 'templates', 'demo.manifest.json'), 'utf8')) as TemplateManifest;
const icon = new Uint8Array(readFileSync(join(root, 'examples', 'icons', 'star.png')));

/** Every character half an em wide, so that lines can be counted by hand. */
const fixedWidth: FontMetrics = { face: () => ({ has: () => true, advances: (text) => [...text].map(() => 0.5) }) };
const plain = loadTemplate(pptx, manifest);
const measured = loadTemplate(pptx, manifest, { fonts: fixedWidth });
const real = loadTemplate(pptx, manifest, { fonts: fontkitMetrics(fontkit, openFonts()) });
const measure = textMeasurer(measured);

const words = (count: number) => Array.from({ length: count }, () => 'word').join(' ');
const patternOf = (id: string, from: TemplateManifest = manifest) => from.patterns.find((pattern) => pattern.id === id) as PatternDef;
const fields = { organizationDepartment: 'Name' };
const deckOf = (...slides: SlideSpec[]): DeckSpec => ({ version: 1, fields, slides });

/** A slide of a pattern with a row, at a number of items: every text is `text`, every picture the same icon. */
function rowSlide(pattern: PatternDef, items: number, text: string | ((slot: string) => string) = 'word'): SlideSpec {
  const slots: Record<string, SlotValue> = {};
  for (const slot of effectiveSlots(pattern, items)) {
    if (slot.empty === 'drop' || slot.empty === 'keep') continue;
    if (slot.kind === 'text' && !slot.deckField) slots[slot.name] = typeof text === 'string' ? text : text(slot.name);
    else if (slot.kind === 'image') slots[slot.name] = { asset: 'icon' };
  }
  return { pattern: pattern.id, items, slots };
}

/** The fixture template with a typeface stated on every run: it names none, and text is measured in a typeface. */
function fixtureWithTypeface(): Uint8Array {
  const files = unzipSync(buildFixturePptx());
  for (const [name, data] of Object.entries(files)) {
    if (!/^ppt\/slides\/slide\d+\.xml$/.test(name)) continue;
    files[name] = strToU8(strFromU8(data).replace(/<a:rPr([^>]*)\/>/g, '<a:rPr$1><a:latin typeface="Test"/><a:cs typeface="Test"/></a:rPr>'));
  }
  return zipSync(files);
}

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

// --- rows that widen ---------------------------------------------------------------------------

test('a text is measured in the box the renderer gives it, at every number of items a row takes', () => {
  // The rendered file is the truth: the width of each text box is read back from it and compared
  // with the width the measurer laid the same text out in.
  let widened = 0;
  let checked = 0;
  for (const pattern of manifest.patterns) {
    const range = itemRange(pattern);
    if (!pattern.repeat || !range) continue;
    for (let items = range.min; items <= range.max; items++) {
      const { pptx: file } = renderDeck(plain, deckOf(rowSlide(pattern, items)), { assets: { icon }, layout: 'adaptive' });
      const rendered = new TextStyles(PptxPackage.load(file));
      const drawn = new TextStyles(plain.pkg);
      for (const slot of effectiveSlots(pattern, items)) {
        if (slot.kind !== 'text' || slot.deckField || slot.empty === 'drop' || slot.empty === 'keep') continue;
        const frame = rendered.frame(1, slot.shapeId, slot.styleRun ?? 'first');
        assert.ok(frame, `${pattern.id} with ${items}: ${slot.name} is on the rendered slide`);
        const paragraph = frame.paragraphs[0];
        const room = frame.width - frame.insets.left - frame.insets.right - (paragraph?.marginStart ?? 0) - (paragraph?.marginEnd ?? 0);
        const fit = measure.fit(pattern, slot, 'word', { items });
        assert.ok(fit, `${pattern.id} with ${items}: ${slot.name} is measured`);
        assert.ok(Math.abs((fit.lineRoom[0] ?? 0) - room) < 0.02, `${pattern.id} with ${items}: ${slot.name} is measured ${fit.lineRoom[0]} pt wide and rendered ${room.toFixed(2)} pt wide`);
        checked++;
        const before = drawn.frame(pattern.sourceSlide, slot.shapeId, slot.styleRun ?? 'first');
        if (before && frame.width > before.width + 1) widened++;
      }
    }
  }
  // Title and texts of two to four cards, two to six figures and the two columns.
  assert.equal(checked, 71);
  // Two and three cards, and two and four figures, have wider texts than the template draws.
  assert.equal(widened, 22);
});

test('a card that widens keeps its height and how full it was drawn: only its lines get longer', () => {
  const at = (items?: number) => measure.fit('icon-cards', 'card1Text', words(20), items === undefined ? {} : { items });
  const [drawn, four, three, two] = [at(), at(4), at(3), at(2)];
  assert.ok(drawn && four && three && two);
  // 15 pt text, 7.5 pt a character: four words to a line in a card as drawn, six and ten in wider ones.
  assert.deepEqual([drawn.lineRoom[0], four.lineRoom[0], three.lineRoom[0], two.lineRoom[0]], [162, 162, 237.6, 388.8]);
  assert.deepEqual([drawn.lines, four.lines, three.lines, two.lines], [5, 5, 4, 2]);
  for (const fit of [four, three, two]) {
    assert.equal(fit.room, drawn.room, 'the room under the first line is the card\'s height, which does not change');
    assert.equal(fit.capacity, drawn.capacity);
    assert.equal(fit.sampleLines, drawn.sampleLines, 'how full the designer drew the card is counted in the card as drawn');
  }
  assert.ok(two.fitsChars > three.fitsChars && three.fitsChars > four.fitsChars, 'a wider card holds more characters');

  // A grid is as wide as its fullest line: five figures leave a line of three, four make two lines of two.
  const caption = (items: number) => measure.fit('stat-cards', 'stat1Caption', words(20), { items })?.lineRoom[0];
  assert.deepEqual([caption(6), caption(5), caption(4), caption(3), caption(2)], [230.4, 230.4, 381.6, 230.4, 381.6]);

  // Without the setting, the number of items changes nothing.
  const still: TemplateManifest = { ...manifest, patterns: manifest.patterns.map((pattern) => (pattern.repeat ? { ...pattern, repeat: { ...pattern.repeat, widen: 1 } } : pattern)) };
  const narrow = textMeasurer(loadTemplate(pptx, still, { fonts: fixedWidth }));
  assert.equal(narrow.fit('icon-cards', 'card1Text', words(20), { items: 2 })?.lineRoom[0], 162);
  // Texts outside the row are not touched by it.
  assert.equal(measure.fit('icon-cards', 'title', words(4), { items: 2 })?.lineRoom[0], measure.fit('icon-cards', 'title', words(4))?.lineRoom[0]);
});

test('text that fits two wide cards is accepted there and reported in four narrow ones', () => {
  // Thirty words: eight lines of four in a card as drawn, which holds seven; three lines of ten in one of two cards.
  const cards = patternOf('icon-cards');
  const body = (name: string) => (name.endsWith('Text') ? words(30) : 'Heading');
  const deck = deckOf(rowSlide(cards, 4, body), rowSlide(cards, 2, body));
  const over = (issues: Array<{ slide?: number; slot?: string; message: string }>) => issues.filter((issue) => /too many/.test(issue.message)).map((issue) => `${issue.slide} ${issue.slot}`);

  const expected = ['1 card1Text', '1 card2Text', '1 card3Text', '1 card4Text'];
  assert.deepEqual(over(validateDeck(manifest, deck, { layout: 'adaptive', measure })), expected);
  assert.deepEqual(over(renderDeck(measured, deck, { assets: { icon }, layout: 'adaptive' }).warnings), expected);
  assert.deepEqual(checkDesign(measured, deck, { layout: 'adaptive', kinds: ['overflow'] }).map((issue) => `${issue.slide} ${issue.slot}`), expected);

  // With character budgets alone the same text is over budget on both slides: 149 characters
  // against 117, and against 234 where the two cards are wider.
  const counted = validateDeck(manifest, deck, { layout: 'adaptive' }).map((issue) => `${issue.slide} ${issue.slot}`);
  assert.deepEqual(counted, expected);
});

test('text is shrunk only where the card it sits in is too narrow for it', () => {
  const shrinking: TemplateManifest = { ...manifest, shrinkText: { to: 0.6 } };
  const template = loadTemplate(pptx, shrinking, { fonts: fixedWidth });
  const cards = patternOf('icon-cards');
  const body = (name: string) => (name.endsWith('Text') ? words(30) : 'Heading');
  const { shrunk, warnings, pptx: file } = renderDeck(template, deckOf(rowSlide(cards, 4, body), rowSlide(cards, 2, body)), { assets: { icon }, layout: 'adaptive' });
  assert.deepEqual(warnings.filter((issue) => /too many/.test(issue.message)), []);
  assert.deepEqual(shrunk.map((entry) => `${entry.slide} ${entry.slot}`), ['1 card1Text', '1 card2Text', '1 card3Text', '1 card4Text']);
  // The two wide cards keep the designed size in the file.
  const sizes = (slide: number) => [...PptxPackage.load(file).readText(`ppt/slides/slide${slide}.xml`).matchAll(/ sz="(\d+)"/g)].map((match) => Number(match[1]));
  assert.ok(sizes(1).some((size) => size < 1500), 'the narrow cards are set smaller');
  assert.ok(!sizes(2).some((size) => size < 1500 && size > 0 && size !== 1200 && size !== 1000), `the wide cards keep their sizes: ${sizes(2).join(' ')}`);
});

test('a short text in a wide card looks emptier, and the check says how much text fills it', () => {
  const cards = patternOf('icon-cards');
  const body = (name: string) => (name.endsWith('Text') ? words(8) : 'Heading');
  // Eight words are two lines of four in a card as drawn, where the template's own text takes four: fine.
  assert.deepEqual(checkDesign(measured, deckOf(rowSlide(cards, 4, body)), { layout: 'adaptive', kinds: ['empty'] }), []);
  // In one of two cards they are a single line of the four the card was drawn with.
  const issues = checkDesign(measured, deckOf(rowSlide(cards, 2, body)), { layout: 'adaptive', kinds: ['empty'] });
  assert.deepEqual(issues.map((issue) => issue.slot), ['card1Text', 'card2Text']);
  assert.equal(issues[0]?.message, "text takes 1 line where the template's own text takes 4, so its box will look empty. Text of about 156 characters fills it as designed.");
  // That length is within the budget the wider card has, so the advice can be followed.
  const budget = effectiveSlots(cards, 2).find((slot) => slot.name === 'card1Text');
  assert.ok(budget?.kind === 'text' && (textBudget(budget).maxChars ?? 0) >= 156);
});

test('checkManifest tries the larger budget of a widened row in the wider box', () => {
  // The demo's budgets hold at every number of items, with fonts of the real metrics.
  assert.deepEqual(checkManifest(real), []);
  // The budget a row hands out grows by the factor its items widen by at least, so the box must too.
  let tried = 0;
  for (const pattern of manifest.patterns) {
    const range = itemRange(pattern);
    if (!pattern.repeat || !range) continue;
    for (let items = range.min; items <= range.max; items++) {
      const factor = widenFactor(pattern.repeat, items);
      const names = pattern.repeat.slots;
      const inRow: Set<string> = new Set(Array.from({ length: range.max }, (_, index) => names.map((name) => name.replace('{n}', String(index + 1)))).flat());
      for (const slot of effectiveSlots(pattern, items)) {
        if (slot.kind !== 'text' || slot.maxChars === undefined || factor === 1 || !inRow.has(slot.name)) continue;
        tried++;
        const fit = measure.fit(pattern, slot, 'word', { items });
        const drawn = measure.fit(pattern, slot, 'word');
        assert.ok(fit && drawn && (fit.lineRoom[0] ?? 0) >= (drawn.lineRoom[0] ?? 0) * factor - 0.01, `${pattern.id} ${slot.name} with ${items}: the box widens by at least the budget's factor`);
      }
    }
  }

  assert.equal(tried, 22);

  // A budget that is too generous is reported for the row as drawn and for each number of items that widens it.
  const generous: TemplateManifest = {
    ...manifest,
    patterns: manifest.patterns.map((pattern) =>
      pattern.id !== 'icon-cards' ? pattern : { ...pattern, slots: pattern.slots.map((slot) => (slot.kind === 'text' && /^card\dText$/.test(slot.name) ? { ...slot, maxChars: 260 } : slot)) },
    ),
  };
  const issues = checkManifest(loadTemplate(pptx, generous, { fonts: fixedWidth })).filter((issue) => issue.pattern === 'icon-cards' && issue.slot === 'card1Text');
  assert.deepEqual(
    issues.map((issue) => issue.message.replace(/takes \d+ lines where \d+ fit; about \d+ characters fit/, '...')),
    [
      'has a budget of 260 characters, but sample text of that length ...',
      'has a budget of 520 characters on a slide with 2 of its 4 items, but sample text of that length ...',
      'has a budget of 346 characters on a slide with 3 of its 4 items, but sample text of that length ...',
    ],
  );
  // Without fonts nothing is measured, at any number of items.
  assert.deepEqual(checkManifest(loadTemplate(pptx, generous)), []);
});

test('the sample deck fills a widened card to what the wider box holds', () => {
  const lengths = (template: Template, items: 'min' | undefined) => {
    const { deck } = sampleDeck(template, items ? { fit: 'box', items } : { fit: 'box' });
    const slide = deck.slides.find((entry) => entry.pattern === 'icon-cards');
    return String(slide?.slots.card1Text ?? '').length;
  };
  const four = lengths(measured, undefined);
  const two = lengths(measured, 'min');
  assert.ok(two > four * 2, `two cards are filled with ${two} characters, four with ${four}`);
  // And what it writes fits: the sample of the fewest items has no text over its box.
  const { deck, assets } = sampleDeck(measured, { fit: 'box', items: 'min' });
  assert.deepEqual(renderDeck(measured, deck, { assets, layout: 'adaptive' }).warnings.filter((issue) => /too many/.test(issue.message)), []);
});

test('a row of a template with groups and copies is measured as it is rendered', () => {
  // The fixture's cards are a group and two loose shapes, in a right-to-left row that takes one
  // to five: fewer than the three drawn widen, more are copies of the last.
  const wide: TemplateManifest = { ...cardsManifest, patterns: cardsManifest.patterns.map((pattern) => (pattern.repeat ? { ...pattern, repeat: { ...pattern.repeat, widen: 2 } } : pattern)) };
  const file = fixtureWithTypeface();
  const template = loadTemplate(file, wide, { fonts: fixedWidth });
  const measurer = textMeasurer(template);
  const cards = patternOf('cards', wide);
  const rooms: number[] = [];
  for (let items = 1; items <= 5; items++) {
    const slide: SlideSpec = { pattern: 'cards', items, slots: { title: 'T', ...Object.fromEntries(Array.from({ length: items }, (_, index) => [`card${index + 1}Text`, 'word'])) } };
    const rendered = new TextStyles(PptxPackage.load(renderDeck(loadTemplate(file, wide), { version: 1, slides: [slide] }, { layout: 'adaptive' }).pptx));
    for (const slot of effectiveSlots(cards, items)) {
      if (slot.kind !== 'text' || slot.name === 'title') continue;
      const frame = rendered.frame(1, slot.shapeId);
      assert.ok(frame, `${slot.name} of ${items} is on the rendered slide`);
      const fit = measurer.fit(cards, slot, 'word', { items });
      assert.ok(fit && Math.abs((fit.lineRoom[0] ?? 0) - (frame.width - frame.insets.left - frame.insets.right)) < 0.02, `${slot.name} of ${items}: measured ${fit?.lineRoom[0]}, rendered ${frame.width - frame.insets.left - frame.insets.right}`);
      if (slot.name === 'card1Text') rooms.push(Math.round(fit.lineRoom[0] ?? 0));
    }
  }
  // One card and two cards are twice as wide, which is the limit the manifest sets and less than
  // the row has room for; three are as drawn, and so are the copies that make four and five.
  // The box is 78.74 pt wide as drawn and its text keeps 7.2 pt from each side.
  assert.deepEqual(rooms, [143, 143, 64, 64, 64]);
});

test('the planner accepts text for two wide cards that it would send back for three narrow ones', async () => {
  // 9 pt text, 4.5 pt a character. Sixty characters fit a card of two; a card as drawn holds fewer.
  const wide: TemplateManifest = {
    ...cardsManifest,
    patterns: cardsManifest.patterns.map((pattern) =>
      pattern.repeat ? { ...pattern, slots: pattern.slots.map((slot) => (slot.kind === 'text' && slot.name !== 'title' ? { ...slot, maxChars: 400 } : slot)), repeat: { ...pattern.repeat, widen: 2 } } : pattern,
    ),
  };
  const template = loadTemplate(fixtureWithTypeface(), wide, { fonts: fixedWidth });
  const measurer = textMeasurer(template);
  const cards = patternOf('cards', wide);
  // The longest text that fits a card of two, by measurement.
  let text = 'word';
  while (measurer.fit(cards, 'card1Text', `${text} word`, { items: 2 })?.fits) text = `${text} word`;
  assert.equal(measurer.fit(cards, 'card1Text', text, { items: 3 })?.fits, false, 'the same text is too long for a card as drawn');

  const two = scripted({ title: 'T', slides: [{ pattern: 'cards', purpose: 'Two things', items: 2 }] }, { s1: { title: 'Two', card1Text: text, card2Text: text } });
  const planned = await planDeck(template, { prompt: 'Two things', language: 'he', imageDensity: 'none' }, { llm: two.llm, patterns: ['cards'], review: designReview(template) });
  assert.equal(two.requests.length, 2, 'no repair round: the draft fits the wider cards');
  assert.deepEqual(planned.warnings, []);

  const three = scripted(
    { title: 'T', slides: [{ pattern: 'cards', purpose: 'Three things', items: 3 }] },
    { s1: { title: 'Three', card1Text: text, card2Text: text, card3Text: text } },
    { s1: { title: 'Three', card1Text: 'word', card2Text: 'word', card3Text: 'word' } },
  );
  await planDeck(template, { prompt: 'Three things', language: 'he', imageDensity: 'none' }, { llm: three.llm, patterns: ['cards'] });
  assert.equal(three.requests.length, 3, 'one repair round: the draft runs over the cards as drawn');
  assert.match(three.requests[2]?.prompt ?? '', /s1\.card1Text: text takes \d+ lines at 9 pt; the box holds \d+/);
});

// --- derived patterns ----------------------------------------------------------------------------

test('a mirrored pattern is measured like the pattern it mirrors, and a text-only one in its wider box', () => {
  for (const [source, mirror] of [['text-and-photo', 'text-and-photo-mirrored'], ['photo-stat', 'photo-stat-mirrored'], ['list-with-photo', 'list-with-photo-mirrored']] as const) {
    const texts = patternOf(mirror).slots.filter((slot) => slot.kind === 'text' && !slot.deckField && slot.empty !== 'keep' && slot.empty !== 'drop');
    assert.ok(texts.length >= 2, mirror);
    for (const slot of texts) {
      const [a, b] = [measure.fit(source, slot.name, words(40)), measure.fit(mirror, slot.name, words(40))];
      assert.ok(a && b, `${mirror} ${slot.name}`);
      // The mirror image states its boxes to four decimal places of the slide, so they are the same to a twentieth of a point.
      assert.ok(Math.abs((b.lineRoom[0] ?? 0) - (a.lineRoom[0] ?? 0)) < 0.05 && Math.abs(b.room - a.room) < 0.05, `${mirror} ${slot.name}: ${b.lineRoom[0]} by ${b.room} against ${a.lineRoom[0]} by ${a.room}`);
      assert.deepEqual([b.lines, b.capacity, b.size], [a.lines, a.capacity, a.size], `${mirror} ${slot.name}: the same text breaks the same way`);
      // The box only moved, so how full the designer drew it is still known.
      assert.equal(b.sampleLines, a.sampleLines, `${mirror} ${slot.name}`);
      assert.ok(a.sampleLines > 0);
    }
  }
  const [beside, alone] = [measure.fit('text-and-photo', 'body', words(40)), measure.fit('text-only', 'body', words(40))];
  assert.ok(beside && alone && (alone.lineRoom[0] ?? 0) > (beside.lineRoom[0] ?? 0), 'without the photo the text has a wider box');
});

test('variants made by the studio have budgets that hold when they are measured', () => {
  // The demo manifest without its own variants, so that the operation makes them here.
  const base: TemplateManifest = { ...manifest, patterns: manifest.patterns.filter((pattern) => !pattern.id.endsWith('-mirrored') && pattern.id !== 'text-only') };
  const result = changeManifest(loadTemplate(pptx, base), [
    { op: 'derivePattern', pattern: 'text-and-photo', variant: 'mirror', id: 'made-mirror' },
    { op: 'derivePattern', pattern: 'text-and-photo', variant: 'text-only', id: 'made-text-only' },
    { op: 'derivePattern', pattern: 'list-with-photo', variant: 'text-only', id: 'made-list-only' },
    { op: 'derivePattern', pattern: 'text-and-photo', variant: 'without', slot: 'subtitle', id: 'made-without' },
  ]);
  assert.deepEqual(result.rejected, []);
  const made = loadTemplate(pptx, result.manifest, { fonts: fontkitMetrics(fontkit, openFonts()) });
  // Every budget the operation worked out is tried with real metrics: none lets text run over its box.
  assert.deepEqual(checkManifest(made), []);
  // And the measured budgets are at least what the operation gave: it was cautious, not generous.
  const budgets = measuredBudgets(made).filter((entry) => entry.pattern.startsWith('made-'));
  assert.ok(budgets.length >= 8, `${budgets.length} text slots of the made patterns were measured`);
  for (const entry of budgets) assert.ok(entry.fits >= (entry.maxChars ?? 0), `${entry.pattern} ${entry.slot}: holds ${entry.fits}, budget ${entry.maxChars}`);
  // A deck that uses them renders with nothing over its box.
  const { deck, assets } = sampleDeck(made);
  assert.deepEqual(renderDeck(made, deck, { assets, layout: 'adaptive' }).warnings, []);
});

// --- alternatives --------------------------------------------------------------------------------

test('an alternative that is within its budgets also fits when it is measured', () => {
  const deck = JSON.parse(readFileSync(join(root, 'examples', 'decks', 'patterns.en.deck.json'), 'utf8')) as DeckSpec;
  const measurer = textMeasurer(real);
  let alternatives = 0;
  deck.slides.forEach((_, index) => {
    for (const alternative of slideAlternatives(manifest, deck, index + 1, { layout: 'adaptive' })) {
      alternatives++;
      const one: DeckSpec = { ...deck, slides: [alternative.slide] };
      const issues = validateDeck(manifest, one, { layout: 'adaptive', measure: measurer });
      assert.deepEqual(issues.filter((issue) => issue.severity === 'error'), [], `slide ${index + 1} as ${alternative.pattern}`);
      // Budgets are cautious, so what passes them passes the measurement too.
      if (alternative.warnings.length === 0) assert.deepEqual(issues, [], `slide ${index + 1} as ${alternative.pattern}`);
      // The design checks run on it and name only slots of the new slide.
      for (const issue of checkDesign(real, one, { layout: 'adaptive' })) assert.ok(issue.slot in alternative.slide.slots || issue.slot in (deck.fields ?? {}), `${alternative.pattern}: ${issue.slot}`);
    }
  });
  assert.ok(alternatives >= 20, `${alternatives} alternatives were measured`);
});

// --- previews ------------------------------------------------------------------------------------

test('every pattern gets a picture of its own, derived ones included, and a widened row reaches the renderer widened', async () => {
  const seen: Uint8Array[] = [];
  const counting: PreviewProvider = {
    async renderSlides({ pptx: file }) {
      seen.push(file);
      return readPresentation(PptxPackage.load(file)).slides.map((_, index) => ({ slide: index + 1, mimeType: 'image/svg+xml', data: new TextEncoder().encode(`<svg data-slide="${index + 1}"/>`) }));
    },
  };
  const pictures = await patternPreviews(plain, counting);
  assert.deepEqual(Object.keys(pictures), manifest.patterns.map((pattern) => pattern.id));
  const text = (id: string) => new TextDecoder().decode(pictures[id]?.data);
  // A derived pattern is the slide of the sample deck made for it, not the slide it was derived from.
  const at = (id: string) => manifest.patterns.findIndex((pattern) => pattern.id === id) + 1;
  assert.equal(text('text-and-photo-mirrored'), `<svg data-slide="${at('text-and-photo-mirrored')}"/>`);
  assert.notEqual(at('text-and-photo-mirrored'), at('text-and-photo'));

  // What the provider is handed is the file a deck gets: two cards, each wider than a card as drawn.
  const cards = patternOf('icon-cards');
  await previewDeck(plain, deckOf(rowSlide(cards, 2)), counting, { assets: { icon }, layout: 'adaptive' });
  const handed = new TextStyles(PptxPackage.load(seen.at(-1) as Uint8Array));
  const drawn = new TextStyles(plain.pkg).frame(cards.sourceSlide, patternOf('icon-cards').slots.find((slot) => slot.name === 'card1Text')?.shapeId ?? '');
  const wide = handed.frame(1, cards.slots.find((slot) => slot.name === 'card1Text')?.shapeId ?? '');
  assert.ok(drawn && wide && wide.width > drawn.width * 2, `${wide?.width} pt against ${drawn?.width} pt`);
});

test('Pagus draws mirrored slides and widened rows from the same file PowerPoint gets', async () => {
  const pagus = pagusPreviews(pagusCore, pagusRenderer);
  const cards = patternOf('icon-cards');
  const mirrored: SlideSpec = { pattern: 'text-and-photo-mirrored', slots: { title: 'A mirrored title', subtitle: 'A supporting line', body: ['The body of the mirrored slide.'], photo: { asset: 'icon' } } };
  const deck = deckOf(mirrored, rowSlide(cards, 2, (name) => (name.endsWith('Text') ? 'Text of a wide card' : name.endsWith('Heading') ? 'Wide heading' : 'Two wide cards')));
  const { previews } = await previewDeck(plain, deck, pagus, { assets: { icon }, layout: 'adaptive' });
  assert.deepEqual(previews.map((picture) => [picture.slide, picture.mimeType]), [[1, 'image/svg+xml'], [2, 'image/svg+xml']]);
  const [first, second] = previews.map((picture) => new TextDecoder().decode(picture.data));
  assert.match(first ?? '', /A mirrored title/);
  assert.match(second ?? '', /Wide heading/);
  assert.match(second ?? '', /Text of a wide card/);
  // Two cards are drawn, not the four of the template.
  assert.equal((second ?? '').split('Wide heading').length - 1, 2);
});

// --- nothing changes in the file ---------------------------------------------------------------------

test('measuring changes what is reported, never the file, unless text is to be shrunk', () => {
  for (const items of ['min', 'max', undefined] as const) {
    const { deck, assets } = sampleDeck(plain, items ? { items } : {});
    const a = renderDeck(plain, deck, { assets, layout: 'adaptive' });
    const b = renderDeck(measured, deck, { assets, layout: 'adaptive' });
    assert.ok(Buffer.from(a.pptx).equals(Buffer.from(b.pptx)), `the sample deck with ${items ?? 'drawn'} items is the same file with and without fonts`);
    assert.deepEqual(b.shrunk, []);
  }
});
