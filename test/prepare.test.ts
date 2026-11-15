// prepareManifest: the one step a studio runs after a draft, in place of going through every
// suggestion and issue by hand. No model is asked anywhere in it.
import assert from 'node:assert/strict';
import { readFileSync } from 'node:fs';
import { dirname, join } from 'node:path';
import { test } from 'node:test';
import { fileURLToPath } from 'node:url';
import * as fontkit from 'fontkit';
import { changeManifest, checkManifest, fontkitMetrics, loadTemplate, planOutline, prepareManifest, renderDeck, sampleDeck, suggestChanges, validateDeck } from '../src/index.ts';
import type { JsonRequest, LlmProvider, PatternDef, TemplateManifest } from '../src/index.ts';
import { buildFixturePptx, cardsManifest, fixtureManifest, gridManifest } from './fixture.ts';
import { openFonts } from './fonts.ts';

const root = join(dirname(fileURLToPath(import.meta.url)), '..');
const pptx = buildFixturePptx();
const demoPptx = readFileSync(join(root, 'templates', 'demo.pptx'));
const demoManifest = JSON.parse(readFileSync(join(root, 'templates', 'demo.manifest.json'), 'utf8')) as TemplateManifest;

const patternOf = (manifest: TemplateManifest, id: string) => manifest.patterns.find((pattern) => pattern.id === id) as PatternDef;
const namesOf = (manifest: TemplateManifest) => manifest.patterns.flatMap((pattern) => [pattern.id, ...pattern.slots.map((slot) => `${pattern.id}.${slot.name}`)]);

// A manifest as a draft of 0.2.0 left it: the pictograms of the cards are photo slots, and the row takes what it draws.
const cards = patternOf(cardsManifest, 'cards');
const asDrafted: TemplateManifest = {
  ...cardsManifest,
  patterns: [
    ...fixtureManifest.patterns,
    {
      ...cards,
      slots: cards.slots.map((slot) => (slot.kind === 'image' ? { name: slot.name, kind: 'image', imageKind: 'photo', shapeId: slot.shapeId } : slot)),
      repeat: { item: 'card', slots: ['card{n}Text', 'card{n}Icon'], shapes: [['10', '13'], ['20'], ['30']], min: 2, max: 3 },
    },
  ],
};

test('prepareManifest takes what the slides allow in one step, and leaves a short list for a person', () => {
  const template = loadTemplate(pptx, asDrafted);
  assert.deepEqual(suggestChanges(template).map((suggestion) => suggestion.kind), ['icon', 'row-max', 'row-widen'], 'what a person would otherwise go through');
  const prepared = prepareManifest(template, { icons: true });

  // The pictograms are icons, the row takes what the slide has room for, and fewer cards fill the line.
  const row = patternOf(prepared.manifest, 'cards');
  assert.deepEqual(row.slots.filter((slot) => slot.kind === 'image').map((slot) => slot.imageKind), ['icon', 'icon', 'icon']);
  assert.deepEqual(row.repeat, { item: 'card', slots: ['card{n}Text', 'card{n}Icon'], shapes: [['10', '13'], ['20'], ['30']], min: 2, max: 6, span: { left: 0.03, right: 0.97 }, widen: 2 });
  assert.equal(prepared.sound, true);
  assert.deepEqual([prepared.added, prepared.removed], [[], []]);
  assert.equal(prepared.done.length, 3, prepared.done.join(' | '));
  assert.match(prepared.done[0] ?? '', /"card1Icon", "card2Icon", "card3Icon" are now an icon/);

  // Nothing of what it takes is left to suggest, and the result is a manifest a deck can use at once.
  const after = loadTemplate(pptx, prepared.manifest);
  assert.deepEqual(suggestChanges(after), []);
  assert.deepEqual(checkManifest(after).filter((issue) => issue.severity === 'error'), []);
  const { deck, assets } = sampleDeck(after, { items: 'max' });
  assert.deepEqual(validateDeck(prepared.manifest, deck, { assets: Object.keys(assets) }), []);
  assert.ok(renderDeck(after, deck, { assets }).pptx.length > 1000);

  // What is left is for a person, and is short: the deck's own texts, the fonts, the descriptions, and warnings it cannot answer.
  assert.equal(prepared.left.length, 4, prepared.left.join(' | '));
  assert.match(prepared.left[0] ?? '', /^3 warnings to look at: content\.body has no character budget/);
  assert.match(prepared.left[1] ?? '', /^One text is the same on every slide and shows what the template shows until you set your own: closing \("SAMPLE CLOSING"\)\.$/);
  assert.match(prepared.left[2] ?? '', /^Text was not measured with the template's fonts/);
  assert.match(prepared.left[3] ?? '', /^Read the description of each layout once\./);
  assert.deepEqual(prepared.fields, { closing: 'SAMPLE CLOSING' });

  // Nothing is renamed, the template passed in is untouched, and a second run finds nothing to do.
  assert.deepEqual(namesOf(prepared.manifest), namesOf(asDrafted));
  assert.equal(patternOf(template.manifest, 'cards').repeat?.max, 3);
  const again = prepareManifest(after, { icons: true });
  assert.deepEqual([again.done, again.added, again.removed], [[], [], []]);
  assert.deepEqual(again.manifest, prepared.manifest);
});

test('without an icon set, icon slots keep the template\'s pictograms, so their layouts are still offered', async () => {
  const prepared = prepareManifest(loadTemplate(pptx, asDrafted), { icons: false });
  assert.deepEqual(patternOf(prepared.manifest, 'cards').slots.filter((slot) => slot.kind === 'image').map((slot) => [slot.imageKind, slot.empty]), [['icon', 'keep'], ['icon', 'keep'], ['icon', 'keep']]);
  assert.equal(prepared.sound, true);

  // The planner is offered the row with pictures off and no icons at all. Before, it was not.
  const offered = async (manifest: TemplateManifest) => {
    const requests: JsonRequest[] = [];
    const llm: LlmProvider = { async generateJson(request) { requests.push(request); throw new Error('stop'); } };
    await planOutline(loadTemplate(pptx, manifest), { prompt: 'Three things', language: 'en', imageDensity: 'none', fields: { closing: 'Bye' } }, { llm }).catch(() => undefined);
    return /\n- cards: /.test(requests[0]?.prompt ?? '');
  };
  assert.equal(await offered(asDrafted), false, 'photo slots, and pictures are off');
  assert.equal(await offered(prepareManifest(loadTemplate(pptx, asDrafted), { icons: true }).manifest), false, 'icon slots, and there is no icon set');
  assert.equal(await offered(prepared.manifest), true);
});

test('variants and widening are taken by default and can be left out', () => {
  // The demo manifest ships with some variants. The step adds the ones it lacks, and none twice.
  const template = loadTemplate(demoPptx, demoManifest);
  const prepared = prepareManifest(template, { icons: true });
  assert.ok(prepared.added.length >= 3, prepared.added.join(', '));
  assert.ok(prepared.added.every((id) => /-(mirrored|text-only)$/.test(id)), prepared.added.join(', '));
  assert.equal(prepared.manifest.patterns.length, demoManifest.patterns.length + prepared.added.length);
  assert.equal(prepared.sound, true);
  assert.deepEqual(namesOf(prepared.manifest).filter((name) => namesOf(demoManifest).includes(name)), namesOf(demoManifest), 'every name the manifest had is still there');
  const rest = suggestChanges(loadTemplate(demoPptx, prepared.manifest));
  assert.ok(rest.every((suggestion) => suggestion.variant === 'without'), 'only the variants it leaves to a person remain');
  assert.deepEqual(prepareManifest(loadTemplate(demoPptx, prepared.manifest), { icons: true }).done, []);

  // With no variants and no widening it settles the facts only.
  const plain = prepareManifest(template, { icons: true, variants: [], widen: false });
  assert.deepEqual(plain.added, []);
  assert.ok(plain.manifest.patterns.every((pattern, index) => pattern.repeat?.widen === demoManifest.patterns[index]?.repeat?.widen));
  // Asked for, the shortened variants are added as well.
  const all = prepareManifest(template, { icons: true, variants: ['mirror', 'text-only', 'without'] });
  assert.ok(all.added.length > prepared.added.length && all.added.some((id) => /-no-/.test(id)), all.added.join(', '));
  assert.deepEqual(suggestChanges(loadTemplate(demoPptx, all.manifest)), []);

  assert.throws(() => prepareManifest(template, {} as never), /say whether the app has an icon set for this template: icons is true or false/);
  assert.throws(() => prepareManifest(template, { icons: true, widen: 9 }), /widen is a factor above 1 and at most 4, such as 2, or false/);
  assert.throws(() => prepareManifest(template, { icons: true, variants: ['flip' as never] }), /"flip" is not a variant; the variants are mirror, text-only, without/);
});

test('issues that have one answer are solved, and a pattern that cannot be rendered is taken out and said', () => {
  // One pattern points at a shape the slide does not have; a grid is said to take more than it draws.
  const grid = patternOf(gridManifest, 'grid');
  const faulty: TemplateManifest = {
    ...cardsManifest,
    patterns: [
      ...cardsManifest.patterns,
      { ...grid, repeat: { ...(grid.repeat as NonNullable<typeof grid.repeat>), max: 6 } },
      { id: 'broken', sourceSlide: 1, description: 'A pattern from a slide that was edited since.', slots: [{ name: 'gone', kind: 'text', shapeId: '99', maxChars: 10 }] },
    ],
  };
  const template = loadTemplate(pptx, faulty);
  assert.ok(checkManifest(template).some((issue) => issue.severity === 'error'));
  const prepared = prepareManifest(template, { icons: true });
  assert.deepEqual(prepared.removed, ['broken']);
  assert.ok(prepared.done.some((sentence) => /^"broken" was taken out, because it cannot be rendered: gone points at shape #99, which is not on slide 1\.$/.test(sentence)), prepared.done.join(' | '));
  assert.equal(patternOf(prepared.manifest, 'grid').repeat?.max, 4, 'a grid takes what it draws');
  assert.equal(prepared.sound, true);
  assert.deepEqual(checkManifest(loadTemplate(pptx, prepared.manifest)).filter((issue) => issue.severity === 'error'), []);

  // A manifest with nothing in it is not fit to use, and says what to do.
  const empty = prepareManifest(loadTemplate(pptx, { ...fixtureManifest, patterns: [] }), { icons: true });
  assert.equal(empty.sound, false);
  assert.deepEqual(empty.left, ['The manifest has no layout. Draft the template first.']);
});

test('with the template\'s fonts, a budget that lets text run over its box gets the measured one', () => {
  const fonts = fontkitMetrics(fontkit, openFonts());
  const generous: TemplateManifest = {
    ...demoManifest,
    patterns: demoManifest.patterns.map((pattern) => (pattern.id !== 'text-only' ? pattern : { ...pattern, slots: pattern.slots.map((slot) => (slot.name === 'body' && slot.kind === 'text' ? { ...slot, maxChars: 3000 } : slot)) })),
  };
  const template = loadTemplate(demoPptx, generous, { fonts });
  assert.equal(checkManifest(template).filter((issue) => issue.message.startsWith('has a budget of 3000')).length, 1);
  const prepared = prepareManifest(template, { icons: true, variants: [], widen: false });
  const budget = (patternOf(prepared.manifest, 'text-only').slots.find((slot) => slot.name === 'body') as { maxChars: number }).maxChars;
  assert.ok(budget > 100 && budget < 3000, `the body now takes ${budget} characters`);
  assert.deepEqual(checkManifest(loadTemplate(demoPptx, prepared.manifest, { fonts })).filter((issue) => issue.message.startsWith('has a budget of')), []);
  // Measured, so the sentence about estimates is not said.
  assert.ok(!prepared.left.some((sentence) => sentence.startsWith('Text was not measured')));
  assert.deepEqual(changeManifest(template, []).manifest, generous, 'the template passed in is as it was');
});
