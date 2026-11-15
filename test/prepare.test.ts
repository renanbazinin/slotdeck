// prepareManifest: the one step a studio runs after a draft, in place of going through every
// suggestion and issue by hand. No model is asked anywhere in it.
import assert from 'node:assert/strict';
import { readFileSync } from 'node:fs';
import { dirname, join } from 'node:path';
import { test } from 'node:test';
import { fileURLToPath } from 'node:url';
import { strFromU8, strToU8, unzipSync, zipSync } from 'fflate';
import * as fontkit from 'fontkit';
import { placeholderPhoto } from '../examples/placeholder.ts';
import { changeManifest, checkManifest, fontkitMetrics, loadTemplate, planOutline, prepareManifest, renderDeck, sampleDeck, suggestChanges, validateDeck } from '../src/index.ts';
import type { DeckSpec, JsonRequest, LlmProvider, PatternDef, TemplateManifest } from '../src/index.ts';
import { buildFixturePptx, cardsManifest, fixtureManifest, fixtureWithSlide, fixtureWithTextCards, gridManifest } from './fixture.ts';
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
  // It says which slots it set, so that they can be filled again when the app gets an icon set.
  assert.deepEqual(prepared.kept, [1, 2, 3].map((n) => ({ pattern: 'cards', slot: `card${n}Icon` })));
  assert.deepEqual(prepareManifest(loadTemplate(pptx, asDrafted), { icons: true }).kept, []);

  // An icon slot a deck may leave out is left as it is: the layout is offered anyway, and decks that left it out look as before.
  const optional: TemplateManifest = { ...asDrafted, patterns: asDrafted.patterns.map((pattern) => ({ ...pattern, slots: pattern.slots.map((slot) => (slot.kind === 'image' && pattern.id === 'cards' ? { ...slot, imageKind: 'icon' as const, empty: 'remove' as const } : slot)) })) };
  const left = prepareManifest(loadTemplate(pptx, optional), { icons: false });
  assert.deepEqual(patternOf(left.manifest, 'cards').slots.filter((slot) => slot.kind === 'image').map((slot) => slot.empty), ['remove', 'remove', 'remove']);
  assert.deepEqual(left.kept, []);

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

test('issues that have one answer are solved, and a pattern that cannot be used is taken out and said', () => {
  const template = loadTemplate(pptx, faulty);
  assert.ok(checkManifest(template).some((issue) => issue.severity === 'error'));
  const prepared = prepareManifest(template, { icons: true });
  assert.deepEqual(prepared.removed, ['broken']);
  assert.ok(prepared.done.includes('"broken" was taken out: "gone" points at shape #99, which is not on slide 1.'), prepared.done.join(' | '));
  // A removal is the one thing a person must hear of first: decks made with that layout are stuck.
  assert.equal(prepared.left[0], 'One layout was taken out because it cannot be used: "broken" ("gone" points at shape #99, which is not on slide 1). A presentation made with it can no longer be changed.');
  assert.deepEqual(patternOf(prepared.manifest, 'grid').repeat, { ...grid.repeat, min: 1, max: 4, widen: 2 }, 'a grid takes what it draws, and then widens like any row');
  assert.equal(prepared.sound, true);
  const after = loadTemplate(pptx, prepared.manifest);
  assert.deepEqual(checkManifest(after).filter((issue) => issue.severity === 'error'), []);

  // What it solved opened nothing: a second run finds nothing to do.
  assert.deepEqual(suggestChanges(after), []);
  const again = prepareManifest(after, { icons: true });
  assert.deepEqual([again.done, again.removed, again.manifest], [[], [], prepared.manifest]);

  // A grid that does not say its fewest takes exactly what it draws, before and after.
  const { min: _min, ...exact } = grid.repeat as NonNullable<typeof grid.repeat>;
  const capped = prepareManifest(loadTemplate(pptx, { ...gridManifest, patterns: [...fixtureManifest.patterns, { ...grid, repeat: { ...exact, max: 6 } }] }), { icons: true });
  const now = patternOf(capped.manifest, 'grid').repeat;
  assert.deepEqual([now?.min ?? now?.shapes.length, now?.max], [4, 4]);

  // A manifest with nothing in it is not fit to use, and says what to do. The last pattern is never taken out.
  const empty = prepareManifest(loadTemplate(pptx, { ...fixtureManifest, patterns: [] }), { icons: true });
  assert.equal(empty.sound, false);
  assert.deepEqual(empty.left, ['The manifest has no layout. Draft the template first.']);
  const only = prepareManifest(loadTemplate(pptx, { ...fixtureManifest, patterns: faulty.patterns.slice(-1) }), { icons: true });
  assert.deepEqual([only.sound, only.removed, only.manifest.patterns.length], [false, [], 1]);
  assert.match(only.left[0] ?? '', /^One error could not be solved: broken\.gone points at shape #99/);
});

test('for a manifest in use, remove: false keeps every pattern and says what is wrong with it', () => {
  const prepared = prepareManifest(loadTemplate(pptx, faulty), { icons: true, remove: false });
  assert.deepEqual(prepared.removed, []);
  assert.deepEqual(namesOf(prepared.manifest), namesOf(faulty), 'every pattern and every slot a stored deck may name is still there');
  assert.equal(prepared.sound, false);
  assert.equal(prepared.left[0], 'One error could not be solved: broken.gone points at shape #99, which is not on slide 1.');
  assert.match(prepared.left[1] ?? '', /^The sample deck does not render: /);
  // What can be done without it is done all the same.
  assert.equal(patternOf(prepared.manifest, 'cards').repeat?.max, 6);
  assert.equal(patternOf(prepared.manifest, 'grid').repeat?.max, 4);

  // Two patterns with one id: nothing says which of them is meant, so both stay, also when patterns may be taken out.
  const [content, closing] = fixtureManifest.patterns as [PatternDef, PatternDef];
  const twice = prepareManifest(loadTemplate(pptx, { ...fixtureManifest, patterns: [content, closing, { ...closing, description: 'Another closing.' }] }), { icons: true });
  assert.deepEqual([twice.removed, twice.sound, twice.manifest.patterns.length], [[], false, 3]);
  assert.match(twice.left[0] ?? '', /^One error could not be solved: closing another pattern has the same id\.$/);
});

test('a pattern that cannot be used is taken out before anything is made of it', () => {
  // The demo manifest without its variants. One pattern has a slot on a shape its slide does not have.
  const bare: TemplateManifest = {
    ...demoManifest,
    patterns: demoManifest.patterns
      .filter((pattern) => !/-mirrored$|^text-only$|^chart-native$/.test(pattern.id))
      .map((pattern): PatternDef => (pattern.id !== 'text-and-photo' ? pattern : { ...pattern, slots: [...pattern.slots, { name: 'gone', kind: 'text', shapeId: '999', maxChars: 10 }] })),
  };
  const prepared = prepareManifest(loadTemplate(demoPptx, bare), { icons: true });
  assert.deepEqual(prepared.removed, ['text-and-photo'], 'only what the manifest came with');
  assert.ok(!prepared.done.some((sentence) => sentence.includes('"text-and-photo-')), 'no variant was made of it');
  assert.ok(prepared.added.length > 0 && prepared.sound);
  assert.deepEqual(prepareManifest(loadTemplate(demoPptx, prepared.manifest), { icons: true }).done, []);

  // A variant that cannot be used is taken out and made again from its pattern: the id is not gone.
  const stale: TemplateManifest = {
    ...demoManifest,
    patterns: demoManifest.patterns.map((pattern): PatternDef => (pattern.id !== 'photo-stat-mirrored' ? pattern : { ...pattern, slots: pattern.slots.map((slot, index) => (index === 0 ? { ...slot, shapeId: '999' } : slot)) })),
  };
  const mended = prepareManifest(loadTemplate(demoPptx, stale), { icons: true });
  assert.ok(mended.done.some((sentence) => sentence.startsWith('"photo-stat-mirrored" was taken out: ')), mended.done.join(' | '));
  assert.deepEqual([mended.removed, mended.added.includes('photo-stat-mirrored'), mended.sound], [[], false, true]);
  assert.deepEqual(namesOf(mended.manifest).filter((name) => namesOf(demoManifest).includes(name)), namesOf(demoManifest));
  assert.ok(!mended.left.some((sentence) => /taken out/.test(sentence)), 'nothing a deck names is gone, so nothing is said');
  assert.deepEqual(prepareManifest(loadTemplate(demoPptx, mended.manifest), { icons: true }).done, []);
});

test('a pattern whose sample cannot be rendered is taken out, though no error is reported for it', () => {
  // A text slot on a drawn rectangle that holds no text: a warning for checkManifest, and nothing to fill for the renderer.
  const drawn = fixtureWithSlide(1, (xml) => xml.replace('</p:spTree>', '<p:sp><p:nvSpPr><p:cNvPr id="50" name="Box"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x="100000" y="4000000"/><a:ext cx="800000" cy="300000"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr></p:sp></p:spTree>'));
  const odd: TemplateManifest = { ...fixtureManifest, patterns: [...fixtureManifest.patterns, { id: 'odd', sourceSlide: 1, description: 'A title and a label.', slots: [{ name: 'title', kind: 'text', shapeId: '2', maxChars: 20 }, { name: 'label', kind: 'text', shapeId: '50', maxChars: 10 }] }] };
  const template = loadTemplate(drawn, odd);
  assert.deepEqual(checkManifest(template).filter((issue) => issue.severity === 'error'), []);
  assert.throws(() => renderDeck(template, sampleDeck(template).deck, { assets: sampleDeck(template).assets }));

  const prepared = prepareManifest(template, { icons: true });
  assert.deepEqual([prepared.removed, prepared.sound], [['odd'], true]);
  assert.ok(prepared.done.some((sentence) => /^"odd" was taken out: its sample cannot be rendered \(/.test(sentence)), prepared.done.join(' | '));
  assert.match(prepared.left[0] ?? '', /^One layout was taken out because it cannot be used: "odd" \(its sample cannot be rendered/);

  const stays = prepareManifest(template, { icons: true, remove: false });
  assert.deepEqual([stays.removed, stays.sound], [[], false]);
  assert.ok(stays.left.some((sentence) => sentence.startsWith('The sample deck does not render: ')), stays.left.join(' | '));
});

test('a row whose items hold a slot that is not the row\'s is repaired, and every deck that rendered still does', () => {
  // As an earlier draft made it: the row is of the texts only, and each card's pictogram is a slot beside the row.
  const inUse = (empty?: 'keep'): TemplateManifest => ({
    ...fixtureManifest,
    patterns: [
      ...fixtureManifest.patterns,
      {
        ...cards,
        slots: cards.slots.map((slot) => (slot.kind === 'image' ? { name: slot.name, kind: 'image', imageKind: 'icon', shapeId: slot.shapeId, ...(empty ? { empty } : {}) } : slot)),
        repeat: { item: 'card', slots: ['card{n}Text'], shapes: [['10'], ['20'], ['30']], min: 2, max: 5, span: { left: 0.05, right: 0.95 } },
      },
    ],
  });
  const settle = { icons: true, variants: [], widen: false, remove: false } as const;
  const png = placeholderPhoto(8, 8, [0, 0, 0], [255, 255, 255]);
  const deckOf = (items: number, icons: boolean): DeckSpec => ({
    version: 1,
    title: 'Stored',
    direction: 'rtl',
    fields: { closing: 'Bye' },
    slides: [{ pattern: 'cards', items, slots: { title: 'T', ...Object.fromEntries(Array.from({ length: items }, (_, index) => [`card${index + 1}Text`, 'x'])), ...(icons ? { card1Icon: { asset: 'i' }, card2Icon: { asset: 'i' }, card3Icon: { asset: 'i' } } : {}) } }],
  });
  const renders = (manifest: TemplateManifest, deck: DeckSpec) => {
    try {
      return renderDeck(loadTemplate(pptx, manifest), deck, { assets: { i: png }, layout: 'adaptive' }).pptx.length > 1000;
    } catch {
      return false;
    }
  };

  // Pictograms kept as designed, which no deck fills: they become part of the row, and two cards can now be rendered.
  const kept = inUse('keep');
  assert.deepEqual([2, 3, 4].map((items) => renders(kept, deckOf(items, false))), [false, true, true], 'two cards could not be rendered');
  const joined = prepareManifest(loadTemplate(pptx, kept), settle);
  assert.deepEqual(patternOf(joined.manifest, 'cards').repeat, { item: 'card', slots: ['card{n}Text', 'card{n}Icon'], shapes: [['10'], ['20'], ['30']], min: 2, max: 6, span: { left: 0.05, right: 0.95 } });
  assert.equal(joined.done[0], '"cards": "card{n}Icon" is now part of the row of cards, because the shape lies in each card. A slide with fewer cards could not be rendered.');
  assert.deepEqual([2, 3, 4].map((items) => renders(joined.manifest, deckOf(items, false))), [true, true, true]);
  assert.equal(joined.sound, true);

  // Pictograms that decks fill: joining them would ask a stored deck of four cards for a fourth it never gave.
  // The row takes no fewer than the three it draws instead, which is what could be rendered all along.
  const filled = inUse();
  assert.deepEqual([2, 3, 4].map((items) => renders(filled, deckOf(items, true))), [false, true, true]);
  const held = prepareManifest(loadTemplate(pptx, filled), settle);
  assert.deepEqual([patternOf(held.manifest, 'cards').repeat?.slots, patternOf(held.manifest, 'cards').repeat?.min], [['card{n}Text'], 3]);
  assert.equal(held.done[0], '"cards": the row of cards takes no fewer than 3, because "card3Icon" lies in card 3 and is not part of the row. A slide with fewer could not be rendered.');
  assert.deepEqual([3, 4].map((items) => renders(held.manifest, deckOf(items, true))), [true, true]);
  assert.ok(validateDeck(held.manifest, deckOf(2, true), { assets: ['i'] }).some((issue) => issue.severity === 'error'), 'two cards are refused before rendering now');
  assert.equal(held.sound, true);

  for (const result of [joined, held]) {
    assert.deepEqual(namesOf(result.manifest), namesOf(kept));
    const again = prepareManifest(loadTemplate(pptx, result.manifest), settle);
    assert.deepEqual([again.done, again.manifest], [[], result.manifest]);
  }
});

test('a deck field the app has a value for is not asked about', () => {
  const template = loadTemplate(pptx, cardsManifest);
  const asked = prepareManifest(template, { icons: true });
  assert.deepEqual(asked.fields, { closing: 'SAMPLE CLOSING' });
  assert.ok(asked.left.some((sentence) => sentence.startsWith('One text is the same on every slide')));

  const set = prepareManifest(template, { icons: true, fields: { closing: 'Thank you' } });
  assert.deepEqual(set.fields, { closing: 'Thank you' });
  assert.ok(!set.left.some((sentence) => sentence.startsWith('One text is the same on every slide')), set.left.join(' | '));
  // A blank value is no value.
  assert.deepEqual(prepareManifest(template, { icons: true, fields: { closing: '  ' } }).fields, { closing: 'SAMPLE CLOSING' });
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

test('with fonts, the items of a row get one budget, the smallest any of them measures', () => {
  const fonts = fontkitMetrics(fontkit, openFonts());
  // Three cards as text boxes, each showing other words, so each measures a little differently.
  const files = unzipSync(fixtureWithTextCards([6000000, 4000000, 2000000]));
  const words = ['MARKET GROWTH WINS', 'OUR WARM COMMUNITY', 'NEW MEMBERS WELCOME'];
  let next = 0;
  files['ppt/slides/slide4.xml'] = strToU8(strFromU8(files['ppt/slides/slide4.xml'] as Uint8Array).replace(/<a:r><a:rPr sz="900"\/><a:t>SAMPLE CARD<\/a:t><\/a:r>/g, () => `<a:r><a:rPr sz="900" b="1"><a:latin typeface="Calibri"/></a:rPr><a:t>${words[next++]}</a:t></a:r>`));
  const wordy = zipSync(files);
  const built = changeManifest(loadTemplate(wordy, fixtureManifest), [
    { op: 'addPattern', slide: 4, id: 'cards', description: 'A row of cards.' },
    { op: 'addSlot', pattern: 'cards', shape: '2', role: 'title', name: 'title' },
    { op: 'addSlot', pattern: 'cards', shape: '11', role: 'text', name: 'card1Text' },
    { op: 'addSlot', pattern: 'cards', shape: '21', role: 'text', name: 'card2Text' },
    { op: 'addSlot', pattern: 'cards', shape: '31', role: 'text', name: 'card3Text' },
    { op: 'setRepeat', pattern: 'cards' },
  ]);
  assert.deepEqual(built.rejected, []);
  const over = (manifest: TemplateManifest) => checkManifest(loadTemplate(wordy, manifest, { fonts })).filter((issue) => issue.pattern === 'cards' && /^card\dText$/.test(issue.slot ?? '') && issue.message.startsWith('has a budget of'));
  assert.ok(over(built.manifest).length > 0, 'the estimates let text run over');

  const prepared = prepareManifest(loadTemplate(wordy, built.manifest, { fonts }), { icons: true, widen: false });
  const budgets = patternOf(prepared.manifest, 'cards').slots.filter((slot) => /^card\dText$/.test(slot.name)).map((slot) => (slot as { maxChars?: number }).maxChars);
  assert.equal(new Set(budgets).size, 1, `one budget for the three: ${budgets.join(', ')}`);
  assert.equal(prepared.done.filter((sentence) => /"card1Text", "card2Text", "card3Text" are now text, up to \d+ characters/.test(sentence)).length, 1, prepared.done.join(' | '));
  assert.deepEqual(over(prepared.manifest), [], 'and it fits every one of them');
  const again = prepareManifest(loadTemplate(wordy, prepared.manifest, { fonts }), { icons: true, widen: false });
  assert.deepEqual([again.done, again.manifest], [[], prepared.manifest]);
});

test('a budget that is the measured one already is not set again, and a second run says nothing', () => {
  const fonts = fontkitMetrics(fontkit, openFonts());
  // The rows of the demo without their widening, and the headings of the cards at what their box measures.
  const flat: TemplateManifest = {
    ...demoManifest,
    patterns: demoManifest.patterns.map((pattern): PatternDef => {
      if (!pattern.repeat) return pattern;
      const { widen: _widen, ...repeat } = pattern.repeat;
      return { ...pattern, repeat, slots: pattern.slots.map((slot) => (pattern.id === 'icon-cards' && /Heading$/.test(slot.name) && slot.kind === 'text' ? { ...slot, maxChars: 21 } : slot)) };
    }),
  };
  const options = { icons: true, variants: [], widen: 1.2 } as const;
  const prepared = prepareManifest(loadTemplate(demoPptx, flat, { fonts }), options);
  assert.equal(patternOf(prepared.manifest, 'icon-cards').repeat?.widen, 1.2);
  assert.ok(!prepared.done.some((sentence) => /are now text, up to 21 characters/.test(sentence)), prepared.done.join(' | '));
  assert.deepEqual(prepareManifest(loadTemplate(demoPptx, prepared.manifest, { fonts }), options).done, []);
});
