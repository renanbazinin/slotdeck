// Other layouts for a slide that already has its content: the same words in another pattern,
// worked out by code. A title stays a title; a row of items becomes other items or a list.
import assert from 'node:assert/strict';
import { readFileSync } from 'node:fs';
import { dirname, join } from 'node:path';
import { test } from 'node:test';
import { fileURLToPath } from 'node:url';
import { loadTemplate, renderDeck, slideAlternatives, validateDeck } from '../src/index.ts';
import type { DeckSpec, PatternDef, SlideSpec, SlotDef, TemplateManifest } from '../src/index.ts';
import { withoutCodes } from './issues.ts';

const root = join(dirname(fileURLToPath(import.meta.url)), '..');
const manifest = JSON.parse(readFileSync(join(root, 'templates', 'demo.manifest.json'), 'utf8')) as TemplateManifest;
const template = loadTemplate(readFileSync(join(root, 'templates', 'demo.pptx')), manifest);
const example = JSON.parse(readFileSync(join(root, 'examples', 'decks', 'patterns.en.deck.json'), 'utf8')) as DeckSpec;
const picture = new Uint8Array(readFileSync(join(root, 'examples', 'icons', 'star.png')));
const fields = { organizationDepartment: 'Org' };

const deckOf = (...slides: SlideSpec[]): DeckSpec => ({ version: 1, fields, slides });
const ids = (found: Array<{ pattern: string }>) => found.map((alternative) => alternative.pattern);

const cards = (items: number): SlideSpec => ({
  pattern: 'icon-cards',
  items,
  slots: {
    title: 'Changes for next season',
    ...Object.fromEntries(
      Array.from({ length: items }, (_, index) => [
        [`card${index + 1}Icon`, { asset: 'clock' }],
        [`card${index + 1}Heading`, ['Longer hours', 'Safer docks', 'Group passes'][index] as string],
        [`card${index + 1}Text`, ['Bikes unlock from five.', 'Every station gets lighting.', 'One pass for five riders.'][index] as string],
      ]).flat(),
    ),
  },
});

test('a row of cards can be a list or columns: the words are the same, the slots are not', () => {
  const three = slideAlternatives(manifest, deckOf(cards(3)), 1, { layout: 'adaptive' });
  // Three cards do not fit the two columns; the grid of figures wants numbers; the slides with a photo want a photo.
  assert.deepEqual(ids(three), ['text-only']);
  assert.deepEqual(three[0], {
    pattern: 'text-only',
    description: manifest.patterns.find((pattern) => pattern.id === 'text-only')?.description,
    slide: {
      pattern: 'text-only',
      slots: { title: 'Changes for next season', body: ['Longer hours: Bikes unlock from five.', 'Safer docks: Every station gets lighting.', 'Group passes: One pass for five riders.'] },
    },
    changes: ['The 3 cards became a list of 3 points, each written as "heading: text".'],
    dropped: ['card1Icon', 'card2Icon', 'card3Icon'],
    warnings: [],
  });

  const two = slideAlternatives(manifest, deckOf(cards(2)), 1, { layout: 'adaptive' });
  assert.deepEqual(ids(two), ['text-only', 'two-columns']);
  assert.deepEqual(two[1]?.slide, {
    pattern: 'two-columns',
    items: 2,
    slots: { title: 'Changes for next season', col1Heading: 'Longer hours', col1Text: 'Bikes unlock from five.', col2Heading: 'Safer docks', col2Text: 'Every station gets lighting.' },
  });
  assert.deepEqual([two[1]?.changes, two[1]?.dropped], [['The 2 cards became 2 columns.'], ['card1Icon', 'card2Icon']]);

  // In strict layout every slot of the new pattern must be filled, and the text-only slide has a subtitle the cards cannot fill.
  assert.deepEqual(ids(slideAlternatives(manifest, deckOf(cards(2)), 1)), ['two-columns']);
  assert.deepEqual(ids(slideAlternatives(manifest, deckOf(cards(2)), 1, { layout: 'adaptive', patterns: ['text-only', 'cover'] })), ['text-only']);
});

test('a list becomes items when it is written "heading: text", and only then', () => {
  const list = (body: string[]): SlideSpec => ({ pattern: 'text-only', slots: { title: 'Two ways to grow', body } });
  const [columns] = slideAlternatives(manifest, deckOf(list(['More stations: reaches two districts.', 'More bikes – ends the empty docks.'])), 1, { layout: 'adaptive' });
  assert.deepEqual(columns?.slide, {
    pattern: 'two-columns',
    items: 2,
    slots: { title: 'Two ways to grow', col1Heading: 'More stations', col1Text: 'reaches two districts.', col2Heading: 'More bikes', col2Text: 'ends the empty docks.' },
  });
  assert.deepEqual(columns?.changes, ['The 2 points became 2 columns, each split at the colon into a heading and a text.']);
  // Without a heading in each point there is nothing to put in the columns' headings, and code does not invent one.
  assert.deepEqual(slideAlternatives(manifest, deckOf(list(['Reaches two districts.', 'Ends the empty docks.'])), 1, { layout: 'adaptive' }), []);
  // Three points are one too many for two columns.
  assert.deepEqual(slideAlternatives(manifest, deckOf(list(['A: one.', 'B: two.', 'C: three.'])), 1, { layout: 'adaptive' }), []);
});

test('variants of one slide take each other’s content slot for slot; a photo may be left out, and that is said', () => {
  const slide: SlideSpec = { pattern: 'text-and-photo', slots: { title: 'T', subtitle: 'S', body: ['One.', 'Two.'], photo: { asset: 'photo', alt: 'A' } } };
  const found = slideAlternatives(manifest, deckOf(slide), 1);
  assert.deepEqual(ids(found), ['text-and-photo-mirrored', 'text-only']);
  assert.deepEqual(found[0]?.slide, { ...slide, pattern: 'text-and-photo-mirrored' });
  assert.deepEqual([found[0]?.changes, found[0]?.dropped], [[], []]);
  assert.deepEqual([found[1]?.slide.slots, found[1]?.dropped], [{ title: 'T', subtitle: 'S', body: ['One.', 'Two.'] }, ['photo']]);

  // A list beside a photo can also be paragraphs beside a photo: the change is named.
  const listed = slideAlternatives(manifest, example, example.slides.findIndex((entry) => entry.pattern === 'list-with-photo') + 1, { layout: 'adaptive' });
  assert.deepEqual(ids(listed), ['text-and-photo', 'text-and-photo-mirrored', 'list-with-photo-mirrored', 'text-only']);
  assert.deepEqual(listed[0]?.changes, ['The 4 points of the list are written as 4 paragraphs.']);
  assert.deepEqual(listed[2]?.changes, []);
});

test('what does not fit is reported, and comes after what fits', () => {
  const slide: SlideSpec = { pattern: 'cover', slots: { title: 'A year of Riverbend Bike Share', subtitle: 'What the season taught us' } };
  const found = slideAlternatives(manifest, deckOf({ pattern: 'section', slots: { title: 'x' } }, slide), 2);
  assert.deepEqual(ids(found), ['section', 'closing']);
  // The warning is the validator's own, under the slide's number in the deck.
  assert.deepEqual(withoutCodes(found[0]?.warnings), [{ severity: 'warning', slide: 2, slot: 'title', message: 'text is 30 characters; the slot fits about 20' }]);

  const short = slideAlternatives(manifest, deckOf({ pattern: 'closing', slots: { title: 'Thank you', subtitle: 'Questions?' } }), 1);
  assert.deepEqual(short.map((alternative) => [alternative.pattern, alternative.warnings.length]), [['cover', 0], ['section', 0]]);
  // Four points in a text-only slide that holds three: offered, with the warning, after the alternatives that fit.
  const points = slideAlternatives(manifest, example, example.slides.findIndex((entry) => entry.pattern === 'list-with-photo') + 1, { layout: 'adaptive' });
  assert.match(points.at(-1)?.warnings[0]?.message ?? '', /text has 4 paragraphs; the slot fits 3/);
});

test('content that has no place makes a pattern no alternative', () => {
  const none = (pattern: string) => slideAlternatives(manifest, example, example.slides.findIndex((entry) => entry.pattern === pattern) + 1, { layout: 'adaptive' });
  // A table and a native chart fit nowhere else; the chart slide's takeaway has no slot on the diagram slide.
  for (const pattern of ['table', 'chart-native', 'chart-with-takeaway', 'quote']) assert.deepEqual(none(pattern), [], pattern);
  // Without its takeaway, the supplied chart picture and its text fit the wide diagram, and the other way round.
  const chart: SlideSpec = { pattern: 'chart-with-takeaway', slots: { title: 'T', chartGraphic: { asset: 'chart' }, chartDescription: 'What it shows.' } };
  assert.deepEqual(slideAlternatives(manifest, deckOf(chart), 1, { layout: 'adaptive' })[0]?.slide, {
    pattern: 'wide-diagram',
    slots: { title: 'T', diagramGraphic: { asset: 'chart' }, diagramCaption: 'What it shows.' },
  });
  // Figures do not become cards, and cards do not become figures.
  const stats: SlideSpec = { pattern: 'stat-cards', items: 2, slots: { title: 'T', stat1Number: '48k', stat1Caption: 'rides', stat2Number: '310', stat2Caption: 'bikes' } };
  const figures = slideAlternatives(manifest, deckOf(stats), 1, { layout: 'adaptive' });
  assert.deepEqual(ids(figures), ['text-only']);
  assert.deepEqual([figures[0]?.slide.slots.body, figures[0]?.changes], [['48k: rides', '310: bikes'], ['The 2 stats became a list of 2 points, each written as "figure: text".']]);

  assert.throws(() => slideAlternatives(manifest, deckOf(chart), 2), /the deck has 1 slides; there is no slide 2/);
  assert.throws(() => slideAlternatives(manifest, deckOf({ pattern: 'nothing', slots: {} }), 1), /uses the pattern "nothing", which the manifest does not have/);
});

test('every alternative of every example slide is a slide the renderer accepts', () => {
  const assets = { 'photo-tall': picture, 'photo-square': picture, 'photo-wide': picture, 'chart-months': picture, 'chart-wide': picture };
  const icons = (JSON.parse(readFileSync(join(root, 'examples', 'icons', 'icons.json'), 'utf8')) as Array<{ id: string }>).map((icon) => [icon.id, picture]);
  let count = 0;
  example.slides.forEach((_, index) => {
    for (const alternative of slideAlternatives(manifest, example, index + 1, { layout: 'adaptive' })) {
      const deck = { ...example, slides: [alternative.slide] };
      assert.deepEqual(validateDeck(manifest, deck, { layout: 'adaptive' }).filter((issue) => issue.severity === 'error'), [], `${index + 1} -> ${alternative.pattern}`);
      assert.ok(renderDeck(template, deck, { assets: { ...assets, ...Object.fromEntries(icons) }, layout: 'adaptive' }).pptx.length > 1000);
      count++;
    }
  });
  assert.ok(count >= 20, `${count} alternatives rendered`);
});

// --- a manifest of its own, to show the rules without a pptx -----------------------------------

const text = (name: string, maxChars: number, more: Partial<SlotDef> = {}): SlotDef => ({ name, kind: 'text', shapeId: name, maxChars, ...more }) as SlotDef;
const row = (prefix: string, count: number, slots: (n: number) => SlotDef[], item: string, more: Partial<PatternDef['repeat']> = {}): Pick<PatternDef, 'slots' | 'repeat'> => ({
  slots: [text('title', 40, { size: 30 }), ...Array.from({ length: count }, (_, index) => slots(index + 1)).flat()],
  repeat: { item, slots: slots(0).map((slot) => slot.name.replace(`${prefix}0`, `${prefix}{n}`)), shapes: Array.from({ length: count }, (_, index) => [`${prefix}${index + 1}`]), min: 2, max: count, ...more },
});
const own: TemplateManifest = {
  version: 1,
  name: 'own',
  direction: 'ltr',
  patterns: [
    { id: 'cards', sourceSlide: 1, description: 'Cards in a row.', ...row('card', 4, (n) => [text(`card${n}Heading`, 20, { size: 18 }), text(`card${n}Text`, 100, { size: 12 }), { name: `card${n}Icon`, kind: 'image', imageKind: 'icon', shapeId: `i${n}`, empty: 'remove' }], 'card') },
    { id: 'grid', sourceSlide: 2, description: 'Cards in a grid.', ...row('box', 4, (n) => [text(`box${n}Title`, 20, { size: 18 }), text(`box${n}Body`, 100, { size: 12 })], 'tile', { columns: 2 }) },
    { id: 'steps', sourceSlide: 3, description: 'Steps, one text each.', ...row('step', 4, (n) => [text(`step${n}`, 120, { size: 12 })], 'step') },
    { id: 'list', sourceSlide: 4, description: 'A list.', slots: [text('title', 40, { size: 30 }), text('points', 400, { size: 14, maxParagraphs: 5 })] },
    { id: 'figures', sourceSlide: 5, description: 'Figures.', figures: true, ...row('stat', 3, (n) => [text(`stat${n}Number`, 6, { size: 40 }), text(`stat${n}Label`, 40, { size: 12 })], 'figure') },
    { id: 'diagram', sourceSlide: 6, description: 'A diagram with a note.', slots: [text('title', 40, { size: 30 }), text('note', 200, { size: 12 }), { name: 'drawing', kind: 'image', imageKind: 'graphic', shapeId: 'g' }] },
    { id: 'note', sourceSlide: 7, description: 'A note.', slots: [text('title', 40, { size: 30 }), text('note', 200, { size: 12 })] },
  ],
};
const plain = (slides: SlideSpec[]): DeckSpec => ({ version: 1, slides });

test('a title plus items of a heading and a text can be cards, a grid, single texts or a list', () => {
  const three: SlideSpec = {
    pattern: 'cards',
    items: 3,
    slots: { title: 'T', card1Heading: 'One', card1Text: 'first.', card2Heading: 'Two', card2Text: 'second.', card3Heading: 'Three', card3Text: 'third.', card1Icon: { asset: 'a' } },
  };
  const found = slideAlternatives(own, plain([three]), 1, { layout: 'adaptive' });
  assert.deepEqual(ids(found), ['grid', 'steps', 'list']);
  assert.deepEqual(found[0]?.slide, { pattern: 'grid', items: 3, slots: { title: 'T', box1Title: 'One', box1Body: 'first.', box2Title: 'Two', box2Body: 'second.', box3Title: 'Three', box3Body: 'third.' } });
  assert.deepEqual([found[0]?.changes, found[0]?.dropped], [['The 3 cards became 3 tiles.'], ['card1Icon']]);
  assert.deepEqual(found[1]?.slide.slots, { title: 'T', step1: 'One: first.', step2: 'Two: second.', step3: 'Three: third.' });
  assert.deepEqual(found[1]?.changes, ['The 3 cards became 3 steps.', 'Each step is written as one text, "heading: text".']);
  assert.deepEqual(found[2]?.slide.slots.points, ['One: first.', 'Two: second.', 'Three: third.']);

  // And back: single texts written "heading: text" fill cards, and so does such a list.
  const steps = found[1]?.slide as SlideSpec;
  const back = slideAlternatives(own, plain([steps]), 1, { layout: 'adaptive' });
  assert.deepEqual(ids(back), ['cards', 'grid', 'list']);
  assert.deepEqual(back[0]?.slide.slots, { title: 'T', card1Heading: 'One', card1Text: 'first.', card2Heading: 'Two', card2Text: 'second.', card3Heading: 'Three', card3Text: 'third.' });
  assert.deepEqual(back[0]?.changes, ['The 3 steps became 3 cards.', 'Each card was split at the colon into a heading and a text.']);
  // The list's three points can also be three paragraphs of a plain note.
  assert.deepEqual(ids(slideAlternatives(own, plain([found[2]?.slide as SlideSpec]), 1, { layout: 'adaptive' })), ['cards', 'grid', 'steps', 'note']);

  // A card whose text has two paragraphs cannot be one point of a list without running them together.
  const long: SlideSpec = { ...three, slots: { ...three.slots, card1Text: ['first.', 'and more.'] } };
  assert.deepEqual(ids(slideAlternatives(own, plain([long]), 1, { layout: 'adaptive' })), ['grid']);
  // Five points are more than any of the rows takes; they can still be five paragraphs of a note.
  const five: SlideSpec = { pattern: 'list', slots: { title: 'T', points: ['A: a.', 'B: b.', 'C: c.', 'D: d.', 'E: e.'] } };
  assert.deepEqual(ids(slideAlternatives(own, plain([five]), 1)), ['note']);
});

test('figures stay figures, and a supplied graphic is never left out', () => {
  const stats: SlideSpec = { pattern: 'figures', items: 2, slots: { title: 'T', stat1Number: '42', stat1Label: 'things', stat2Number: '7', stat2Label: 'days' } };
  // As single texts or as a list a figure leads its label; cards and the grid would need a heading it is not.
  assert.deepEqual(ids(slideAlternatives(own, plain([stats]), 1)), ['steps', 'list']);
  // Cards have no number for the figures' slots.
  const two: SlideSpec = { pattern: 'cards', items: 2, slots: { title: 'T', card1Heading: 'One', card1Text: 'first.', card2Heading: 'Two', card2Text: 'second.' } };
  assert.ok(!ids(slideAlternatives(own, plain([two]), 1, { layout: 'adaptive' })).includes('figures'));

  const diagram: SlideSpec = { pattern: 'diagram', slots: { title: 'T', note: 'How to read it.', drawing: { asset: 'd' } } };
  assert.deepEqual(slideAlternatives(own, plain([diagram]), 1), [], 'the note slide has no place for the drawing');
  const note: SlideSpec = { pattern: 'note', slots: { title: 'T', note: 'Just words.' } };
  assert.deepEqual(ids(slideAlternatives(own, plain([note]), 1)), ['list'], 'the diagram slide needs a drawing the note does not have; the list takes the words');
});
