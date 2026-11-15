// prepareManifest on the slides and manifests that are easy to get wrong: a title that takes its
// place from the layout, cards that are not groups, two patterns with one id, a grid said to take
// more than it draws, deck fields without a value. No model is asked anywhere.
import assert from 'node:assert/strict';
import { readFileSync } from 'node:fs';
import { dirname, join } from 'node:path';
import { test } from 'node:test';
import { fileURLToPath } from 'node:url';
import { strFromU8, unzipSync } from 'fflate';
import { changeManifest, checkManifest, loadTemplate, planOutline, prepareManifest, renderDeck, sampleDeck, suggestChanges, validateDeck } from '../src/index.ts';
import type { DeckSpec, JsonRequest, LayoutMode, LlmProvider, PatternDef, TemplateManifest } from '../src/index.ts';
import { buildFixturePptx, card, fixtureManifest, fixtureWithSlide, fixtureWithTextCards, gridManifest } from './fixture.ts';

const root = join(dirname(fileURLToPath(import.meta.url)), '..');
const pptx = buildFixturePptx();
const demoPptx = readFileSync(join(root, 'templates', 'demo.pptx'));
const demoManifest = JSON.parse(readFileSync(join(root, 'templates', 'demo.manifest.json'), 'utf8')) as TemplateManifest;

const patternOf = (manifest: TemplateManifest, id: string) => manifest.patterns.find((pattern) => pattern.id === id) as PatternDef;
const namesOf = (manifest: TemplateManifest) => manifest.patterns.flatMap((pattern) => [pattern.id, ...pattern.slots.map((slot) => `${pattern.id}.${slot.name}`)]);
const [content, closing] = fixtureManifest.patterns as [PatternDef, PatternDef];

const text = (id: number, name: string, x: number, y: number, w: number, h: number, words = 'SAMPLE WORDS') =>
  `<p:sp><p:nvSpPr><p:cNvPr id="${id}" name="${name}"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr>` +
  `<p:spPr><a:xfrm><a:off x="${x}" y="${y}"/><a:ext cx="${w}" cy="${h}"/></a:xfrm></p:spPr>` +
  `<p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr sz="900"/><a:t>${words}</a:t></a:r></a:p></p:txBody></p:sp>`;
const pic = (id: number, name: string, x: number, y: number, w: number, h: number) =>
  `<p:pic><p:nvPicPr><p:cNvPr id="${id}" name="${name}"/><p:cNvPicPr/><p:nvPr/></p:nvPicPr>` +
  `<p:blipFill><a:blip r:embed="rId2"/><a:stretch><a:fillRect/></a:stretch></p:blipFill>` +
  `<p:spPr><a:xfrm><a:off x="${x}" y="${y}"/><a:ext cx="${w}" cy="${h}"/></a:xfrm></p:spPr></p:pic>`;
/** A drawn rectangle that holds no text: a text slot on it is a warning for checkManifest, and nothing to fill for the renderer. */
const rectangle = (id: number) =>
  `<p:sp><p:nvSpPr><p:cNvPr id="${id}" name="Box"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x="100000" y="4000000"/><a:ext cx="800000" cy="300000"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr></p:sp>`;
/** Slide 4 of the fixture with its title (#2) and what is given after it. */
const slide4 = (rest: string) => fixtureWithSlide(4, (xml) => `${xml.slice(0, xml.indexOf('<p:grpSp>'))}${rest}</p:spTree></p:cSld></p:sld>`);
/** Slide 4 with nothing of its own. */
const slide4Bare = (rest: string) => fixtureWithSlide(4, (xml) => `${xml.slice(0, xml.indexOf('<p:sp>'))}${rest}</p:spTree></p:cSld></p:sld>`);
const withPattern = (pattern: PatternDef): TemplateManifest => ({ ...fixtureManifest, patterns: [...fixtureManifest.patterns, pattern] });
const samples = (bytes: Uint8Array, manifest: TemplateManifest) => {
  const template = loadTemplate(bytes, manifest);
  for (const items of ['min', 'max'] as const) {
    const { deck, assets } = sampleDeck(template, { items });
    renderDeck(template, deck, { assets, layout: 'adaptive' });
  }
  return true;
};
/** Whether the planner is offered a pattern, with the fields and the layout mode a chat side would pass and no icon set. */
const offered = async (bytes: Uint8Array, manifest: TemplateManifest, id: string, layout?: LayoutMode) => {
  const requests: JsonRequest[] = [];
  const llm: LlmProvider = { async generateJson(request) { requests.push(request); throw new Error('stop'); } };
  await planOutline(loadTemplate(bytes, manifest), { prompt: 'Three things', language: 'en', fields: { closing: 'Bye' } }, { llm, ...(layout ? { layout } : {}) }).catch(() => undefined);
  return new RegExp(`\\n- ${id}: `).test(requests[0]?.prompt ?? '');
};

test('a shape that takes its place from the layout is not moved, so no variant is made that cannot be rendered', () => {
  // The usual title of a real template: a placeholder with no position of its own on the slide.
  const placeholder =
    `<p:sp><p:nvSpPr><p:cNvPr id="40" name="Title 1"/><p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr><p:nvPr><p:ph type="title"/></p:nvPr></p:nvSpPr><p:spPr/>` +
    `<p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr lang="en-US" sz="2600"/><a:t>SAMPLE TITLE HERE</a:t></a:r></a:p></p:txBody></p:sp>`;
  const bytes = slide4Bare(placeholder + text(41, 'Body', 500000, 1500000, 3500000, 2500000, 'SAMPLE BODY TEXT THAT SAYS SOMETHING') + pic(42, 'Photo', 4800000, 1200000, 3800000, 3200000));
  const built = changeManifest(loadTemplate(bytes, fixtureManifest), [
    { op: 'addPattern', slide: 4, id: 'text-photo', description: 'A title, a text and a photo.' },
    { op: 'addSlot', pattern: 'text-photo', shape: '40', role: 'title', name: 'title' },
    { op: 'addSlot', pattern: 'text-photo', shape: '41', role: 'text', name: 'body' },
    { op: 'addSlot', pattern: 'text-photo', shape: '42', role: 'photo', name: 'photo' },
  ]);
  assert.deepEqual(built.rejected, []);
  const template = loadTemplate(bytes, built.manifest);
  assert.ok(samples(bytes, built.manifest));

  // Mirroring would move the title, and it has no place to move from.
  const mirror = changeManifest(template, [{ op: 'derivePattern', pattern: 'text-photo', variant: 'mirror' }]);
  assert.match(mirror.rejected[0]?.reason ?? '', /takes its place from the slide's layout and has none of its own on the slide, so it cannot be moved/);
  assert.deepEqual(suggestChanges(template, { kinds: ['variant'] }).map((suggestion) => suggestion.variant).filter((variant) => variant !== 'without'), ['text-only']);

  // A manifest in use that was sound comes back sound, with a text-only variant whose title stays where the layout puts it.
  const prepared = prepareManifest(template, { icons: true, remove: false, fields: { closing: 'Bye' } });
  assert.deepEqual([prepared.sound, prepared.added], [true, ['text-photo-text-only']]);
  assert.equal('box' in (patternOf(prepared.manifest, 'text-photo-text-only').slots.find((slot) => slot.name === 'title') ?? {}), false);
  assert.ok(samples(bytes, prepared.manifest));
  const again = prepareManifest(loadTemplate(bytes, prepared.manifest), { icons: true, remove: false, fields: { closing: 'Bye' } });
  assert.deepEqual([again.done, again.manifest], [[], prepared.manifest]);
});

test('no variant is made of a pattern that cannot be used, and with remove: false left names the pattern', () => {
  // A manifest in use: one layout has a slot on a shape the file no longer has.
  const bare: TemplateManifest = {
    ...demoManifest,
    patterns: demoManifest.patterns
      .filter((pattern) => !/-mirrored$|^text-only$|^chart-native$/.test(pattern.id))
      .map((pattern): PatternDef => (pattern.id !== 'list-with-photo' ? pattern : { ...pattern, slots: [...pattern.slots, { name: 'gone', kind: 'text', shapeId: '777', maxChars: 10 }] })),
  };
  const prepared = prepareManifest(loadTemplate(demoPptx, bare), { icons: true, remove: false });
  assert.deepEqual(prepared.added.filter((id) => id.startsWith('list-with-photo-')), [], 'one layout that cannot be used does not become three');
  assert.ok(prepared.added.length > 0);
  assert.deepEqual(checkManifest(loadTemplate(demoPptx, prepared.manifest)).filter((issue) => issue.severity === 'error').map((issue) => `${issue.pattern}.${issue.slot}`), ['list-with-photo.gone']);
  assert.deepEqual([prepared.sound, prepared.removed], [false, []]);

  // A sample that fails without an error from checkManifest: left says which layout, not only what.
  const drawn = fixtureWithSlide(1, (xml) => xml.replace('</p:spTree>', `${rectangle(50)}</p:spTree>`));
  const odd = withPattern({ id: 'odd', sourceSlide: 1, description: 'A title, a label and a photo.', slots: [{ name: 'title', kind: 'text', shapeId: '2', maxChars: 20 }, { name: 'label', kind: 'text', shapeId: '50', maxChars: 10 }, { name: 'photo', kind: 'image', imageKind: 'photo', shapeId: '4' }] });
  const stays = prepareManifest(loadTemplate(drawn, odd), { icons: true, remove: false });
  assert.deepEqual([stays.sound, stays.removed, stays.added.filter((id) => id.startsWith('odd-'))], [false, [], []]);
  assert.equal(stays.left[0], 'The sample deck does not render: "odd" (shape has no text body to fill)');
  // Taken out, nothing was made of it first, so nothing is said of a variant.
  const out = prepareManifest(loadTemplate(drawn, odd), { icons: true });
  assert.deepEqual([out.sound, out.removed], [true, ['odd']]);
  assert.ok(!out.done.some((sentence) => sentence.includes('"odd-')), out.done.join(' | '));
});

test('without an icon set, a row found in the same run takes its pictograms along, and kept names every slot that was set', () => {
  // Cards that are not groups: a text box and a pictogram above it, each directly on the slide.
  const lefts: [number, number, number] = [6000000, 4000000, 2000000];
  const bytes = fixtureWithTextCards(lefts, lefts.map((x, index) => pic(12 + index * 10, 'Icon', x + 300000, 1500000, 400000, 400000)).join(''));
  const cards = withPattern({
    id: 'cards',
    sourceSlide: 4,
    description: 'Cards.',
    slots: [
      { name: 'title', kind: 'text', shapeId: '2', size: 26, maxChars: 30 },
      ...[1, 2, 3].flatMap((n): PatternDef['slots'] => [
        { name: `card${n}Text`, kind: 'text', shapeId: `${n}1`, size: 9, maxChars: 20 },
        { name: `card${n}Icon`, kind: 'image', imageKind: 'icon', shapeId: `${n}2` },
      ]),
    ],
  });
  const prepared = prepareManifest(loadTemplate(bytes, cards), { icons: false, variants: [], widen: false });
  const repeat = patternOf(prepared.manifest, 'cards').repeat;
  assert.deepEqual([repeat?.slots.slice().sort(), repeat?.shapes.map((item) => item.slice().sort())], [['card{n}Icon', 'card{n}Text'], [['11', '12'], ['21', '22'], ['31', '32']]]);
  assert.deepEqual(prepared.kept, [1, 2, 3].map((n) => ({ pattern: 'cards', slot: `card${n}Icon`, was: 'error' })));
  // Two cards: two pictograms, each still above its text.
  const template = loadTemplate(bytes, prepared.manifest);
  const { deck, assets } = sampleDeck(template, { items: 'min' });
  const only: DeckSpec = { ...deck, slides: deck.slides.filter((slide) => slide.pattern === 'cards') };
  assert.equal(only.slides[0]?.items, 2);
  const xml = strFromU8(unzipSync(renderDeck(template, only, { assets, layout: 'adaptive' }).pptx)['ppt/slides/slide1.xml'] as Uint8Array);
  assert.equal((xml.match(/<p:pic>/g) ?? []).length, 2);
  assert.deepEqual(prepareManifest(template, { icons: false, variants: [], widen: false }).done, []);

  // A pattern that gets variants: their copies of the slot are kept and named as well, and can be given back.
  const beside = slide4Bare(text(40, 'Title', 500000, 300000, 8000000, 600000, 'SAMPLE TITLE HERE') + pic(43, 'Icon', 500000, 1300000, 400000, 400000) + text(41, 'Body', 500000, 1900000, 3500000, 2500000, 'SAMPLE BODY TEXT THAT SAYS SOMETHING') + pic(42, 'Photo', 4800000, 1200000, 3800000, 3200000));
  const textPhoto = withPattern({
    id: 'text-photo',
    sourceSlide: 4,
    description: 'A title, a pictogram, a text and a photo.',
    slots: [
      { name: 'title', kind: 'text', shapeId: '40', size: 26, maxChars: 40 },
      { name: 'icon', kind: 'image', imageKind: 'icon', shapeId: '43' },
      { name: 'body', kind: 'text', shapeId: '41', size: 9, maxChars: 300 },
      { name: 'photo', kind: 'image', imageKind: 'photo', shapeId: '42' },
    ],
  });
  const withVariants = prepareManifest(loadTemplate(beside, textPhoto), { icons: false });
  assert.deepEqual(withVariants.added.slice().sort(), ['text-photo-mirrored', 'text-photo-text-only']);
  assert.deepEqual(withVariants.kept.map((entry) => `${entry.pattern}.${entry.slot}:${entry.was}`).sort(), ['text-photo-mirrored.icon:error', 'text-photo-text-only.icon:error', 'text-photo.icon:error']);
  const back = changeManifest(loadTemplate(beside, withVariants.manifest), withVariants.kept.map((entry) => ({ op: 'setSlot' as const, pattern: entry.pattern, slot: entry.slot, empty: entry.was })));
  assert.deepEqual(back.rejected, []);
  assert.deepEqual(back.manifest.patterns.flatMap((pattern) => pattern.slots.filter((slot) => slot.kind === 'image' && slot.imageKind === 'icon' && slot.empty === 'keep').map((slot) => `${pattern.id}.${slot.name}`)), ['content.icon'], 'only the one the fixture keeps by hand');

  // A pattern that is taken out is not named.
  const drawn = fixtureWithSlide(1, (xml) => xml.replace('</p:spTree>', `${rectangle(50)}</p:spTree>`));
  const odd = withPattern({ id: 'odd', sourceSlide: 1, description: 'A title, a label and a pictogram.', slots: [{ name: 'title', kind: 'text', shapeId: '2', maxChars: 20 }, { name: 'label', kind: 'text', shapeId: '50', maxChars: 10 }, { name: 'icon', kind: 'image', imageKind: 'icon', shapeId: '8' }] });
  const gone = prepareManifest(loadTemplate(drawn, odd), { icons: false });
  assert.deepEqual([gone.removed, gone.kept], [['odd'], []]);
});

test('an icon slot that may be left out is kept too when the chat side plans strictly, which is the default', async () => {
  // The fixture's cards with pictograms a deck may leave out.
  const optional = withPattern({
    id: 'cards',
    sourceSlide: 4,
    description: 'A row of cards.',
    slots: [
      { name: 'title', kind: 'text', shapeId: '2', size: 26, maxChars: 30 },
      ...[1, 2, 3].flatMap((n): PatternDef['slots'] => [
        { name: `card${n}Text`, kind: 'text', shapeId: `${n}1`, size: 9, maxChars: 20 },
        { name: `card${n}Icon`, kind: 'image', imageKind: 'icon', shapeId: `${n}2`, empty: 'remove' },
      ]),
    ],
    repeat: { item: 'card', slots: ['card{n}Text', 'card{n}Icon'], shapes: [['10', '13'], ['20'], ['30']], min: 2, max: 3 },
  });
  // With no icon set a strict planner must fill the slot and cannot: the layout is not offered. An adaptive one leaves it out.
  assert.deepEqual([await offered(pptx, optional, 'cards'), await offered(pptx, optional, 'cards', 'adaptive')], [false, true]);

  const strict = prepareManifest(loadTemplate(pptx, optional), { icons: false });
  assert.deepEqual(strict.kept, [1, 2, 3].map((n) => ({ pattern: 'cards', slot: `card${n}Icon`, was: 'remove' })));
  assert.equal(await offered(pptx, strict.manifest, 'cards'), true);

  const adaptive = prepareManifest(loadTemplate(pptx, optional), { icons: false, layout: 'adaptive' });
  assert.deepEqual(adaptive.kept, []);
  assert.deepEqual(patternOf(adaptive.manifest, 'cards').slots.filter((slot) => slot.kind === 'image').map((slot) => slot.empty), ['remove', 'remove', 'remove'], 'decks that left the pictogram out look as before');
  assert.equal(await offered(pptx, adaptive.manifest, 'cards', 'adaptive'), true);

  assert.throws(() => prepareManifest(loadTemplate(pptx, optional), { icons: false, layout: 'loose' as never }), /"loose" is not a layout mode; it is strict or adaptive/);
});

test('two patterns with one id both stay, whichever of them cannot be used', () => {
  // The second closing has a slot on a shape that is not there.
  const twice: TemplateManifest = { ...fixtureManifest, patterns: [content, closing, { id: 'closing', sourceSlide: 2, description: 'Another closing.', slots: [{ name: 'gone', kind: 'text', shapeId: '99', maxChars: 10 }] }] };
  for (const options of [{ icons: true, variants: [], widen: false }, { icons: true, variants: [], widen: false, remove: false }] as const) {
    let manifest = twice;
    for (let run = 1; run <= 2; run++) {
      const prepared = prepareManifest(loadTemplate(pptx, manifest), options);
      assert.deepEqual(prepared.manifest.patterns.map((pattern) => [pattern.id, pattern.slots[0]?.name]), [['content', 'title'], ['closing', 'title'], ['closing', 'gone']], `run ${run}`);
      assert.deepEqual([prepared.removed, prepared.sound], [[], false]);
      assert.ok(!prepared.done.some((sentence) => sentence.includes('was taken out')), prepared.done.join(' | '));
      assert.match(prepared.left[0] ?? '', /^2 errors could not be solved: closing another pattern has the same id; closing\.gone points at shape #99/);
      manifest = prepared.manifest;
    }
  }
  // The second one only fails its sample: no error names it, and still neither is taken out.
  const drawn = fixtureWithSlide(1, (xml) => xml.replace('</p:spTree>', `${rectangle(50)}</p:spTree>`));
  const sampled: TemplateManifest = { ...fixtureManifest, patterns: [content, closing, { id: 'closing', sourceSlide: 1, description: 'Another closing, with a label.', slots: [{ name: 'title', kind: 'text', shapeId: '2', maxChars: 20 }, { name: 'label', kind: 'text', shapeId: '50', maxChars: 10 }] }] };
  const prepared = prepareManifest(loadTemplate(drawn, sampled), { icons: true });
  assert.deepEqual(prepared.manifest.patterns.filter((pattern) => pattern.id === 'closing').map((pattern) => pattern.sourceSlide), [2, 1]);
  assert.deepEqual([prepared.removed, prepared.sound], [[], false]);
  assert.match(prepared.left[0] ?? '', /^One error could not be solved: closing another pattern has the same id\.$/);
});

test('with remove: false a grid said to take more than it draws keeps its number, and left says so', () => {
  const grid = patternOf(gridManifest, 'grid');
  const wide = withPattern({ ...grid, repeat: { ...(grid.repeat as NonNullable<typeof grid.repeat>), max: 6 } });
  const stored: DeckSpec = { version: 1, title: 'Stored', direction: 'rtl', fields: { closing: 'Bye' }, slides: [{ pattern: 'grid', items: 5, slots: { title: 'T', card1Text: 'a', card2Text: 'b', card3Text: 'c', card4Text: 'd', card5Text: 'e' } }] };
  const fine = (manifest: TemplateManifest) => validateDeck(manifest, stored).filter((issue) => issue.severity === 'error').length === 0 && renderDeck(loadTemplate(pptx, manifest), stored).pptx.length > 1000;
  assert.ok(fine(wide), 'a deck with five cards is rendered today');

  const inUse = prepareManifest(loadTemplate(pptx, wide), { icons: true, remove: false, variants: [], widen: false });
  assert.equal(patternOf(inUse.manifest, 'grid').repeat?.max, 6);
  assert.ok(fine(inUse.manifest), 'and still is');
  assert.ok(inUse.left.includes('"grid" says its grid takes up to 6 cards, and the template draws 4: a slide with more runs past what is drawn. The number was left, in case a stored presentation has that many. Set it to 4 when none does.'), inUse.left.join(' | '));
  assert.ok(!inUse.left.some((sentence) => sentence.includes('a grid cannot hold more items than it draws')), 'said once, in words for a person');
  assert.deepEqual(prepareManifest(loadTemplate(pptx, inUse.manifest), { icons: true, remove: false, variants: [], widen: false }).done, []);

  // After a draft nothing depends on it: the grid takes what it draws.
  const fresh = prepareManifest(loadTemplate(pptx, wide), { icons: true, variants: [], widen: false });
  assert.equal(patternOf(fresh.manifest, 'grid').repeat?.max, 4);
  assert.ok(!fresh.left.some((sentence) => sentence.includes('says its grid')));
});

test('a row held at its last item takes no more than it took: a further item would be a copy with the template\'s own sample in it', () => {
  // Cards as groups; the row is of the texts, and each card's pictogram is a required slot beside the row.
  const cards = withPattern({
    id: 'cards',
    sourceSlide: 4,
    description: 'A row of cards.',
    slots: [
      { name: 'title', kind: 'text', shapeId: '2', size: 26, maxChars: 30 },
      ...[1, 2, 3].flatMap((n): PatternDef['slots'] => [
        { name: `card${n}Text`, kind: 'text', shapeId: `${n}1`, size: 9, maxChars: 20 },
        { name: `card${n}Icon`, kind: 'image', imageKind: 'icon', shapeId: `${n}2` },
      ]),
    ],
    repeat: { item: 'card', slots: ['card{n}Text'], shapes: [['10'], ['20'], ['30']], min: 2, max: 3 },
  });
  const template = loadTemplate(pptx, cards);
  assert.deepEqual(suggestChanges(template, { kinds: ['row-max'] }), []);
  const prepared = prepareManifest(template, { icons: true, remove: false, variants: [], widen: false });
  const repeat = patternOf(prepared.manifest, 'cards').repeat;
  assert.deepEqual([repeat?.min, repeat?.max], [3, 3]);
  assert.equal(prepared.sound, true);
  // With the pictograms kept as designed there is nothing a copy would lack, and the row takes what the slide has room for.
  const kept: TemplateManifest = { ...cards, patterns: cards.patterns.map((pattern) => ({ ...pattern, slots: pattern.slots.map((slot) => (slot.kind === 'image' && pattern.id === 'cards' ? { ...slot, empty: 'keep' as const } : slot)) })) };
  assert.equal(patternOf(prepareManifest(loadTemplate(pptx, kept), { icons: true, remove: false, variants: [], widen: false }).manifest, 'cards').repeat?.max, 6);
});

test('a text of the whole deck inside the last item means there is no row', () => {
  // The third card holds a second text, which is a value of the deck.
  const extra = text(33, 'Department', 2000000, 3300000, 1000000, 200000, 'SAMPLE DEPARTMENT');
  const bytes = fixtureWithSlide(4, (xml) => xml.replace(card(30, 2000000), card(30, 2000000).replace('</p:grpSp>', `${extra}</p:grpSp>`)));
  const cards = withPattern({
    id: 'cards',
    sourceSlide: 4,
    description: 'Cards.',
    slots: [
      { name: 'title', kind: 'text', shapeId: '2', size: 26, maxChars: 30 },
      ...[1, 2, 3].map((n): PatternDef['slots'][number] => ({ name: `card${n}Text`, kind: 'text', shapeId: `${n}1`, size: 9, maxChars: 20 })),
      { name: 'department', kind: 'text', shapeId: '33', deckField: 'department' },
    ],
  });
  const template = loadTemplate(bytes, cards);
  assert.ok(samples(bytes, cards));
  assert.match(changeManifest(template, [{ op: 'setRepeat', pattern: 'cards' }]).rejected[0]?.reason ?? '', /has no row of like items/);
  assert.deepEqual(suggestChanges(template, { kinds: ['row'] }), []);
  for (const remove of [true, false]) {
    const prepared = prepareManifest(template, { icons: true, variants: [], widen: false, remove });
    assert.deepEqual([prepared.removed, prepared.sound, patternOf(prepared.manifest, 'cards').repeat], [[], true, undefined]);
    assert.deepEqual(prepareManifest(loadTemplate(bytes, prepared.manifest), { icons: true, variants: [], widen: false, remove }).done, []);
  }
});

test('a deck field with no value and no text in the template is asked for', () => {
  // The closing text of the template was cleared.
  const cleared = fixtureWithSlide(2, (xml) => xml.replace('SAMPLE CLOSING', ''));
  const prepared = prepareManifest(loadTemplate(cleared, fixtureManifest), { icons: true });
  assert.deepEqual(prepared.fields, {});
  assert.ok(prepared.left.includes('One text that is the same on every slide has no value yet, and the template shows none in its place: closing. A layout with it is not offered until you set one.'), prepared.left.join(' | '));
  // Given a value, nothing is asked.
  const given = prepareManifest(loadTemplate(cleared, fixtureManifest), { icons: true, fields: { closing: 'Thank you' } });
  assert.deepEqual(given.fields, { closing: 'Thank you' });
  assert.ok(!given.left.some((sentence) => sentence.includes('no value yet')));
});

test('a row that only a variant has room for is found in the same run', () => {
  // Two cards and a lead beside a photo: no room for a third card until the photo is left out.
  const bytes = fixtureWithSlide(4, (xml) =>
    `${xml.slice(0, xml.indexOf('<p:sp>'))}${text(2, 'Title', 300000, 200000, 2400000, 500000, 'SAMPLE TITLE')}${card(10, 300000)}${card(20, 1700000)}${text(60, 'Lead', 300000, 3700000, 2400000, 600000, 'SAMPLE LEAD TEXT OF WORDS')}${pic(50, 'Photo', 2900000, 1000000, 5600000, 3400000)}</p:spTree></p:cSld></p:sld>`,
  );
  const duo: TemplateManifest = {
    ...fixtureManifest,
    patterns: [
      closing,
      {
        id: 'duo',
        sourceSlide: 4,
        description: 'Two cards, a lead and a photo beside them.',
        slots: [
          { name: 'title', kind: 'text', shapeId: '2', maxChars: 30 },
          { name: 'card1Text', kind: 'text', shapeId: '11', maxChars: 20 },
          { name: 'card2Text', kind: 'text', shapeId: '21', maxChars: 20 },
          { name: 'lead', kind: 'text', shapeId: '60', maxChars: 60 },
          { name: 'photo', kind: 'image', imageKind: 'photo', shapeId: '50' },
        ],
      },
    ],
  };
  const template = loadTemplate(bytes, duo);
  assert.deepEqual(suggestChanges(template, { kinds: ['row'] }), [], 'beside the photo the two cards have no room for more');
  const prepared = prepareManifest(template, { icons: true });
  assert.ok(prepared.added.includes('duo-text-only') && prepared.sound);
  const row = patternOf(prepared.manifest, 'duo-text-only').repeat;
  assert.ok(row !== undefined && (row.max ?? 0) > 2, JSON.stringify(row));
  const after = loadTemplate(bytes, prepared.manifest);
  assert.ok(suggestChanges(after).every((suggestion) => suggestion.variant === 'without'));
  const again = prepareManifest(after, { icons: true });
  assert.deepEqual([again.done, again.manifest], [[], prepared.manifest]);
  assert.deepEqual(namesOf(prepared.manifest).filter((name) => namesOf(duo).includes(name)), namesOf(duo));
});
