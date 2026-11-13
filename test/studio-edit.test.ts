// Changing a manifest after it was drafted: by hand, one operation at a time, and by an
// instruction that a model turns into the same operations.
import assert from 'node:assert/strict';
import { test } from 'node:test';
import { changeManifest, describeManifest, editManifest, loadTemplate, renderDeck, sampleDeck, studioView, validateDeck } from '../src/index.ts';
import type { JsonRequest, LlmProvider, ManifestChange, ManifestIssue, TemplateManifest } from '../src/index.ts';
import { buildFixturePptx, cardsManifest, fixtureManifest, gridManifest, tableManifest } from './fixture.ts';

const pptx = buildFixturePptx();
const template = (manifest: TemplateManifest) => loadTemplate(pptx, manifest);

const errorsOf = (issues: ManifestIssue[]) => issues.filter((issue) => issue.severity === 'error');

/** Applies changes that are all expected to be made. */
function change(manifest: TemplateManifest, changes: ManifestChange[]) {
  const result = changeManifest(template(manifest), changes);
  assert.deepEqual(result.rejected, [], 'every change is made');
  return result;
}

test('a pattern is built by hand from a slide: slots measured, the row found, the range checked', () => {
  const before = structuredClone(fixtureManifest);
  const built = change(fixtureManifest, [
    { op: 'addPattern', slide: 4, id: 'Card Row', description: 'A row of cards.' },
    { op: 'addSlot', pattern: 'card-row', shape: '2', role: 'title' },
    ...[1, 2, 3].flatMap((n): ManifestChange[] => [
      { op: 'addSlot', pattern: 'card-row', shape: `${n}1`, role: 'text', name: `card${n} text` },
      { op: 'addSlot', pattern: 'card-row', shape: `${n}2`, role: 'icon', name: `card${n} icon` },
    ]),
    { op: 'addSlot', pattern: 'card-row', shape: '13', role: 'decoration', name: 'stray label' },
    { op: 'setRepeat', pattern: 'card-row', max: 5 },
    { op: 'setSlot', pattern: 'card-row', slot: 'card2Icon', empty: 'remove' },
  ]);
  assert.deepEqual(fixtureManifest, before, 'the manifest that was passed in is not touched');
  assert.deepEqual(errorsOf(built.issues), []);

  const pattern = built.manifest.patterns.find((candidate) => candidate.id === 'card-row');
  assert.deepEqual(pattern?.slots.map((slot) => slot.name), ['title', 'card1Text', 'card1Icon', 'card2Text', 'card2Icon', 'card3Text', 'card3Icon', 'strayLabel']);
  // Limits come from the shape, as they do in a draft.
  assert.deepEqual(pattern?.slots[1], { name: 'card1Text', kind: 'text', shapeId: '11', size: 9, maxChars: 40 });
  assert.deepEqual(pattern?.slots.at(-1), { name: 'strayLabel', kind: 'text', shapeId: '13', size: 8, maxChars: 24, empty: 'drop' });
  // One item's slot was changed; every item's slot follows, because the items of a row must match.
  assert.deepEqual(pattern?.slots.filter((slot) => slot.kind === 'image').map((slot) => slot.empty), ['remove', 'remove', 'remove']);
  // Three cards of 11% in a row drawn 55% wide hold four; five need the row to spread, which the slide has room for.
  assert.deepEqual(pattern?.repeat, { item: 'card', slots: ['card{n}Text', 'card{n}Icon'], shapes: [['10'], ['20'], ['30']], min: 2, max: 5, span: { left: 0.03, right: 0.97 } });
  assert.match(built.applied.at(-2)?.message ?? '', /found a row of 3 cards; it takes 2 to 5\. The row may now spread from 3% to 97%/);
  assert.match(built.applied.at(-1)?.message ?? '', /"card1Icon", "card2Icon", "card3Icon" are now an icon, optional/);

  // What was built is usable: a full row renders.
  const made = template(built.manifest);
  const { deck, assets } = sampleDeck(made, { items: 'max' });
  assert.equal(deck.slides.at(-1)?.items, 5);
  assert.deepEqual(validateDeck(built.manifest, deck, { assets: Object.keys(assets), layout: 'adaptive' }), []);
  assert.ok(renderDeck(made, deck, { assets, layout: 'adaptive' }).pptx.length > 1000);
});

test('a change that cannot be made is refused with a reason, and the rest are still made', () => {
  const result = changeManifest(template(cardsManifest), [
    { op: 'setRepeat', pattern: 'cards', max: 12 },
    { op: 'setSlot', pattern: 'nothing', slot: 'title', maxChars: 5 },
    { op: 'addSlot', pattern: 'cards', shape: '99', role: 'text' },
    { op: 'addSlot', pattern: 'content', shape: '4', role: 'title' },
    { op: 'addSlot', pattern: 'content', shape: '8', role: 'chart' },
    { op: 'setSlot', pattern: 'content', slot: 'title', maxColumns: 3 },
    { op: 'setSlot', pattern: 'content', slot: 'title', maxChars: 0 },
    { op: 'setSlot', pattern: 'cards', slot: 'card1Text', name: 'heading' },
    { op: 'removeSlot', pattern: 'cards', slot: 'card3Icon' },
    { op: 'setSlot', pattern: 'content', slot: 'icon', box: { x: 0.1, y: 0.1, w: 0.2, h: 0.2 } },
    { op: 'setPattern', pattern: 'cards', id: 'Closing' },
    { op: 'removeRepeat', pattern: 'content' },
    { op: 'addPattern', slide: 9, id: 'ninth' },
    { op: 'setSlot', pattern: 'content', slot: 'title', maxChars: 30 },
  ]);
  const reasons = result.rejected.map((entry) => `${entry.index + 1}. ${entry.reason}`);
  for (const expected of [
    /^1\. 12 cards do not fit side by side: each is 11% of the slide wide and the row has 94%\. The most is 8$/,
    /^2\. there is no pattern "nothing"$/,
    /^3\. slide 4 has no shape #99$/,
    /^4\. shape #4 is a picture; it can be photo, icon, graphic, chart, decoration$/,
    /^5\. shape #8 is a picture; .* \(a chart cannot replace a picture inside a group\)$/,
    /^6\. "title" is a text slot; maxColumns does not apply to it$/,
    /^7\. maxChars must be a whole number of 1 or more$/,
    /^8\. "card1Text" is part of the row of cards, so it cannot be renamed/,
    /^9\. "card3Icon" is part of the row of cards; remove the row first$/,
    /^10\. icon has a box, but its shape is inside a group$/,
    /^11\. another pattern is already called "closing"$/,
    /^12\. "content" has no row of items$/,
    /^13\. the template has 5 slides; there is no slide 9$/,
  ]) {
    assert.ok(reasons.some((reason) => expected.test(reason)), `${expected} not in:\n${reasons.join('\n')}`);
  }
  assert.equal(result.rejected.length, 13);
  assert.deepEqual(result.applied.map((entry) => entry.message), ['"content": "title" is now text, up to 30 characters, required.']);
  assert.equal(result.manifest.patterns[0]?.slots[0]?.kind === 'text' && result.manifest.patterns[0].slots[0].maxChars, 30);
});

test('budgets follow a change of size, area or table shape unless the change states one', () => {
  const sized = change(fixtureManifest, [
    { op: 'setSlot', pattern: 'content', slot: 'title', fontSize: 13 },
    { op: 'setSlot', pattern: 'content', slot: 'note', maxChars: 40 },
    { op: 'setSlot', pattern: 'content', slot: 'note', box: { x: 0.1, y: 0.5, w: 0.6, h: 0.2 } },
  ]);
  const [title, , , note] = sized.manifest.patterns[0]?.slots ?? [];
  // Half the font size holds four times the text: 20 characters at 26 pt become 80 at 13 pt.
  assert.deepEqual(title, { name: 'title', kind: 'text', shapeId: '2', maxChars: 80, styleRun: 'last', fontSize: 13 });
  // The note's box was 33% by 10% of the slide; at 60% by 20% it holds 3.8 times as much.
  assert.equal(note?.kind === 'text' && note.maxChars, 151);
  assert.match(sized.applied[0]?.message ?? '', /maxChars follows at 80/);

  const table = change(tableManifest, [
    { op: 'setSlot', pattern: 'table', slot: 'table', bodyText: { size: 18 } },
    { op: 'setSlot', pattern: 'table', slot: 'table', headerText: { size: 12 }, maxColumns: 2, maxRows: 5 },
  ]);
  const slot = table.manifest.patterns.at(-1)?.slots[1];
  // Two columns of 236 pt hold 24 characters of 18 pt text a line; six rows leave room for one line each.
  assert.deepEqual(slot?.kind === 'table' && [slot.maxColumns, slot.maxRows, slot.maxCellChars], [2, 5, 21]);
  // A style is changed setting by setting: the header keeps its colour and weight.
  assert.deepEqual(slot?.kind === 'table' && slot.headerText, { size: 12, bold: true, color: 'FFFFFF', font: 'Assistant' });

  const stated = change(tableManifest, [{ op: 'setSlot', pattern: 'table', slot: 'table', maxColumns: 2, maxCellChars: 15 }]);
  const kept = stated.manifest.patterns.at(-1)?.slots[1];
  assert.equal(kept?.kind === 'table' && kept.maxCellChars, 15, 'a stated budget is not second-guessed');
});

test('patterns are copied, renamed and removed; a slot changes role; any shape can be taken away', () => {
  const result = change(gridManifest, [
    { op: 'copyPattern', pattern: 'content', id: 'Text Only', description: 'A title and text.' },
    { op: 'addSlot', pattern: 'text-only', shape: '4', role: 'decoration' },
    { op: 'addSlot', pattern: 'text-only', shape: '7', role: 'decoration', name: 'icon group' },
    { op: 'removeSlot', pattern: 'text-only', slot: 'icon' },
    { op: 'addSlot', pattern: 'content', shape: '4', role: 'chart' },
    { op: 'addSlot', pattern: 'content', shape: '6', role: 'fixed' },
    { op: 'setPattern', pattern: 'grid', id: 'Card Grid', figures: true },
    { op: 'setRepeat', pattern: 'card-grid', min: 2 },
    { op: 'removePattern', pattern: 'closing' },
    { op: 'setManifest', name: 'studio', direction: 'ltr', alignOnDirectionChange: 'keep' },
  ]);
  const { manifest } = result;
  assert.deepEqual(manifest.patterns.map((pattern) => pattern.id), ['content', 'text-only', 'card-grid']);
  assert.deepEqual([manifest.name, manifest.direction, manifest.alignOnDirectionChange], ['studio', 'ltr', 'keep']);
  assert.deepEqual(errorsOf(result.issues), [], 'a slot that only removes its shape may point at a group');

  const textOnly = manifest.patterns[1];
  assert.deepEqual(textOnly?.slots.find((slot) => slot.shapeId === '4'), { name: 'photo', kind: 'image', imageKind: 'graphic', shapeId: '4', empty: 'drop' }, 'the replaced slot keeps its name');
  assert.deepEqual(textOnly?.slots.at(-1), { name: 'iconGroup', kind: 'text', shapeId: '7', empty: 'drop' });
  // The copy changed; the pattern it was copied from did not, apart from what was asked of it.
  assert.deepEqual(manifest.patterns[0]?.slots.map((slot) => `${slot.name}:${slot.kind}`), ['title:text', 'body:text', 'photo:chart', 'icon:image']);
  assert.equal(manifest.patterns[2]?.figures, true);
  assert.deepEqual([manifest.patterns[2]?.repeat?.min, manifest.patterns[2]?.repeat?.max], [2, 4]);
  assert.match(changeManifest(template(manifest), [{ op: 'setRepeat', pattern: 'card-grid', max: 6 }]).rejected[0]?.reason ?? '', /a grid holds at most the 4 cards it draws/);

  // The text-only pattern renders without the picture and without the group.
  const made = template(manifest);
  const rendered = renderDeck(made, { version: 1, slides: [{ pattern: 'text-only', slots: { title: 'T', body: ['B'], note: 'N' } }] });
  assert.ok(rendered.pptx.length > 1000);
});

test('a header text known from other patterns is filled in for a pattern added by hand', () => {
  // Slide 2 says "SAMPLE CLOSING", which the closing pattern takes from the deck.
  const result = change(fixtureManifest, [{ op: 'addPattern', slide: 2, id: 'closing-again' }]);
  assert.deepEqual(result.manifest.patterns.at(-1)?.slots, [{ name: 'closing', kind: 'text', shapeId: '2', size: 52, maxChars: 14, deckField: 'closing' }]);
  assert.match(result.applied[0]?.message ?? '', /with the deck's closing\. It has no content slots yet\./);
});

test('studioView joins the slide and the manifest for a screen that edits by hand', () => {
  const view = studioView(template(cardsManifest));
  assert.equal(Math.round(view.aspectRatio * 1000), 1778);
  assert.deepEqual(view.slides.map((slide) => slide.patterns), [['content'], ['closing'], [], ['cards'], []]);

  const shapes = new Map(view.slides[3]?.shapes.map((shape) => [shape.id, shape]));
  assert.deepEqual(shapes.get('11'), {
    id: '11',
    kind: 'text',
    box: { x: 0.6562, y: 0.5444, w: 0.1094, h: 0.0778 },
    frame: false,
    roles: ['title', 'subtitle', 'text', 'list', 'number', 'caption', 'decoration'],
    slots: [{ pattern: 'cards', name: 'card1Text', kind: 'text', empty: 'error', summary: 'text, up to 20 characters, required' }],
    groupId: '10',
    text: 'SAMPLE CARD',
    fontSize: 9,
  });
  assert.deepEqual(shapes.get('12')?.slots, [{ pattern: 'cards', name: 'card1Icon', kind: 'image', empty: 'keep', summary: 'an icon, kept as designed' }]);
  assert.deepEqual([shapes.get('12')?.roles, shapes.get('10')?.roles], [['photo', 'icon', 'graphic', 'chart', 'decoration'], ['decoration']]);
  assert.deepEqual(shapes.get('13')?.slots, [], 'a shape that is not a slot has none');
});

test('editManifest has a model turn an instruction into checked changes, and lets it correct a refusal', async () => {
  const requests: JsonRequest[] = [];
  const llm: LlmProvider = {
    async generateJson(request) {
      requests.push(request);
      if (requests.length === 1) {
        return {
          changes: [
            { op: 'setRepeat', pattern: 'cards', shape: '', slot: '', values: [{ property: 'max', value: '12' }] },
            // Values an operation does not read are ignored, and a shape may come with its #.
            { op: 'addSlot', pattern: 'cards', shape: '#13', slot: '', values: [{ property: 'role', value: 'decoration' }, { property: 'name', value: 'stray label' }, { property: 'maxChars', value: '3' }] },
            { op: 'setSlot', pattern: 'content', shape: '', slot: 'title', values: [{ property: 'maxChars', value: '28' }] },
            { op: 'explode', pattern: 'cards', shape: '', slot: '', values: [] },
          ],
          note: '',
        };
      }
      return { changes: [{ op: 'setRepeat', pattern: 'cards', shape: '', slot: '', values: [{ property: 'max', value: '8' }] }], note: 'A row holds eight cards at most.' };
    },
  };
  const result = await editManifest(template(cardsManifest), 'Let the cards go up to 12, drop the stray label, and give the title 28 characters.', { llm });

  assert.equal(requests.length, 2);
  // The model reads the manifest next to the measured shapes, including shapes that are not slots.
  const first = requests[0]?.prompt ?? '';
  assert.match(first, /Pattern "cards", slide 4\. A row of cards\.\n  row of 3 cards \(card\{n\}Text, card\{n\}Icon\); takes 1 to 5/);
  assert.match(first, /    card1Text: text, up to 20 characters, required \| #11 text at x66 y54, 11 wide 8 high, in group #10 \| 9 pt/);
  assert.match(first, /other shapes on the slide, which stay as designed:\n    #13 text at x67 y70, 9 wide 4 high \| 8 pt \| 1 paragraph, 5 characters: "first"/);
  assert.match(first, /Slides without a pattern: 3, 5\./);
  assert.match(first, /The person asks:\nLet the cards go up to 12/);
  // The second call says what was refused and why.
  const second = requests[1]?.prompt ?? '';
  assert.match(second, /It refused these, and nothing of them was done:\n1\. 12 cards do not fit side by side: .* The most is 8/);
  assert.match(second, /The software made these:\n2\. "cards": shape #13 is a new slot, "strayLabel", a decoration that is always removed\./);

  assert.deepEqual(result.applied.map((entry) => entry.change.op), ['addSlot', 'setSlot', 'setRepeat']);
  assert.deepEqual(result.rejected, []);
  assert.equal(result.note, 'A row holds eight cards at most.');
  assert.deepEqual(errorsOf(result.issues), []);
  const cards = result.manifest.patterns.find((pattern) => pattern.id === 'cards');
  assert.equal(cards?.repeat?.max, 8);
  assert.equal(cards?.slots.at(-1)?.empty, 'drop');
  assert.equal(cardsManifest.patterns[2]?.repeat?.max, 5, 'the manifest that was passed in is not touched');
});

test('editManifest keeps a refusal the model does not correct', async () => {
  let calls = 0;
  const llm: LlmProvider = {
    async generateJson() {
      calls++;
      return calls === 1
        ? { changes: [{ op: 'removeSlot', pattern: 'cards', shape: '', slot: 'card1Icon', values: [] }], note: '' }
        : { changes: [], note: 'The icons belong to the row of cards.' };
    },
  };
  const result = await editManifest(template(cardsManifest), 'Remove the first icon.', { llm });
  assert.equal(calls, 2);
  assert.deepEqual(result.applied, []);
  assert.match(result.rejected[0]?.reason ?? '', /is part of the row of cards/);
  assert.equal(result.note, 'The icons belong to the row of cards.');
  assert.deepEqual(result.manifest, cardsManifest);
});

test('describeManifest reads like the manifest: slots with limits, rows with their range', () => {
  const text = describeManifest(template(gridManifest));
  assert.match(text, /^Template "fixture", written in he, reads right to left; mirrors alignment/);
  assert.match(text, /Pattern "grid", slide 5\. A grid of cards\.\n  grid of 4 cards, 2 to a line \(card\{n\}Text\); takes 1 to 4/);
  assert.match(text, /    title: text from the deck's "closing", required \| #2 text/);
});
