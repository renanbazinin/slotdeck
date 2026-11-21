// How many items a row takes: the arithmetic on its own, on slides described as plain data, and
// then what it promises against what the renderer draws.
import assert from 'node:assert/strict';
import { test } from 'node:test';
import { PptxPackage, changeManifest, checkManifest, loadTemplate, renderDeck, sampleDeck, suggestChanges } from '../src/index.ts';
import type { DeckSpec, TemplateManifest } from '../src/index.ts';
import { NS, childElement, descendants, parseXml, path, rootElement } from '../src/pptx/xml.ts';
import type { Element } from '../src/pptx/xml.ts';
import type { ShapeInfo, SlideInfo } from '../src/template/inspect.ts';
import type { RepeatDef } from '../src/template/manifest.ts';
import { movedRowSlots, rowCapacity, rowComfort, standsOut } from '../src/template/rows.ts';
import { buildFixturePptx, card, cardsManifest, fixtureManifest, fixtureWithSlide, fixtureWithTextCards, gridManifest } from './fixture.ts';
import { applyRepeat } from '../src/render/repeat.ts';
import { indexShapes } from '../src/render/render.ts';

/** A slide with a title and one text box per item, given as left edge and width in fractions of the slide. */
function slideOf(items: Array<[x: number, w: number]>, more: Array<Partial<ShapeInfo> & { id: string }> = [], titleX = 0.05): SlideInfo {
  const shape = (id: string, x: number, y: number, w: number, h: number, extra: Partial<ShapeInfo> = {}): ShapeInfo => ({ id, name: id, kind: 'text', box: { x, y, w, h }, ...extra }) as ShapeInfo;
  return {
    number: 1,
    part: 'ppt/slides/slide1.xml',
    layout: '',
    shapes: [shape('1', titleX, 0.05, 0.4, 0.1), ...items.map(([x, w], index) => shape(String(10 + index), x, 0.4, w, 0.3)), ...more.map((extra) => shape(extra.id, 0, 0, 0, 0, extra))],
  } as SlideInfo;
}
const rowOf = (count: number, more: Partial<RepeatDef> = {}): RepeatDef => ({ item: 'card', slots: ['card{n}Text'], shapes: Array.from({ length: count }, (_, index) => [String(10 + index)]), min: 2, max: count, ...more });

test('a row of like items takes what fits with a third of the drawn spacing, and says the span it needs', () => {
  // Four cards 12% wide, 4% apart, drawn from 20% to 80%, with free room from 5% to 95%.
  const slide = slideOf([[0.2, 0.12], [0.36, 0.12], [0.52, 0.12], [0.68, 0.12]]);
  assert.deepEqual(rowComfort(slide, rowOf(4)), { most: 6, span: { left: 0.05, right: 0.95 } });
  assert.equal(rowCapacity(slide, rowOf(4)), 7, 'seven fit when they nearly touch, which a person may ask for');
  // Taken, the answer is final: asking again gives the same number and nothing more to do.
  assert.deepEqual(rowComfort(slide, rowOf(4, { max: 6, span: { left: 0.05, right: 0.95 } })), { most: 6 });
  // A grid takes what it draws, and so does a row of one.
  assert.deepEqual(rowComfort(slide, rowOf(4, { columns: 2 })), { most: 4 });
  assert.deepEqual(rowComfort(slideOf([[0.2, 0.12]]), rowOf(1)), { most: 1 });
});

test('the count is made on the span that is returned, so one ask is final when the room is not a round number', () => {
  // The title sits 3.28% from the edge, which is the margin the row may spread to.
  const slide = slideOf([[0.4, 0.05], [0.452, 0.05]], [], 0.0328);
  const first = rowComfort(slide, rowOf(2));
  assert.ok(first.span && first.most > 2, JSON.stringify(first));
  assert.deepEqual(first.span, { left: 0.032, right: 0.968 }, 'rounded outward to the precision spans are stored at');
  assert.deepEqual(rowComfort(slide, rowOf(2, { max: first.most, span: first.span })), { most: first.most });
  // A span a person set that sticks out on one side is kept, and counted with the room on the other.
  const lopsided = rowComfort(slide, rowOf(2, { span: { left: 0.01, right: 0.6 } }));
  assert.deepEqual(lopsided.span, { left: 0.01, right: 0.968 });
  assert.deepEqual(rowComfort(slide, rowOf(2, { max: lopsided.most, span: lopsided.span })), { most: lopsided.most });
});

test('a row whose items differ in width, or sit closer than one is wide, takes what it draws', () => {
  // The renderer steps further items by the distance from the first to the second, and copies the
  // last. Here the last is wider than that step leaves room for, so a fourth would cover the third.
  const uneven = slideOf([[0.2, 0.11], [0.33, 0.11], [0.46, 0.14]]);
  assert.deepEqual(rowComfort(uneven, rowOf(3)), { most: 3 });
  assert.equal(rowCapacity(uneven, rowOf(3)), 3);
  // Equal widths with a first gap far smaller than the others: the step is the small one.
  const bunched = slideOf([[0.2, 0.1], [0.305, 0.1], [0.6, 0.1]]);
  assert.deepEqual(rowComfort(bunched, rowOf(3)), { most: 3 });
  // Items that touch get nothing more by themselves. A person may still ask: a fourth is drawn
  // touching the third, as the template draws them.
  const touching = slideOf([[0.3, 0.1], [0.4, 0.1], [0.5, 0.1]]);
  assert.deepEqual(rowComfort(touching, rowOf(3)), { most: 3 });
  assert.ok(rowCapacity(touching, rowOf(3)) > 3);
  // Positions are read to four decimals, so cards that touch can measure a step one ten-thousandth
  // short of their width. They are still cards that touch, not cards that overlap.
  assert.ok(rowCapacity(slideOf([[0.1278, 0.0917], [0.2194, 0.0917], [0.3111, 0.0917]]), rowOf(3)) > 3);
  // An item made of two shapes has its width from four rounded values, and can come out a quarter of a thousandth wide.
  assert.ok(rowCapacity(slideOf([[0.1, 0.1153], [0.21505, 0.1153], [0.3301, 0.1153]]), rowOf(3)) > 3);
  // Cards that really lie on each other, by 2% of the slide, are another matter.
  assert.equal(rowCapacity(slideOf([[0.2, 0.1], [0.28, 0.1], [0.36, 0.1]]), rowOf(3)), 3);
  // A last item a little wider than the others: the step from the first to the second leaves
  // only a quarter of the drawn space after a copy of it, so the promise of a third cannot be kept.
  const slightly = slideOf([[0.2187, 0.1094], [0.35, 0.1094], [0.4812, 0.1258]]);
  assert.deepEqual(rowComfort(slightly, rowOf(3)), { most: 3 });
  // With widths that differ by less than the spacing can absorb, the row still takes more.
  const barely = slideOf([[0.2, 0.1], [0.34, 0.1], [0.48, 0.102]]);
  assert.ok(rowComfort(barely, rowOf(3)).most > 3);
});

test('a row does not spread past what it is drawn on, nor past a line through it', () => {
  const cards: Array<[number, number]> = [[0.3, 0.1], [0.45, 0.1], [0.6, 0.1]];
  assert.deepEqual(rowComfort(slideOf(cards), rowOf(3)).span, { left: 0.05, right: 0.95 }, 'free to the margins without a panel');
  // A tinted panel behind the row, from 25% to 75%: the row stays 2% inside its edges.
  const onPanel = slideOf(cards, [{ id: '90', kind: 'shape', box: { x: 0.25, y: 0.35, w: 0.5, h: 0.4 } }]);
  assert.deepEqual(rowComfort(onPanel, rowOf(3)), { most: 4, span: { left: 0.27, right: 0.73 } });
  // Arrows between the items stay where they are. Without them four of these cards would fit in
  // the same extent; with them a fourth would be drawn across an arrow, so the row takes three.
  const spaced: Array<[number, number]> = [[0.2, 0.1], [0.4, 0.1], [0.6, 0.1]];
  assert.equal(rowComfort(slideOf(spaced), rowOf(3)).most > 3, true);
  const arrows = slideOf(spaced, [{ id: '91', kind: 'line', box: { x: 0.31, y: 0.55, w: 0.08, h: 0 } }, { id: '95', kind: 'line', box: { x: 0.51, y: 0.55, w: 0.08, h: 0 } }]);
  assert.deepEqual(rowComfort(arrows, rowOf(3)), { most: 3 });
  assert.equal(rowCapacity(arrows, rowOf(3)), 3);
  // So does a highlight behind some of the items.
  assert.deepEqual(rowComfort(slideOf(spaced, [{ id: '96', kind: 'shape', box: { x: 0.18, y: 0.38, w: 0.34, h: 0.34 } }]), rowOf(3)), { most: 3 });
  // An arrow the pattern takes off every slide is in nobody's way, also when it is drawn as a group of two shapes.
  const grouped = slideOf(spaced, [{ id: '60', kind: 'group', box: { x: 0.31, y: 0.5, w: 0.08, h: 0.1 } }, { id: '61', kind: 'shape', box: { x: 0.31, y: 0.53, w: 0.06, h: 0.04 }, groupId: '60' }, { id: '62', kind: 'shape', box: { x: 0.37, y: 0.5, w: 0.02, h: 0.1 }, groupId: '60' }]);
  assert.deepEqual(rowComfort(grouped, rowOf(3)), { most: 3 });
  assert.ok(rowComfort(grouped, rowOf(3), new Set(['60'])).most > 3);
  // A rule along the row's top edge is not among the items, and pins nothing.
  assert.equal(rowComfort(slideOf(spaced, [{ id: '97', kind: 'line', box: { x: 0.2, y: 0.4, w: 0.5, h: 0 } }]), rowOf(3)).most > 3, true);
  // A neighbour whose box reaches a little into the last card is beside the row: the row stops
  // there and is free on the other side.
  const neighbour = slideOf(spaced, [{ id: '98', kind: 'picture', box: { x: 0.69, y: 0.3, w: 0.25, h: 0.5 } }]);
  assert.deepEqual(rowComfort(neighbour, rowOf(3)).span, { left: 0.05, right: 0.7 });
  // A background from edge to edge limits nothing, and neither does a heading that only brushes the row's height.
  const onBackground = slideOf(cards, [{ id: '92', kind: 'picture', box: { x: 0, y: 0, w: 1, h: 1 } }, { id: '93', kind: 'text', box: { x: 0.1, y: 0.2, w: 0.8, h: 0.21 } }]);
  assert.deepEqual(rowComfort(onBackground, rowOf(3)).span, { left: 0.05, right: 0.95 });
  // A shape beside the row limits the room on that side, as before.
  const beside = slideOf(cards, [{ id: '94', kind: 'picture', box: { x: 0.8, y: 0.3, w: 0.15, h: 0.5 } }]);
  assert.deepEqual(rowComfort(beside, rowOf(3)).span, { left: 0.05, right: 0.78 });
});

test('a span counts only for a row on the slide, which is the only row the renderer spreads', () => {
  // The same three cards, as children of one group.
  const grouped = slideOf([[0.3, 0.1], [0.45, 0.1], [0.6, 0.1]]);
  grouped.shapes.push({ id: '50', name: 'group', kind: 'group', box: { x: 0.3, y: 0.4, w: 0.4, h: 0.3 } } as ShapeInfo);
  for (const shape of grouped.shapes) if (['10', '11', '12'].includes(shape.id)) shape.groupId = '50';
  const stale = rowOf(3, { span: { left: 0.05, right: 0.95 } });
  assert.deepEqual(rowComfort(grouped, stale), { most: 3 }, 'the span is not counted and none is proposed');
  assert.equal(rowCapacity(grouped, stale), 3);
});

// --- against the renderer ----------------------------------------------------------------

const pptx = buildFixturePptx();

/** Left edge and width, in EMU, of every card group on a rendered slide, left to right. */
function cardFrames(bytes: Uint8Array): Array<{ x: number; w: number }> {
  const slide = rootElement(parseXml(PptxPackage.load(bytes).readText('ppt/slides/slide1.xml')));
  return descendants(slide, NS.p, 'grpSp')
    .map((group) => path(group as Element, [NS.p, 'grpSpPr'], [NS.a, 'xfrm']) as Element)
    .map((xfrm) => ({ x: Number(childElement(xfrm, NS.a, 'off')?.getAttribute('x')), w: Number(childElement(xfrm, NS.a, 'ext')?.getAttribute('cx')) }))
    .sort((a, b) => a.x - b.x);
}

const plain: TemplateManifest = {
  ...fixtureManifest,
  patterns: [
    {
      id: 'cards',
      sourceSlide: 4,
      description: 'A row of cards.',
      slots: [{ name: 'title', kind: 'text', shapeId: '2', maxChars: 30 }, ...[1, 2, 3].map((n): TemplateManifest['patterns'][number]['slots'][number] => ({ name: `card${n}Text`, kind: 'text', shapeId: `${n}1`, maxChars: 20 }))],
      repeat: { item: 'card', slots: ['card{n}Text'], shapes: [['10'], ['20'], ['30']], min: 2, max: 3 },
    },
  ],
};
const deckOf = (items: number): DeckSpec => ({ version: 1, direction: 'ltr', slides: [{ pattern: 'cards', items, slots: { title: 'T', ...Object.fromEntries(Array.from({ length: items }, (_, index) => [`card${index + 1}Text`, `Card ${index + 1}`])) } }] });

test('what row-max promises is what renders: every number up to it keeps a third of the drawn spacing', () => {
  // The fixture draws three cards 1,000,000 EMU wide and 1,000,000 apart.
  const template = loadTemplate(pptx, plain);
  const [more] = suggestChanges(template, { kinds: ['row-max'] });
  assert.deepEqual(more?.changes, [{ op: 'setRepeat', pattern: 'cards', max: 6, span: { left: 0.03, right: 0.97 } }]);
  const taken = changeManifest(template, more?.changes ?? []);
  assert.deepEqual(taken.rejected, []);
  const after = loadTemplate(pptx, taken.manifest);
  for (let items = 2; items <= 6; items++) {
    const frames = cardFrames(renderDeck(after, deckOf(items)).pptx);
    assert.equal(frames.length, items);
    const gaps = frames.slice(1).map((frame, index) => frame.x - ((frames[index] as { x: number; w: number }).x + (frames[index] as { x: number; w: number }).w));
    assert.ok(gaps.every((gap) => gap >= 1000000 / 3 - 1), `${items} cards: gaps ${gaps.join(', ')}`);
    assert.ok((frames[0]?.x ?? 0) >= 0.03 * 9144000 - 1 && (frames.at(-1)?.x ?? 0) + 1000000 <= 0.97 * 9144000 + 1, `${items} cards stay inside the span`);
  }
  assert.deepEqual(suggestChanges(after, { kinds: ['row-max'] }), [], 'taken once, there is nothing more to take');
});

test('a row drawn on a panel is given no more than the panel holds, and renders inside it', () => {
  // A rounded panel behind the three cards, from 6.6% to 93.4% of the slide.
  const panel = `<p:sp><p:nvSpPr><p:cNvPr id="90" name="Panel"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x="600000" y="1900000"/><a:ext cx="7944000" cy="1700000"/></a:xfrm><a:prstGeom prst="roundRect"><a:avLst/></a:prstGeom></p:spPr></p:sp>`;
  const onPanel = fixtureWithSlide(4, (xml) => xml.replace('<p:grpSp><p:nvGrpSpPr><p:cNvPr id="10"', `${panel}<p:grpSp><p:nvGrpSpPr><p:cNvPr id="10"`));
  const template = loadTemplate(onPanel, plain);
  const [more] = suggestChanges(template, { kinds: ['row-max'] });
  assert.deepEqual(more?.changes, [{ op: 'setRepeat', pattern: 'cards', max: 5, span: { left: 0.085, right: 0.915 } }], 'five, not the six the bare slide has room for');
  const after = loadTemplate(onPanel, changeManifest(template, more?.changes ?? []).manifest);
  const frames = cardFrames(renderDeck(after, deckOf(5)).pptx);
  assert.ok((frames[0]?.x ?? 0) >= 600000 && (frames.at(-1)?.x ?? 0) + 1000000 <= 600000 + 7944000, `cards at ${frames.map((frame) => frame.x).join(', ')}`);
});

test('setRepeat refuses a span the drawn row does not fit in, and a span for a row inside a group', () => {
  const template = loadTemplate(pptx, cardsManifest);
  const narrow = changeManifest(template, [{ op: 'setRepeat', pattern: 'cards', span: { left: 0.4, right: 0.6 } }]);
  assert.match(narrow.rejected[0]?.reason ?? '', /cards do not fit side by side: each is 11% of the slide wide and the row has 20%\. The most is 1/);
  // The same span is fine for a row that is told to take the one card that fits in it.
  const one = changeManifest(template, [{ op: 'setRepeat', pattern: 'cards', min: 1, max: 1, span: { left: 0.4, right: 0.6 } }]);
  assert.match(one.applied[0]?.message ?? '', /takes exactly 1\. The row may now spread from 40% to 60% of the slide's width\./);
  // What the template draws always fits in the width it is drawn in, whatever else is set.
  assert.deepEqual(changeManifest(loadTemplate(pptx, plain), [{ op: 'setRepeat', pattern: 'cards', widen: 2 }]).rejected, []);

  // Items that are children of groups are not laid out between the limits of a span.
  const inner: TemplateManifest = { ...plain, patterns: [{ ...(plain.patterns[0] as TemplateManifest['patterns'][number]), repeat: { item: 'card', slots: ['card{n}Text'], shapes: [['11'], ['21'], ['31']], min: 2, max: 3 } }] };
  const grouped = changeManifest(loadTemplate(pptx, inner), [{ op: 'setRepeat', pattern: 'cards', span: { left: 0.05, right: 0.95 } }]);
  assert.match(grouped.rejected[0]?.reason ?? '', /the cards sit inside a group, so they stay where the group draws them and a span would change nothing/);
  assert.deepEqual(changeManifest(loadTemplate(pptx, inner), [{ op: 'setRepeat', pattern: 'cards', span: null }]).rejected, [], 'a span can still be taken away');
});

/** The fixture's card slide with its three cards at other places, first in reading order last. */
const cardsAt = (first: number, second: number, third: number, wrap = false): Uint8Array =>
  fixtureWithSlide(4, (xml) => {
    const moved = xml.replace(card(10, 6000000), card(10, third)).replace(card(20, 4000000), card(20, second)).replace(card(30, 2000000), card(30, first));
    if (!wrap) return moved;
    // The three cards as children of one group, which is how a designer often draws a row.
    const start = moved.indexOf('<p:grpSp><p:nvGrpSpPr><p:cNvPr id="10"');
    const end = moved.indexOf('</p:spTree>');
    const group = `<p:grpSp><p:nvGrpSpPr><p:cNvPr id="50" name="Row"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr><a:xfrm><a:off x="${first}" y="2000000"/><a:ext cx="${third + 1000000 - first}" cy="1800000"/><a:chOff x="${first}" y="2000000"/><a:chExt cx="${third + 1000000 - first}" cy="1800000"/></a:xfrm></p:grpSpPr>`;
    return `${moved.slice(0, start)}${group}${moved.slice(start, end)}</p:grpSp>${moved.slice(end)}`;
  });

test('setRepeat counts the fit when it is asked about the number or the span, and never refuses the row as drawn', () => {
  const template = loadTemplate(pptx, plain);
  // A span the drawn three do not fit in is refused, with or without the number.
  for (const change of [{ op: 'setRepeat', pattern: 'cards', span: { left: 0.4, right: 0.6 } }, { op: 'setRepeat', pattern: 'cards', span: { left: 0.4, right: 0.6 }, max: 3 }] as const) {
    assert.match(changeManifest(template, [change]).rejected[0]?.reason ?? '', /3 cards do not fit side by side: each is 11% of the slide wide and the row has 20%\. The most is 1/);
  }

  // Cards drawn 1% of a card apart: closer than the fit allows for new items, and still the row as drawn.
  const tight = loadTemplate(cardsAt(2000000, 3010000, 4020000), plain);
  assert.deepEqual(changeManifest(tight, [{ op: 'setRepeat', pattern: 'cards', widen: 2 }]).rejected, []);
  assert.deepEqual(changeManifest(tight, [{ op: 'setRepeat', pattern: 'cards', max: 3 }]).rejected, []);
  assert.deepEqual(suggestChanges(tight, { kinds: ['row-max'] }), [], 'nothing more is offered for cards that nearly touch');

  // Cards that touch: a person may ask for a fourth, which is drawn touching like the others.
  const touching = loadTemplate(cardsAt(2000000, 3000000, 4000000), plain);
  const four = changeManifest(touching, [{ op: 'setRepeat', pattern: 'cards', max: 4 }]);
  assert.deepEqual(four.rejected, []);
  const frames = cardFrames(renderDeck(loadTemplate(cardsAt(2000000, 3000000, 4000000), four.manifest), deckOf(4)).pptx);
  assert.deepEqual(frames.slice(1).map((frame, index) => frame.x - ((frames[index] as { x: number; w: number }).x + 1000000)), [0, 0, 0]);
});

test('taking a span away takes away the room it gave', () => {
  // Six cards fit only with the row spread from 3% to 97%.
  const template = loadTemplate(pptx, plain);
  const wide = changeManifest(template, [{ op: 'setRepeat', pattern: 'cards', max: 6, span: { left: 0.03, right: 0.97 } }]);
  assert.deepEqual(wide.rejected, []);
  const back = changeManifest(loadTemplate(pptx, wide.manifest), [{ op: 'setRepeat', pattern: 'cards', span: null }]);
  assert.deepEqual(back.rejected, [], 'a span can always be taken away');
  // Four of these cards fit in the width the three are drawn in, when they sit close.
  assert.match(back.applied[0]?.message ?? '', /takes 2 to 4\. Without the span no more than 4 fit side by side\./);
  assert.deepEqual(back.manifest.patterns[0]?.repeat, { item: 'card', slots: ['card{n}Text'], shapes: [['10'], ['20'], ['30']], min: 2, max: 4 });
  const frames = cardFrames(renderDeck(loadTemplate(pptx, back.manifest), deckOf(4)).pptx);
  assert.ok(frames.slice(1).every((frame, index) => frame.x >= (frames[index] as { x: number; w: number }).x + 1000000), `four cards at ${frames.map((frame) => frame.x).join(', ')} do not overlap`);
  // A row drawn as a fan, each card on a fifth of the next, was given five with a span by 0.2.0.
  // Without the span it goes back to the three it draws: a further card would lie on the last.
  const fan = cardsAt(2000000, 2800000, 3600000);
  const stored: TemplateManifest = { ...plain, patterns: [{ ...(plain.patterns[0] as TemplateManifest['patterns'][number]), repeat: { item: 'card', slots: ['card{n}Text'], shapes: [['10'], ['20'], ['30']], min: 2, max: 5, span: { left: 0.03, right: 0.97 } } }] };
  const unfanned = changeManifest(loadTemplate(fan, stored), [{ op: 'setRepeat', pattern: 'cards', span: null }]);
  assert.match(unfanned.applied[0]?.message ?? '', /takes 2 to 3\. Without the span no more than 3 fit side by side\./);
  // Asked to keep six without the span, it says no.
  assert.match(changeManifest(loadTemplate(pptx, wide.manifest), [{ op: 'setRepeat', pattern: 'cards', span: null, max: 6 }]).rejected[0]?.reason ?? '', /6 cards do not fit side by side/);
});

test('a row whose items are shown where the slide does not draw them takes no more and no wider', () => {
  const slots = [1, 2, 3].map((n) => ({ name: `card${n}Text`, kind: 'text' as const, shapeId: String(9 + n), ...(n === 3 ? { box: { x: 0.3, y: 0.4, w: 0.5, h: 0.3 } } : {}) }));
  assert.deepEqual(movedRowSlots(slots, rowOf(3)), ['card3Text']);
  assert.deepEqual(movedRowSlots(slots.slice(0, 2), rowOf(2)), []);

  // Three text boxes on the slide, the third with a box of its own, as a text-only variant made by 0.2.0 left it.
  const boxes = fixtureWithTextCards([500000, 1900000, 3300000]);
  const moved: TemplateManifest = {
    ...plain,
    patterns: [
      {
        ...(plain.patterns[0] as TemplateManifest['patterns'][number]),
        slots: (plain.patterns[0] as TemplateManifest['patterns'][number]).slots.map((slot) => (slot.name === 'card3Text' ? { ...slot, box: { x: 0.36, y: 0.39, w: 0.5, h: 0.29 } } : slot)),
        repeat: { item: 'card', slots: ['card{n}Text'], shapes: [['11'], ['21'], ['31']], min: 2, max: 3 },
      },
    ],
  };
  const template = loadTemplate(boxes, moved);
  assert.deepEqual(checkManifest(template).filter((issue) => issue.severity === 'error'), []);
  for (const change of [{ op: 'setRepeat', pattern: 'cards', widen: 2 }, { op: 'setRepeat', pattern: 'cards', max: 4 }] as const) {
    assert.match(changeManifest(template, [change]).rejected[0]?.reason ?? '', /"card3Text" is shown in another place than the slide draws it, so the row cannot take more cards or wider ones: .* Take the box away \(setSlot with box: null\) first/);
  }
  assert.deepEqual(suggestChanges(template, { kinds: ['row-max', 'row-widen'] }), []);
  // The way out the refusal names works: the box is taken away, and the row is a row again.
  const freed = changeManifest(template, [{ op: 'setSlot', pattern: 'cards', slot: 'card3Text', box: null }, { op: 'setRepeat', pattern: 'cards', widen: 2 }]);
  assert.deepEqual(freed.rejected, []);
  assert.equal(freed.manifest.patterns[0]?.repeat?.widen, 2);
  // Without a row at all, slots with a box are not made one.
  const { repeat: _repeat, ...loose } = moved.patterns[0] as TemplateManifest['patterns'][number];
  assert.match(changeManifest(loadTemplate(boxes, { ...moved, patterns: [loose] }), [{ op: 'setRepeat', pattern: 'cards' }]).rejected[0]?.reason ?? '', /"card3Text" is shown in another place than the slide draws it, and a row is laid out from where the slide draws its items/);
});

test('a manifest with a span on a row inside a group can still be changed, and the span taken away', () => {
  // 0.2.0 accepted a span and a number of items for such a row, though the renderer uses neither the span nor room outside the group.
  const grouped = cardsAt(2000000, 4000000, 6000000, true);
  const stale: TemplateManifest = { ...plain, patterns: [{ ...(plain.patterns[0] as TemplateManifest['patterns'][number]), repeat: { item: 'card', slots: ['card{n}Text'], shapes: [['10'], ['20'], ['30']], min: 1, max: 5, span: { left: 0.05, right: 0.95 } } }] };
  const template = loadTemplate(grouped, stale);
  assert.deepEqual(checkManifest(template).filter((issue) => issue.severity === 'error'), []);
  for (const change of [{ op: 'setRepeat', pattern: 'cards', widen: 2 }, { op: 'setRepeat', pattern: 'cards', min: 2 }, { op: 'setRepeat', pattern: 'cards', span: null }] as const) {
    assert.deepEqual(changeManifest(template, [change]).rejected, [], JSON.stringify(change));
  }
  assert.equal(changeManifest(template, [{ op: 'setRepeat', pattern: 'cards', span: null }]).manifest.patterns[0]?.repeat?.span, undefined);
  // Asked about the number, it counts the row where the group draws it, and never says fewer than are drawn.
  const asked = changeManifest(template, [{ op: 'setRepeat', pattern: 'cards', max: 5 }]);
  assert.match(asked.rejected[0]?.reason ?? '', /5 cards do not fit side by side: .* The most is 4/);
  assert.deepEqual(changeManifest(template, [{ op: 'setRepeat', pattern: 'cards', max: 4 }]).rejected, []);
  // And nothing is suggested on the strength of the span.
  assert.deepEqual(suggestChanges(template, { kinds: ['row-max'] }), []);
});

test('a sample deck is one the manifest allows, when a row takes fewer than it draws', () => {
  const fewer = changeManifest(loadTemplate(pptx, cardsManifest), [{ op: 'setRepeat', pattern: 'cards', max: 2 }]);
  assert.deepEqual(fewer.rejected, []);
  const template = loadTemplate(pptx, fewer.manifest);
  const { deck, assets } = sampleDeck(template);
  assert.equal(deck.slides.find((slide) => slide.pattern === 'cards')?.items, 2, 'the nearest the row takes to the three it draws');
  assert.ok(renderDeck(template, deck, { assets }).pptx.length > 1000);
});

test('a slot that points at a shape the slide does not have makes no row, and nothing throws', () => {
  const { repeat: _repeat, ...loose } = plain.patterns[0] as TemplateManifest['patterns'][number];
  const stale: TemplateManifest = { ...plain, patterns: [{ ...loose, slots: loose.slots.map((slot) => (slot.name === 'card3Text' ? { ...slot, shapeId: '999' } : slot)) }] };
  const template = loadTemplate(pptx, stale);
  assert.ok(checkManifest(template).some((issue) => /points at shape #999/.test(issue.message)));
  assert.deepEqual(suggestChanges(template), []);
  assert.match(changeManifest(template, [{ op: 'setRepeat', pattern: 'cards' }]).rejected[0]?.reason ?? '', /has no row of like items/);
});

test('a slot whose shape lies in an item is part of the row when every item has one, and holds the row otherwise', () => {
  type Slots = TemplateManifest['patterns'][number]['slots'];
  const made = (slots: Slots, slide = 4) => {
    const manifest: TemplateManifest = { ...fixtureManifest, patterns: [...fixtureManifest.patterns, { id: 'cards', sourceSlide: slide, description: 'Cards.', slots: [{ name: 'title', kind: 'text', shapeId: '2', size: 26, maxChars: 30 }, ...slots] }] };
    const result = changeManifest(loadTemplate(pptx, manifest), [{ op: 'setRepeat', pattern: 'cards' }]);
    return { manifest: result.manifest, repeat: result.manifest.patterns.find((pattern) => pattern.id === 'cards')?.repeat, reason: result.rejected[0]?.reason ?? '' };
  };
  const text = (n: number, name = `card${n}Text`): Slots[number] => ({ name, kind: 'text', shapeId: `${n}1`, size: 9, maxChars: 20 });
  const icon = (n: number, name: string, empty?: 'keep' | 'drop'): Slots[number] => ({ name, kind: 'image', imageKind: 'icon', shapeId: `${n}2`, ...(empty ? { empty } : {}) });
  const rendersAt = (manifest: TemplateManifest, items: 'min' | 'max') => {
    const template = loadTemplate(pptx, manifest);
    const { deck, assets } = sampleDeck(template, { items });
    return renderDeck(template, deck, { assets, layout: 'adaptive' }).pptx.length > 1000;
  };

  // A pictogram kept as designed in each card, which no deck fills: it goes where its card goes.
  // Left out of the row, a slide with two cards could not be rendered: the third pictogram is gone with its card.
  const kept = made([1, 2, 3].flatMap((n) => [text(n), icon(n, `card${n}Icon`, 'keep')]));
  assert.deepEqual([kept.repeat?.slots, kept.repeat?.min], [['card{n}Text', 'card{n}Icon'], 2]);
  assert.ok(rendersAt(kept.manifest, 'min') && rendersAt(kept.manifest, 'max'));

  // Named by its kind and not by its card: step1 and icon1.
  const steps = made([1, 2, 3].flatMap((n) => [text(n, `step${n}`), icon(n, `icon${n}`)]));
  assert.deepEqual([steps.repeat?.item, steps.repeat?.slots, steps.repeat?.min], ['step', ['step{n}', 'icon{n}'], 2]);
  assert.ok(rendersAt(steps.manifest, 'min') && rendersAt(steps.manifest, 'max'));

  // A slot only one item has: the row takes no fewer items than that one.
  const third = made([...[1, 2, 3, 4].map((n) => text(n)), icon(3, 'badge')], 5);
  assert.deepEqual([third.repeat?.slots, third.repeat?.min, third.repeat?.max], [['card{n}Text'], 3, 4]);
  assert.ok(rendersAt(third.manifest, 'min') && rendersAt(third.manifest, 'max'));

  // In the last item there is no row: a further item would be a copy of that one, and a slide with fewer has lost it.
  const last = made([...[1, 2, 3].map((n) => text(n)), icon(3, 'badge')]);
  assert.equal(last.repeat, undefined);
  assert.match(last.reason, /has no row of like items/);
  assert.deepEqual(suggestChanges(loadTemplate(pptx, last.manifest), { kinds: ['row'] }), []);
});

test('a row whose slide draws one item to stand out moves its items in a deck that reads the other way, and keeps it first for its readers', () => {
  // The fixture's row of cards, first at the right, with the first card's text on green.
  const green = fixtureWithSlide(4, (xml) =>
    xml.replace(/(<p:cNvPr id="11" name="Card text"\/>[\s\S]*?<\/a:xfrm>)(<\/p:spPr>)/, '$1<a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:solidFill><a:srgbClr val="00A252"/></a:solidFill>$2'),
  );
  const shapes = cardsManifest.patterns.find((pattern) => pattern.id === 'cards')?.repeat?.shapes ?? [];
  assert.equal(standsOut(PptxPackage.load(green), 'ppt/slides/slide4.xml', shapes), 1);
  assert.equal(standsOut(PptxPackage.load(buildFixturePptx()), 'ppt/slides/slide4.xml', shapes), undefined, 'cards alike');
  assert.equal(standsOut(PptxPackage.load(green), 'ppt/slides/slide4.xml', shapes.slice(0, 2)), undefined, 'two unlike cards do not say which stands out');

  // Suggested, and taken by changeManifest.
  const template = loadTemplate(green, cardsManifest);
  const found = suggestChanges(template, { kinds: ['row-emphasis'] });
  assert.deepEqual(found.map((suggestion) => suggestion.changes), [[{ op: 'setRepeat', pattern: 'cards', emphasis: 1 }]]);
  assert.match(found[0]?.message ?? '', /^Card 1 of "cards" is drawn unlike the other cards\./);
  assert.deepEqual(suggestChanges(loadTemplate(buildFixturePptx(), cardsManifest), { kinds: ['row-emphasis'] }), []);
  const changed = changeManifest(template, [{ op: 'setRepeat', pattern: 'cards', emphasis: 1 }]);
  assert.deepEqual(changed.rejected, []);
  assert.match(changed.applied[0]?.message ?? '', /Card 1 stands out: in a deck that reads the other way the cards change places, each with its look, so it is still card 1 for its readers\./);
  assert.deepEqual(suggestChanges(loadTemplate(green, changed.manifest), { kinds: ['row-emphasis'] }), [], 'asked again, nothing more to do');
  const marked = loadTemplate(green, changed.manifest);

  /** Where each card of the first slide stands, left edge in EMU, and the text it holds. */
  const cards = (file: Uint8Array) => {
    const root = rootElement(parseXml(PptxPackage.load(file).readText('ppt/slides/slide1.xml')));
    return ['10', '20', '30']
      .map((id) => {
        const group = descendants(root, NS.p, 'grpSp').find((element) => path(element, [NS.p, 'nvGrpSpPr'], [NS.p, 'cNvPr'])?.getAttribute('id') === id);
        if (!group) return undefined;
        const x = Number(path(group, [NS.p, 'grpSpPr'], [NS.a, 'xfrm'], [NS.a, 'off'])?.getAttribute('x'));
        const text = descendants(group, NS.a, 't').map((element) => element.textContent).join('');
        return { card: Number(id) / 10, x, text };
      })
      .filter((entry) => entry !== undefined)
      .sort((a, b) => a.x - b.x);
  };
  const deckOf = (direction: 'ltr' | 'rtl', items = 3): DeckSpec => ({
    version: 1,
    direction,
    slides: [{ pattern: 'cards', items, slots: { title: 'T', ...Object.fromEntries(['a', 'b', 'c'].slice(0, items).map((text, at) => [`card${at + 1}Text`, text])) } }],
  });
  // Without it, an English deck's first text goes to the card on the left and the green card on the right holds the last.
  assert.deepEqual(cards(renderDeck(template, deckOf('ltr')).pptx).map((entry) => `${entry.card}:${entry.text}`), ['3:a', '2:b', '1:c']);
  // With it, the green card moves to the left with its text.
  const moved = cards(renderDeck(marked, deckOf('ltr')).pptx);
  assert.deepEqual(moved.map((entry) => `${entry.card}:${entry.text}`), ['1:a', '2:b', '3:c']);
  assert.deepEqual(
    moved.map((entry) => entry.x),
    cards(renderDeck(template, deckOf('ltr')).pptx).map((entry) => entry.x),
    'the same places, taken by other cards',
  );
  assert.deepEqual(cards(renderDeck(marked, deckOf('ltr', 2)).pptx).map((entry) => `${entry.card}:${entry.text}`), ['1:a', '2:b']);
  // A deck that reads as the template does is the same file with it or without it.
  assert.deepEqual(renderDeck(marked, deckOf('rtl')).pptx, renderDeck(template, deckOf('rtl')).pptx);

  // A number that is no item is refused, null takes it away, and a manifest written by hand is checked.
  assert.match(changeManifest(template, [{ op: 'setRepeat', pattern: 'cards', emphasis: 4 }]).rejected[0]?.reason ?? '', /the card that stands out is one of the 3 drawn, by its number from 1/);
  assert.equal('emphasis' in (changeManifest(marked, [{ op: 'setRepeat', pattern: 'cards', emphasis: null }]).manifest.patterns.find((pattern) => pattern.id === 'cards')?.repeat ?? {}), false);
  const wrong: TemplateManifest = { ...cardsManifest, patterns: cardsManifest.patterns.map((pattern) => (pattern.repeat ? { ...pattern, repeat: { ...pattern.repeat, emphasis: 9 } } : pattern)) };
  assert.ok(checkManifest(loadTemplate(green, wrong)).some((issue) => issue.message === 'repeat.emphasis is the number of the item that stands out, from 1 to 3'));
});

test('fewer lines than drawn stand further apart with spread, up to its factor and as far as the area they are drawn in', () => {
  // Six topics one under another, 0.5 cm high and 0.7 cm apart from top to top, from 1,000,000 to 5,000,000 EMU.
  const box = (id: number, y: number) =>
    `<p:sp><p:nvSpPr><p:cNvPr id="${id}" name="Topic"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x="1000000" y="${y}"/><a:ext cx="4000000" cy="500000"/></a:xfrm></p:spPr>` +
    `<p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>Topic</a:t></a:r></a:p></p:txBody></p:sp>`;
  const slide =
    '<p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><p:cSld><p:spTree>' +
    [0, 1, 2, 3, 4, 5].map((n) => box(10 + n, 1000000 + n * 700000)).join('') +
    '</p:spTree></p:cSld></p:sld>';
  const tops = (count: number, spread?: number) => {
    const root = rootElement(parseXml(slide));
    const repeat: RepeatDef = { item: 'topic', slots: ['topic{n}'], shapes: [0, 1, 2, 3, 4, 5].map((n) => [String(10 + n)]), columns: 1, min: 1, max: 6, ...(spread ? { spread } : {}) };
    applyRepeat(indexShapes(root), repeat, count, 12192000);
    return descendants(root, NS.a, 'off').map((off) => Number(off.getAttribute('y')));
  };
  // Three of six bunch in the middle at the drawn step.
  assert.deepEqual(tops(3), [2050000, 2750000, 3450000]);
  // With 1.6 the step grows by 60%; with 3 it grows only as far as the area, from its top to its foot.
  assert.deepEqual(tops(3, 1.6), [1630000, 2750000, 3870000]);
  assert.deepEqual(tops(3, 3), [1000000, 2750000, 4500000]);
  // As many lines as drawn, or one alone, stand as without it.
  assert.deepEqual(tops(6, 1.6), tops(6));
  assert.deepEqual(tops(1, 1.6), tops(1));

  // A studio sets it on a list or grid of several lines, and not on a single row.
  const grid = loadTemplate(buildFixturePptx(), gridManifest);
  const set = changeManifest(grid, [{ op: 'setRepeat', pattern: 'grid', spread: 1.6 }]);
  assert.deepEqual(set.rejected, []);
  assert.match(set.applied[0]?.message ?? '', /With fewer lines of cards than drawn, the space between them grows up to 1\.6 times, as far as the area they are drawn in\./);
  assert.equal(set.manifest.patterns.find((pattern) => pattern.id === 'grid')?.repeat?.spread, 1.6);
  assert.match(changeManifest(grid, [{ op: 'setRepeat', pattern: 'grid', spread: 5 }]).rejected[0]?.reason ?? '', /spread is a factor between 1 and 3/);
  assert.match(changeManifest(loadTemplate(buildFixturePptx(), cardsManifest), [{ op: 'setRepeat', pattern: 'cards', spread: 1.6 }]).rejected[0]?.reason ?? '', /the cards stand on one line, so there are no lines to spread/);
  assert.equal('spread' in (changeManifest(loadTemplate(buildFixturePptx(), set.manifest), [{ op: 'setRepeat', pattern: 'grid', spread: null }]).manifest.patterns.find((pattern) => pattern.id === 'grid')?.repeat ?? {}), false);
});
