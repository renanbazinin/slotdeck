// Free slides: slides the software lays out itself from a list of blocks, any number of them.
import assert from 'node:assert/strict';
import { readFileSync } from 'node:fs';
import { dirname, join } from 'node:path';
import { test } from 'node:test';
import { fileURLToPath } from 'node:url';
import { strFromU8, strToU8, unzipSync, zipSync } from 'fflate';
import * as fontkit from 'fontkit';
import { unstatedChartValues } from '../src/deck/numbers.ts';
import { describeBlocks, freeContent, freeSlide, freeSlotDefs, freeSlots } from '../src/free/blocks.ts';
import type { FreeBlockContent } from '../src/free/blocks.ts';
import { contrast } from '../src/color/space.ts';
import { checkDeck, freeTextFloor, overflowMessage } from '../src/deck/validate.ts';
import { FREE_FRAME_ROOM, freeBaseRefusal, freeKit, freeKitOrReason, freeOffered, widened } from '../src/free/kit.ts';
import type { FreeTableSlot, FreeTextSlot } from '../src/free/plan.ts';
import { arrangeBlocks, settle } from '../src/free/layout.ts';
import { FREE_KIND_NOTES, MOST_BLOCKS, blockRuns, blocksOfRuns, freeOffer, freePattern } from '../src/free/offer.ts';
import type { FreeLayout, LayoutFrame } from '../src/free/layout.ts';
import { freeFor } from '../src/free/resolve.ts';
import type { FreeBlockKind } from '../src/deck/spec.ts';
import { changeManifest, checkDesign, checkManifest, DeckValidationError, editManifest, extractAssets, fontkitMetrics, loadTemplate, renderDeck, slideAlternatives, textMeasurer, validateDeck } from '../src/index.ts';
import type { ChartValue, DeckSpec, JsonRequest, LlmProvider, SlideSpec, TemplateManifest } from '../src/index.ts';
import { placeholderPhoto } from '../examples/placeholder.ts';
import { buildFixturePptx, cardsManifest, fixtureManifest, fixtureWithSlide, gridManifest, tableManifest } from './fixture.ts';
import { openFonts } from './fonts.ts';
import { withoutCodes } from './issues.ts';

const root = join(dirname(fileURLToPath(import.meta.url)), '..');
const manifest = JSON.parse(readFileSync(join(root, 'templates', 'demo.manifest.json'), 'utf8')) as TemplateManifest;
const demoPptx = new Uint8Array(readFileSync(join(root, 'templates', 'demo.pptx')));
const template = loadTemplate(demoPptx, manifest);
const icon = (id: string) => new Uint8Array(readFileSync(join(root, 'examples', 'icons', `${id}.png`)));
const photo = placeholderPhoto(400, 300, [14, 124, 123], [244, 162, 89]);

const chart: ChartValue = { type: 'column', categories: ['Spring', 'Summer'], series: [{ name: '2025', values: [7, 19] }] };
const card = (heading: string, text: string): FreeBlockContent => ({ card: { heading, text } });
const threeCards = freeSlide({ title: 'Three ways in', blocks: [card('Longer hours', 'Bikes unlock from five.'), card('Safer docks', 'Every station gets a light.'), card('Group passes', 'One pass for five riders.')] });
const deckOf = (slides: SlideSpec[], more: Partial<DeckSpec> = {}): DeckSpec => ({ version: 1, fields: { organizationDepartment: 'Riverbend' }, ...more, slides });

const files = (pptx: Uint8Array) => unzipSync(pptx);
/** The demo template with some of its parts changed, for a case the demo does not draw. */
function demoWith(changes: Record<string, (xml: string) => string>): Uint8Array {
  const all = unzipSync(demoPptx);
  for (const [name, change] of Object.entries(changes)) all[name] = strToU8(change(strFromU8(all[name] as Uint8Array)));
  return zipSync(all);
}
/** The demo manifest with only some of its patterns. */
const only = (...ids: string[]): TemplateManifest => ({ ...manifest, patterns: manifest.patterns.filter((pattern) => ids.includes(pattern.id)) });
const slideXml = (pptx: Uint8Array, number: number) => strFromU8(files(pptx)[`ppt/slides/slide${number}.xml`] as Uint8Array);
const textsOf = (xml: string) => [...xml.matchAll(/<a:t>([^<]*)<\/a:t>/g)].map((match) => match[1] as string);
const namesOf = (xml: string) => [...xml.matchAll(/<p:cNvPr id="(\d+)" name="([^"]*)"/g)].map((match) => ({ id: match[1] as string, name: match[2] as string }));
/** Whether a chart is drawn without its value axis. */
const axesGone = (xml: string) => /<c:valAx>(?:(?!<\/c:valAx>).)*?<c:delete val="1"\/>/.test(xml);
/** The chart parts a slide shows, in the order of its relationships. */
const chartsOf = (pptx: Uint8Array, number: number): string[] => {
  const all = files(pptx);
  const rels = strFromU8(all[`ppt/slides/_rels/slide${number}.xml.rels`] as Uint8Array);
  return [...rels.matchAll(/Target="\.\.\/charts\/(chart\d+\.xml)"/g)].map((match) => strFromU8(all[`ppt/charts/${match[1]}`] as Uint8Array));
};

// ---- Writing a free slide -------------------------------------------------------------------

test('freeSlide writes the kinds of the blocks and their content under numbered names', () => {
  assert.deepEqual(threeCards, {
    pattern: 'free',
    blocks: ['card', 'card', 'card'],
    slots: {
      title: 'Three ways in',
      b1Heading: 'Longer hours',
      b1Text: 'Bikes unlock from five.',
      b2Heading: 'Safer docks',
      b2Text: 'Every station gets a light.',
      b3Heading: 'Group passes',
      b3Text: 'One pass for five riders.',
    },
  });
  const mixed = freeSlide({ blocks: [{ chart, caption: 'North' }, { photo: 'pier' }, { note: 'Both grew.' }], columns: 2, lead: 'top', base: 'content', colors: { accent: '1565C0' } });
  assert.deepEqual(mixed.blocks, ['chart', 'photo', 'note']);
  assert.deepEqual(mixed.slots, { b1Caption: 'North', b1Chart: chart, b2Photo: { asset: 'pier' }, b3Text: 'Both grew.' });
  assert.deepEqual([mixed.arrange, mixed.base, mixed.colors], [{ columns: 2, lead: 'top' }, 'content', { accent: '1565C0' }]);
  assert.throws(() => freeSlide({ blocks: [{ poem: 'x' } as unknown as FreeBlockContent] }), /block 1 is none of the kinds a free slide has/);
});

test('freeContent gives back what freeSlide took, so blocks can be added and removed without renumbering', () => {
  const content = {
    title: 'A season in numbers',
    blocks: [
      { text: ['First paragraph.', 'Second paragraph.'], heading: 'Why' },
      { list: ['One', 'Two'] },
      { card: { heading: 'Hours', text: 'From five.', icon: { asset: 'clock' } } },
      { figure: { value: '87%', label: 'would recommend it' } },
      { step: { heading: 'Unlock', text: 'Scan the code.' } },
      { quote: 'I sold my second car.', source: 'A rider' },
      { note: 'The bottom line.' },
      { chart, caption: 'Rides' },
      { table: { header: ['Station', 'Docks'], rows: [['Tram Square', '24']] } },
      { photo: { asset: 'pier', alt: 'The pier' }, caption: 'The pier at dawn' },
      { graphic: { asset: 'map' } },
    ] satisfies FreeBlockContent[],
    columns: 3,
    lead: 'left' as const,
  };
  const slide = freeSlide(content);
  assert.deepEqual(freeContent(slide), content);
  // Dropping the first block renumbers the rest.
  const without = freeSlide({ ...freeContent(slide), blocks: freeContent(slide).blocks.slice(1) });
  assert.deepEqual(without.slots.b1Items, ['One', 'Two']);
  assert.equal(without.blocks?.length, 10);
});

test('freeContent and freeSlide keep a dense slide dense', () => {
  const text = ['A first paragraph of the text a person asked to have whole.', 'And a second one.'];
  const dense: SlideSpec = { pattern: 'free', blocks: ['text'], arrange: { dense: true }, slots: { title: 'All of it', b1Text: text } };
  assert.equal(freeContent(dense).dense, true);
  assert.deepEqual(freeSlide(freeContent(dense)), dense);
  const both: SlideSpec = { ...dense, arrange: { columns: 2, dense: true } };
  assert.deepEqual(freeSlide(freeContent(both)), both);
  // Code that changes the blocks the way the docs say keeps the flag.
  const headed = freeSlide({ ...freeContent(dense), blocks: [{ text, heading: 'Why' }] });
  assert.deepEqual(headed.arrange, { dense: true });
  // A slide without the flag comes back as before: no `dense` in its content, and no `arrange` added.
  const plain = freeSlide({ title: 'All of it', blocks: [{ text }] });
  assert.equal('dense' in freeContent(plain), false);
  assert.equal(JSON.stringify(freeSlide(freeContent(plain))), JSON.stringify(plain));
  assert.equal('arrange' in freeSlide(freeContent(plain)), false);
  const columns = freeSlide({ title: 'All of it', blocks: [{ text }], columns: 1 });
  assert.equal(JSON.stringify(freeSlide(freeContent(columns))), JSON.stringify({ pattern: 'free', blocks: ['text'], slots: { title: 'All of it', b1Text: text }, arrange: { columns: 1 } }));
});

test('the slots of a free slide are known from its blocks alone', () => {
  assert.deepEqual(freeSlots(['card', 'chart']).map((entry) => `${entry.name}${entry.part.required ? '' : '?'}`), ['b1Icon?', 'b1Heading', 'b1Text', 'b2Caption?', 'b2Chart']);
  assert.deepEqual(freeSlotDefs(['figure', 'photo']).map((slot) => `${slot.name}:${slot.kind}`), ['title:text', 'b1Value:text', 'b1Label:text', 'b2Photo:image', 'b2Caption:text']);
  assert.equal(describeBlocks(['chart', 'chart', 'note']), '2 charts and a note');
  assert.equal(describeBlocks(['text']), 'a text');
  assert.equal(describeBlocks(['icon' as FreeBlockKind, 'card', 'card', 'table']), 'an icon, 2 cards and a table');
});

// ---- Where the blocks go ---------------------------------------------------------------------

const frame: LayoutFrame = { x: 0.0375, w: 0.9225, top: 0.3, bottom: 0.88, denseTop: 0.26, denseBottom: 0.95, gap: 0.0225, aspect: 16 / 9 };
const boxes = (layout: FreeLayout) => layout.blocks.map((block) => [block.box.x, block.box.y, block.box.w, block.box.h]);
/** Every block lies in the room it was given, and no two of them touch. */
function assertTidy(layout: FreeLayout, count: number, label: string): void {
  assert.equal(layout.blocks.length, count, label);
  const top = layout.dense ? frame.denseTop : frame.top;
  const bottom = layout.dense ? frame.denseBottom : frame.bottom;
  layout.blocks.forEach((block, index) => {
    const { x, y, w, h } = block.box;
    assert.equal(block.index, index + 1, label);
    assert.ok(w > 0 && h > 0, `${label}: block ${block.index} has a size`);
    assert.ok(x >= frame.x - 0.0006 && x + w <= frame.x + frame.w + 0.0006, `${label}: block ${block.index} is inside the sides`);
    assert.ok(y >= top - 0.0006 && y + h <= bottom + 0.0006, `${label}: block ${block.index} is inside top and bottom (${y} to ${y + h})`);
  });
  for (const a of layout.blocks) {
    for (const b of layout.blocks) {
      if (a.index >= b.index) continue;
      const apart = a.box.x + a.box.w <= b.box.x + 0.0006 || b.box.x + b.box.w <= a.box.x + 0.0006 || a.box.y + a.box.h <= b.box.y + 0.0006 || b.box.y + b.box.h <= a.box.y + 0.0006;
      assert.ok(apart, `${label}: blocks ${a.index} and ${b.index} overlap`);
    }
  }
}

test('one block has the whole room; three cards stand side by side', () => {
  const one = arrangeBlocks(['text'], { frame, rtl: false });
  assert.deepEqual(boxes(one), [[0.0375, 0.3, 0.9225, 0.58]]);
  const three = arrangeBlocks(['card', 'card', 'card'], { frame, rtl: false });
  assert.deepEqual([three.columns, three.tiles, three.dense], [3, false, false]);
  assert.deepEqual(three.blocks.map((block) => [block.box.x, block.box.w]), [[0.0375, 0.2925], [0.3525, 0.2925], [0.6675, 0.2925]]);
  assertTidy(three, 3, 'three cards');
});

test('a deck read from the right has its first block at the right', () => {
  const ltr = arrangeBlocks(['card', 'card', 'card'], { frame, rtl: false });
  const rtl = arrangeBlocks(['card', 'card', 'card'], { frame, rtl: true });
  assert.deepEqual(rtl.blocks.map((block) => block.box.x), ltr.blocks.map((block) => block.box.x).reverse());
  // A side is a side of the slide, whatever the reading direction.
  const left = arrangeBlocks(['chart', 'text'], { frame, rtl: true, arrange: { lead: 'left' } });
  assert.equal(left.blocks[0]?.lead, true);
  assert.equal(left.blocks[0]?.box.x, frame.x);
  assertTidy(rtl, 3, 'three cards from the right');
});

test('two hundred charts are set out in a grid, in the larger room', () => {
  const layout = arrangeBlocks(Array<FreeBlockKind>(200).fill('chart'), { frame, rtl: false });
  assert.deepEqual([layout.columns, layout.dense, layout.tiles], [20, true, false]);
  assert.equal(new Set(layout.blocks.map((block) => block.box.y)).size, 10);
  assertTidy(layout, 200, '200 charts');
});

test('one chart among a few cards leads; a note stands under what it sums up', () => {
  const beside = arrangeBlocks(['chart', 'card', 'card', 'card'], { frame, rtl: false });
  assert.equal(beside.blocks[0]?.lead, true);
  assert.ok((beside.blocks[0]?.box.w ?? 0) > 0.5, 'the chart takes more than half the width');
  assert.deepEqual(new Set(beside.blocks.slice(1).map((block) => block.box.x)).size, 1, 'the cards are stacked beside it');
  assertTidy(beside, 4, 'chart beside cards');

  const noted = arrangeBlocks(['chart', 'chart', 'note'], { frame, rtl: false });
  const note = noted.blocks[2];
  assert.deepEqual([note?.box.x, note?.box.w], [frame.x, frame.w]);
  assert.ok((note?.box.y ?? 0) > (noted.blocks[0]?.box.y ?? 1) + (noted.blocks[0]?.box.h ?? 1), 'the note is under the charts');
  assertTidy(noted, 3, 'two charts and a note');

  // A note that comes first is a band over the rest.
  const opened = arrangeBlocks(['note', 'card', 'card'], { frame, rtl: false });
  assert.equal(opened.blocks[0]?.box.y, frame.top);
  assertTidy(opened, 3, 'a note first');
});

test('notes share two fifths of the room with other blocks, take back what those leave, and never leave the room', () => {
  // A note of three lines is 0.08 of the slide's width high, and lower by as much as its text is set
  // smaller: six of them ask for more than the room has.
  const tall = (_index: number, _width: number, lower: number) => 0.08 * lower;
  const bottomOf = (layout: FreeLayout) => Math.max(...layout.blocks.map((block) => block.box.y + block.box.h));
  // Charts fill their share of the room, so the notes have two fifths, the gap after each counted in.
  const charts = arrangeBlocks(['chart', 'chart', 'note', 'note', 'note', 'note', 'note', 'note'], { frame, rtl: false, noteHeight: tall });
  assertTidy(charts, 8, 'two charts and six notes');
  assert.equal(charts.dense, true, 'a slide whose notes are made lower counts as crowded');
  const room = frame.denseBottom - frame.denseTop;
  const notes = charts.blocks.filter((block) => block.kind === 'note');
  const first = notes[0] as (typeof notes)[number];
  const second = notes[1] as (typeof notes)[number];
  const taken = bottomOf(charts) - first.box.y + (second.box.y - first.box.y - first.box.h);
  assert.ok(taken > 0.39 * room && taken < 0.401 * room, `the notes take ${(taken / room).toFixed(3)} of the room`);
  assert.equal(new Set(notes.map((block) => block.box.h)).size, 1, 'every note is lowered by the same factor');
  assert.ok(first.box.h < 0.5 * 0.08 * frame.aspect);
  assert.ok((charts.blocks[0]?.box.h ?? 0) > 0.55 * room, 'the charts keep three fifths');
  // Three low cards do not keep three fifths of a slide empty: the notes have what the cards leave.
  const cards = arrangeBlocks(['card', 'card', 'card', 'note', 'note', 'note', 'note', 'note', 'note'], { frame, rtl: false, noteHeight: tall, needed: () => 0.1 });
  assertTidy(cards, 9, 'three cards and six notes');
  assert.ok((cards.blocks[3]?.box.h ?? 0) > 1.2 * first.box.h, 'notes beside low cards are higher than beside charts');
  // Notes that fit once the cards have what they need are as high as they asked to be.
  const two = arrangeBlocks(['card', 'card', 'note', 'note'], { frame, rtl: false, noteHeight: () => 0.06, needed: () => 0.1 });
  assert.deepEqual([two.dense, two.blocks[2]?.box.h, two.blocks[3]?.box.h], [false, 0.1067, 0.1067]);
  assertTidy(two, 4, 'two cards and two notes');
  // A slide of notes alone has the whole room, and a note that comes first is lowered with the rest.
  const alone = arrangeBlocks(Array<FreeBlockKind>(12).fill('note'), { frame, rtl: false, noteHeight: tall });
  assertTidy(alone, 12, 'twelve notes');
  assert.ok(bottomOf(alone) > frame.denseBottom - 0.002, 'they share the room to its bottom');
  assertTidy(arrangeBlocks(['note', 'chart', 'note', 'note', 'note', 'note', 'note'], { frame, rtl: false, noteHeight: tall }), 7, 'a note first, a chart and five notes');
  assertTidy(arrangeBlocks(['note', 'note', 'photo', 'note', 'note', 'note'], { frame, rtl: false, noteHeight: tall, arrange: { lead: 'top' } }), 6, 'notes around a photo');
  // Smaller text takes fewer lines, so a note is asked again at every size that is tried. These take three
  // lines at their own size and one from seven tenths of it down: four of them under a chart are set at
  // seven tenths, where lowering what they first asked for would have left them a third.
  const wrapping = (_index: number, _width: number, lower: number) => ((lower > 0.7 ? 3 : 1) * 0.015 + 0.02) * lower;
  const few = arrangeBlocks(['chart', 'note', 'note', 'note', 'note'], { frame, rtl: false, noteHeight: wrapping });
  assertTidy(few, 5, 'a chart and four notes that wrap');
  assert.ok((few.lower ?? 1) > 0.699 && (few.lower ?? 1) <= 0.7, `set at ${few.lower}`);
  assert.equal(few.blocks[1]?.box.h, Math.round(0.035 * 0.7 * frame.aspect * 10000) / 10000);
  assert.equal(charts.lower !== undefined && two.lower === undefined, true, 'the layout says when its notes are set smaller');
  // A caller that does not answer to the size has its notes made lower as they are.
  assertTidy(arrangeBlocks(['chart', 'note', 'note', 'note', 'note', 'note', 'note'], { frame, rtl: false, noteHeight: () => 0.08 }), 7, 'notes that stay as high as they were');
});

test('how many stand side by side, and which block leads, can be asked for', () => {
  const two = arrangeBlocks(Array<FreeBlockKind>(6).fill('chart'), { frame, rtl: false, arrange: { columns: 2 } });
  assert.equal(two.columns, 2);
  assert.equal(new Set(two.blocks.map((block) => block.box.y)).size, 3);
  assertTidy(two, 6, 'six charts in two columns');

  const top = arrangeBlocks(['chart', 'figure', 'figure', 'figure'], { frame, rtl: false, arrange: { lead: 'top' } });
  assert.equal(top.blocks[0]?.lead, true);
  assert.ok(top.blocks.slice(1).every((block) => block.box.y > (top.blocks[0]?.box.y ?? 1)), 'the figures are under the chart');
  assertTidy(top, 4, 'a chart over figures');

  const none = arrangeBlocks(['chart', 'card', 'card'], { frame, rtl: false, arrange: { lead: 'none' } });
  assert.ok(none.blocks.every((block) => !block.lead));
  // A number of columns that is no number is not followed.
  assert.deepEqual(boxes(arrangeBlocks(['card', 'card'], { frame, rtl: false, arrange: { columns: 0 } })), boxes(arrangeBlocks(['card', 'card'], { frame, rtl: false })));
});

test('more blocks asked for on one line than the room holds side by side still have a size each', () => {
  // The gaps alone would take more than the width: 400 gaps of 0.3% of the slide in a room of 92%.
  const wide = arrangeBlocks(Array<FreeBlockKind>(400).fill('text'), { frame, rtl: false, arrange: { columns: 400 } });
  assert.equal(wide.columns, 400);
  assertTidy(wide, 400, '400 texts on one line');
  const high = arrangeBlocks(Array<FreeBlockKind>(300).fill('chart'), { frame, rtl: false, arrange: { columns: 1 } });
  assertTidy(high, 300, '300 charts one under the other');
  const beside = arrangeBlocks(['chart', ...Array<FreeBlockKind>(200).fill('text')], { frame, rtl: false, arrange: { columns: 200, lead: 'left' } });
  assertTidy(beside, 201, '200 texts on one line beside a chart');
  // Such a slide is a file PowerPoint opens: every position and size is a number.
  for (const kind of ['text', 'chart'] as const) {
    const blocks = Array.from({ length: 320 }, (): FreeBlockContent => (kind === 'text' ? { text: 'A line.' } : { chart }));
    const deck = deckOf([freeSlide({ title: 'Every station', blocks, columns: 320 })]);
    assert.deepEqual(validateDeck(template, deck).filter((issue) => issue.severity === 'error'), [], kind);
    const xml = slideXml(renderDeck(template, deck).pptx, 1);
    assert.doesNotMatch(xml, /NaN|="-\d/, kind);
    assert.equal((xml.match(/<a:ext cx="[1-9]\d*" cy="[1-9]\d*"\/>/g) ?? []).length >= 320, true, kind);
  }
});

test('any mix of blocks, from one to a hundred and twenty, is set out without overlap', () => {
  const kinds: FreeBlockKind[] = ['text', 'list', 'card', 'figure', 'step', 'quote', 'note', 'chart', 'table', 'photo', 'graphic'];
  // A fixed sequence, so that a failure can be found again.
  let seed = 7;
  const next = () => (seed = (seed * 1103515245 + 12345) % 2147483648) / 2147483648;
  for (let round = 0; round < 160; round++) {
    const count = round < 60 ? 1 + (round % 12) : 1 + Math.floor(next() * 120);
    const uniform = round % 3 === 0;
    const kind = kinds[Math.floor(next() * kinds.length)] as FreeBlockKind;
    const blocks = Array.from({ length: count }, () => (uniform ? kind : (kinds[Math.floor(next() * kinds.length)] as FreeBlockKind)));
    // Notes are bands of their own: a slide of nothing but notes is not what the grid is for.
    const some = blocks.filter((entry) => entry !== 'note').length > 0 ? blocks.map((entry, index) => (entry === 'note' && index > 2 && index < count - 1 ? 'card' : entry)) : ['text' as FreeBlockKind];
    for (const rtl of [false, true]) {
      const layout = arrangeBlocks(some, { frame, rtl });
      assert.equal(layout.blocks.length, some.length);
      for (const a of layout.blocks) {
        assert.ok(a.box.w > 0 && a.box.h > 0, `round ${round}: block ${a.index} of ${some.join(',')} has a size`);
        assert.ok(a.box.x >= frame.x - 0.0006 && a.box.x + a.box.w <= frame.x + frame.w + 0.0006, `round ${round}: block ${a.index} of ${some.join(',')} is inside the sides`);
        for (const b of layout.blocks) {
          if (a.index >= b.index) continue;
          const apart = a.box.x + a.box.w <= b.box.x + 0.0006 || b.box.x + b.box.w <= a.box.x + 0.0006 || a.box.y + a.box.h <= b.box.y + 0.0006 || b.box.y + b.box.h <= a.box.y + 0.0006;
          assert.ok(apart, `round ${round}: blocks ${a.index} and ${b.index} of ${some.join(',')} overlap`);
        }
      }
    }
  }
});

// ---- What a template gives a free slide ------------------------------------------------------

test('the demo template: the frame a free slide stands on and the styles its texts take', () => {
  const kit = freeKit(template);
  assert.ok(kit);
  assert.deepEqual([kit.base, kit.baseSlide, kit.firstId, kit.background, kit.direction], ['list-with-photo', 4, 8, 'FFFFFF', 'ltr']);
  // The logo, the organisation line, the page number and the title stay; the list and the photo go.
  assert.deepEqual(kit.stays.slice().sort(), ['2', '3', '4', '5']);
  assert.deepEqual(kit.frameSlots.map((slot) => slot.name), ['organizationDepartment']);
  assert.deepEqual([kit.title?.shapeId, kit.title?.size, kit.title?.place], ['5', 34, false]);
  assert.deepEqual(kit.band, { top: 0.3, bottom: 0.88 });
  assert.equal(kit.gap, 0.0225);
  const styles = Object.fromEntries(Object.entries(kit.styles).map(([role, style]) => [role, `${style.slide}#${style.shapeId} ${style.size}${style.bold ? ' bold' : ''}${style.italic ? ' italic' : ''}`]));
  assert.deepEqual(styles, {
    body: '2#7 18',
    heading: '5#10 19 bold',
    figure: '6#8 40 bold',
    caption: '8#7 13 italic',
    list: '7#8 18',
    quote: '11#6 32 italic',
    note: '9#8 17 bold',
  });
  assert.deepEqual([kit.panel?.slide, kit.panel?.shapeId, kit.panel?.fill], [5, '7', 'E8F1F0']);
  assert.deepEqual([kit.table?.slide, kit.picture?.slide, kit.figureOnPanel], [8, 3, true]);
  assert.equal(freeKit(template), kit, 'worked out once for a template');
});

test('a manifest can name the frame, and can stop free slides being offered without breaking stored ones', () => {
  const named = loadTemplate(demoPptx, { ...manifest, free: { base: 'text-only' } });
  assert.equal(freeKit(named)?.base, 'text-only');
  // A slide says which frame it was planned on, and keeps it.
  assert.equal(freeKit(template, 'text-only')?.base, 'text-only');
  const kept = deckOf([{ ...threeCards, base: 'text-only' }]);
  assert.deepEqual(validateDeck(template, kept), []);
  // A frame the template no longer has is told, and the slide stands on the usual one.
  assert.match(validateDeck(template, deckOf([{ ...threeCards, base: 'gone' }])).map((issue) => issue.message).join(' | '), /slide 1 was planned on the frame of the layout "gone".*it stands on "list-with-photo"/);
  assert.equal(freeOffered(manifest), true);

  const off = loadTemplate(demoPptx, { ...manifest, free: false });
  assert.equal(freeOffered(off.manifest), false);
  const deck = deckOf([threeCards]);
  assert.deepEqual(validateDeck(off, deck), []);
  assert.deepEqual(files(renderDeck(off, deck).pptx)['ppt/slides/slide1.xml'], files(renderDeck(template, deck).pptx)['ppt/slides/slide1.xml']);
});

test('small templates carry a free slide too', () => {
  for (const [name, small] of Object.entries({ fixtureManifest, cardsManifest, gridManifest, tableManifest })) {
    const loaded = loadTemplate(buildFixturePptx(), small);
    const kit = freeKitOrReason(loaded);
    assert.ok(typeof kit !== 'string', `${name}: ${String(kit)}`);
    const deck: DeckSpec = { version: 1, slides: [freeSlide({ title: 'Three things', blocks: [{ text: 'A short text.' }, card('A heading', 'And its text.'), { chart }, { table: { header: ['A', 'B'], rows: [['1', '2']] } }] })] };
    assert.deepEqual(validateDeck(loaded, deck).filter((issue) => issue.severity === 'error'), [], name);
    const xml = slideXml(renderDeck(loaded, deck).pptx, 1);
    assert.ok(textsOf(xml).includes('And its text.'), name);
    // A template with no table to copy has one drawn for it.
    assert.match(xml, /<a:tbl>/, name);
  }
});

/** Slide 2 of the demo as PowerPoint makes a slide: its title is the layout's title placeholder, with no position of its own. */
const placeholderTitle = (): Uint8Array =>
  demoWith({
    'ppt/slides/_rels/slide2.xml.rels': (xml) => xml.replace('slideLayout7.xml', 'slideLayout2.xml'),
    'ppt/slides/slide2.xml': (xml) => xml.replace(/(<p:cNvPr id="5" name="Title"\/>)<p:cNvSpPr txBox="1"\/><p:nvPr\/><\/p:nvSpPr><p:spPr>.*?<\/p:spPr>/, '$1<p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr><p:nvPr><p:ph type="title"/></p:nvPr></p:nvSpPr><p:spPr/>'),
  });

test('a title that takes its place from the slide layout is read from there, and is never moved', () => {
  const loaded = loadTemplate(placeholderTitle(), manifest);
  const kit = freeKit(loaded, 'text-and-photo');
  assert.ok(kit);
  assert.equal(kit.base, 'text-and-photo');
  // The layout states no place either: the master's title says where it stands.
  assert.deepEqual(kit.title?.box, { x: 0.0375, y: 0.04, w: 0.675, h: 0.1667 });
  assert.equal(kit.title?.place, false);
  assert.ok(kit.band.top > 0.2067 && kit.band.top < 0.32, JSON.stringify(kit.band));
  const deck = deckOf([{ ...threeCards, base: 'text-and-photo' }]);
  assert.deepEqual(validateDeck(loaded, deck), []);
  const xml = slideXml(renderDeck(loaded, deck).pptx, 1);
  // The title is filled where it is: it has no position on the slide to change.
  assert.match(xml, /<p:ph type="title"\/><\/p:nvPr><\/p:nvSpPr><p:spPr\/>/);
  assert.ok(textsOf(xml).includes('Three ways in'));
  // Without a title the blocks start where it would stand, and never over the header that stays.
  const untitled = freeFor(loaded, {}, { ...threeCards, slots: { ...threeCards.slots, title: '' }, base: 'text-and-photo' });
  assert.ok(typeof untitled !== 'string');
  assert.ok(Math.min(...untitled.plan.layout.blocks.map((block) => block.box.y)) >= 0.107, JSON.stringify(untitled.plan.layout.blocks[0]));

  // A placeholder whose layout and master state no place either gives no frame: another layout is stood on,
  // and where there is none the deck's check says why.
  const nowhere = fixtureWithSlide(1, (xml) => xml.replace(/(<p:cNvPr id="2" name="Title"\/>)<p:cNvSpPr txBox="1"\/><p:nvPr\/><\/p:nvSpPr><p:spPr>.*?<\/p:spPr>/, '$1<p:cNvSpPr/><p:nvPr><p:ph type="title"/></p:nvPr></p:nvSpPr><p:spPr/>'));
  assert.equal(freeKit(loadTemplate(nowhere, tableManifest))?.base, 'table');
  const reason = 'the title of the layout "content" takes its place from the slide layout, and neither that nor the master states one';
  assert.equal(freeKitOrReason(loadTemplate(nowhere, fixtureManifest)), reason);
  assert.deepEqual(validateDeck(loadTemplate(nowhere, fixtureManifest), { version: 1, slides: [threeCards] }), [{ severity: 'error', slide: 1, message: `a free slide cannot be made from this template: ${reason}` }]);
});

test('a layout whose title sits low is no frame: the usual choice passes it over, and one that is named gives way with the reason', () => {
  assert.equal(FREE_FRAME_ROOM, 0.45);
  // A cover, a section divider and a closing slide leave blocks a strip under their title.
  assert.equal(freeKit(loadTemplate(demoPptx, only('cover', 'list-with-photo', 'closing')))?.base, 'list-with-photo');
  assert.equal(freeKit(loadTemplate(demoPptx, only('cover', 'text-and-photo', 'list-with-photo', 'section', 'closing')))?.base, 'list-with-photo');
  // Each of the three has a subtitle under its title, which is the title's line and not content: its blocks would start under that.
  for (const [base, share] of [['cover', 22], ['section', 37], ['closing', 28]] as const) {
    const why = `the layout "${base}" has its title low on the slide: blocks would get ${share}% of the slide's height, and a frame leaves them 45% or more`;
    assert.equal(freeKit(template, base)?.base, 'list-with-photo', base);
    // The slide stands on the usual frame, and the deck is told why, as a note on the deck.
    assert.deepEqual(withoutCodes(validateDeck(template, deckOf([{ ...threeCards, base }]))), [{ severity: 'warning', message: `slide 1 was planned on the frame of the layout "${base}", which cannot carry a free slide (${why}); it stands on "list-with-photo"` }]);
    // A studio that names it is refused with the reason, and a manifest written by hand is told when it is checked.
    assert.equal(changeManifest(template, [{ op: 'setManifest', free: { base } }]).rejected[0]?.reason, `a free slide cannot stand on "${base}": ${why}; name a layout with a title that leaves room under it and has no row of items, such as "list-with-photo"`);
    const named = { pkg: template.pkg, manifest: { ...manifest, free: { base } } };
    assert.equal(freeKit(named)?.base, 'list-with-photo', base);
    assert.deepEqual(checkManifest(named), [{ severity: 'warning', message: `free.base names the pattern "${base}", which cannot be the frame of free slides: ${why}; they stand on "list-with-photo"` }]);
  }
  // A layout with a row of items, and one that is gone, are told as before.
  assert.match(validateDeck(template, deckOf([{ ...threeCards, base: 'two-columns' }]))[0]?.message ?? '', /"two-columns", which cannot carry a free slide \(the layout "two-columns" has a row of columns, which the software lays out itself\); it stands on "list-with-photo"$/);
  assert.match(validateDeck(template, deckOf([{ ...threeCards, base: 'gone' }]))[0]?.message ?? '', /"gone", which the template no longer offers; it stands on "list-with-photo"$/);
  // A template that has no layout with room enough is stood on the one that leaves most: its decks still render.
  const titles = loadTemplate(demoPptx, only('cover', 'section', 'closing'));
  const kit = freeKit(titles);
  assert.deepEqual([kit?.base, kit?.band.top, kit?.dense.bottom], ['section', 0.5822, 0.95]);
  const deck = deckOf([threeCards]);
  assert.deepEqual(validateDeck(titles, deck).filter((issue) => issue.severity === 'error'), []);
  assert.ok(textsOf(slideXml(renderDeck(titles, deck).pptx, 1)).includes('One pass for five riders.'));
  // No layout of that template leaves the room of a frame, so the reason does not speak of it: a layout
  // that is named and passed over is compared with the one the slides stand on.
  const low = 'the layout "closing" has its title low on the slide: blocks would get 28% of the slide\'s height, and "section" leaves them 37%';
  const planned = (base: string) => validateDeck(titles, deckOf([{ ...threeCards, base }])).filter((issue) => /was planned/.test(issue.message));
  assert.deepEqual(withoutCodes(planned('closing')), [{ severity: 'warning', message: `slide 1 was planned on the frame of the layout "closing", which cannot carry a free slide (${low}); it stands on "section"` }]);
  assert.equal(changeManifest(titles, [{ op: 'setManifest', free: { base: 'closing' } }]).rejected[0]?.reason, `a free slide cannot stand on "closing": ${low}; name a layout with a title that leaves room under it and has no row of items, such as "section"`);
  const closing = { pkg: titles.pkg, manifest: { ...titles.manifest, free: { base: 'closing' } } };
  assert.deepEqual(checkManifest(closing).filter((issue) => /free\.base/.test(issue.message)), [{ severity: 'warning', message: `free.base names the pattern "closing", which cannot be the frame of free slides: ${low}; they stand on "section"` }]);
  // The layout that leaves most is the frame there: naming it is no mistake.
  assert.deepEqual(planned('section'), []);
  assert.deepEqual(changeManifest(titles, [{ op: 'setManifest', free: { base: 'section' } }]).rejected, []);
  assert.deepEqual(checkManifest({ pkg: titles.pkg, manifest: { ...titles.manifest, free: { base: 'section' } } }).filter((issue) => /free\.base/.test(issue.message)), []);
});

test('on a dark frame a chart and a table the code draws are set in colours that read on it', () => {
  // Only the divider and the closing slide are left to stand on: a navy frame. The chart layout gives the
  // template's look of a chart without being a frame, and no table is there to copy.
  const native = manifest.patterns.find((pattern) => pattern.id === 'chart-native') as TemplateManifest['patterns'][number];
  const dark = loadTemplate(demoPptx, { ...manifest, patterns: [...only('section', 'closing').patterns, { ...native, id: 'chart-look', slots: native.slots.filter((slot) => slot.kind === 'chart') }] });
  const kit = freeKit(dark);
  assert.deepEqual([kit?.base, kit?.background, kit?.ink, kit?.chart.text?.color, kit?.chart.lineColor, kit?.table], ['section', '1B2A41', '000000', '5C6B73', 'D9E2E1', undefined]);
  const two: ChartValue = { type: 'column', categories: ['Spring', 'Summer'], series: [{ name: '2024', values: [5, 9] }, { name: '2025', values: [7, 19] }] };
  const slide = freeSlide({ title: 'Stations', blocks: [{ chart: two }, { table: { header: ['Station', 'Docks'], rows: [['Tram Square', '24'], ['River Gate', '18'], ['North Bridge', '12']] } }] });
  const { pptx } = renderDeck(dark, deckOf([slide]));
  const drawn = chartsOf(pptx, 1)[0] as string;
  // The template's chart text, 5C6B73, has a contrast of 2.6 with the navy: the slide's own text colour is taken.
  const textColor = /<a:defRPr[^>]*><a:solidFill><a:srgbClr val="([0-9A-F]{6})"/.exec(drawn)?.[1] ?? '';
  assert.equal(textColor, kit?.styles.body.color);
  assert.ok(contrast(textColor, '1B2A41') >= 3, `chart text ${textColor}`);
  for (const [, line] of drawn.matchAll(/<a:ln w="9525" cap="flat"><a:solidFill><a:srgbClr val="([0-9A-F]{6})"/g)) assert.ok(contrast(line as string, '1B2A41') >= 1.1, `chart line ${line}`);
  // The table: no near-white band behind light text, and every body cell states a colour that reads on the navy.
  const xml = slideXml(pptx, 1);
  const rows = [...xml.matchAll(/<a:tr\b.*?<\/a:tr>/g)].map((match) => match[0]);
  assert.equal(rows.length, 4);
  assert.doesNotMatch(xml, /F2F2F2/);
  for (const row of rows.slice(1)) {
    const colors = [...row.matchAll(/<a:rPr[^>]*><a:solidFill><a:srgbClr val="([0-9A-F]{6})"/g)].map((match) => match[1] as string);
    assert.equal(colors.length, 2, row);
    for (const color of colors) assert.ok(contrast(color, '1B2A41') >= 3, `table text ${color}`);
    assert.match(row, /<a:noFill\/><\/a:tcPr>/);
  }
  // On the light frame the template's own layouts have, a chart keeps the template's look.
  const light = chartsOf(renderDeck(template, deckOf([freeSlide({ title: 'Stations', blocks: [{ chart: two }] })])).pptx, 1)[0] as string;
  assert.match(light, /<a:defRPr sz="1200"><a:solidFill><a:srgbClr val="5C6B73"\/>/);
  assert.match(light, /<c:majorGridlines><c:spPr><a:ln w="9525" cap="flat"><a:solidFill><a:srgbClr val="D9E2E1"\/>/);
});

test('a template whose only running text is a bulleted list gives plain text all the same', () => {
  const listed = loadTemplate(demoPptx, only('list-with-photo'));
  const slide = freeSlide({
    title: 'Why it matters',
    blocks: [card('Longer hours', 'Bikes unlock from five.'), { figure: { value: '87%', label: 'would recommend it' } }, { quote: 'I sold my second car.', source: 'A rider' }, { list: ['Open earlier', 'Add forty bikes'] }, { note: 'The bottom line.' }],
  });
  const xml = slideXml(renderDeck(listed, deckOf([slide])).pptx, 1);
  const paragraphs = new Map([...xml.matchAll(/<p:cNvPr id="\d+" name="(Block \d+ [a-z]+)"\/>.*?<p:txBody>(.*?)<\/p:txBody>/g)].map((match) => [match[1] as string, match[2] as string]));
  // The list keeps the template's bullets; every other text is a copy of the same shape without them.
  assert.equal((paragraphs.get('Block 4 list')?.match(/<a:buChar\b/g) ?? []).length, 2);
  for (const name of ['Block 1 heading', 'Block 1 text', 'Block 2 value', 'Block 2 label', 'Block 3 quote', 'Block 3 source', 'Block 5 note']) {
    const body = paragraphs.get(name) ?? '';
    assert.match(body, /<a:pPr[^>]*\bmarL="0"[^>]*\bindent="0"[^>]*>.*<a:buNone\/>/, name);
    assert.doesNotMatch(body, /<a:buChar\b/, name);
  }
});

test('a template with no layout to stand on says so, in the deck\'s own check', () => {
  const bare = loadTemplate(buildFixturePptx(), { ...fixtureManifest, patterns: fixtureManifest.patterns.filter((pattern) => !pattern.slots.some((slot) => /^title$/i.test(slot.name))).slice(0, 1) });
  const reason = freeKitOrReason(bare);
  if (typeof reason !== 'string') return; // every pattern of this fixture has a title: nothing to refuse
  const issues = validateDeck(bare, { version: 1, slides: [freeSlide({ blocks: [{ text: 'x' }] })] });
  assert.match(issues[0]?.message ?? '', /a free slide cannot be made from this template/);
});

// ---- Rendering ----------------------------------------------------------------------------------

test('three cards: the frame stays, the base layout\'s own content goes, the blocks are drawn', () => {
  const deck = deckOf([threeCards]);
  assert.deepEqual(validateDeck(template, deck), []);
  const { pptx, warnings } = renderDeck(template, deck);
  assert.deepEqual(warnings, []);
  const xml = slideXml(pptx, 1);
  assert.deepEqual(textsOf(xml), ['Riverbend', '1', 'Three ways in', 'Longer hours', 'Bikes unlock from five.', 'Safer docks', 'Every station gets a light.', 'Group passes', 'One pass for five riders.']);
  assert.deepEqual(
    namesOf(xml).filter((shape) => shape.name !== '').map((shape) => `${shape.id}:${shape.name}`),
    ['2:Logo', '3:Organization', '4:Page number', '5:Title', '8:Block 1 panel', '9:Block 1 heading', '10:Block 1 text', '11:Block 2 panel', '12:Block 2 heading', '13:Block 2 text', '14:Block 3 panel', '15:Block 3 heading', '16:Block 3 text'],
  );
  // The base layout's photo is gone, and the relationship to it: only the logo is a picture.
  assert.equal((xml.match(/<p:pic\b/g) ?? []).length, 1);
  assert.equal((strFromU8(files(pptx)['ppt/slides/_rels/slide1.xml.rels'] as Uint8Array).match(/relationships\/image"/g) ?? []).length, 1);
  // A card's text is never set as large as its heading.
  const plan = freeFor(template, deck, threeCards);
  assert.ok(typeof plan !== 'string');
  const sizes = Object.fromEntries(plan.plan.shapes.filter((shape) => shape.kind === 'text').map((shape) => [shape.name, shape.size]));
  assert.ok((sizes['Block 1 text'] ?? 0) < (sizes['Block 1 heading'] ?? 0), JSON.stringify(sizes));
});

test('the same deck gives the same file, and a free slide does not depend on the slides around it', () => {
  const deck = deckOf([threeCards, freeSlide({ title: 'One chart', blocks: [{ chart }] })]);
  const first = renderDeck(template, deck).pptx;
  assert.deepEqual(renderDeck(template, deck).pptx, first);
  const among = renderDeck(template, deckOf([{ pattern: 'cover', slots: { title: 'A year of rides', subtitle: 'What the season taught us' } }, threeCards]));
  const strip = (xml: string) => xml.replace(/<a:t>\d+<\/a:t>/g, '');
  assert.equal(strip(slideXml(among.pptx, 2)), strip(slideXml(first, 1)));
});

test('a slide without a title has no empty title box', () => {
  const xml = slideXml(renderDeck(template, deckOf([freeSlide({ blocks: [{ text: 'One sentence, alone.' }] })])).pptx, 1);
  assert.deepEqual(namesOf(xml).map((shape) => shape.name).filter((name) => name === 'Title'), []);
  assert.ok(textsOf(xml).includes('One sentence, alone.'));
});

test('two hundred charts on one slide: every one is a chart of its own, drawn bare', () => {
  const deck = deckOf([freeSlide({ title: 'Every station', blocks: Array.from({ length: 200 }, () => ({ chart })) })]);
  const issues = validateDeck(template, deck);
  assert.deepEqual(issues.filter((issue) => issue.severity === 'error'), []);
  const { pptx } = renderDeck(template, deck);
  const all = files(pptx);
  assert.equal(Object.keys(all).filter((name) => /^ppt\/charts\/chart\d+\.xml$/.test(name)).length, 200);
  assert.equal((slideXml(pptx, 1).match(/<p:graphicFrame\b/g) ?? []).length, 200);
  const charts = chartsOf(pptx, 1);
  assert.equal(charts.length, 200);
  // No axis, no legend, and the plot fills its frame.
  for (const xml of [charts[0] as string, charts[199] as string]) {
    assert.ok(axesGone(xml));
    assert.match(xml, /<c:manualLayout>/);
    assert.doesNotMatch(xml, /<c:legend>/);
  }
  const ids = namesOf(slideXml(pptx, 1)).map((shape) => shape.id);
  assert.equal(new Set(ids).size, ids.length, 'every shape has its own id');
  // The deck is told what became of them, as a note on the deck: no rewriting of text changes it.
  assert.deepEqual(withoutCodes(issues), [{ severity: 'warning', message: 'slide 1 holds 200 blocks, as asked; 200 charts are drawn small, without axes, legend or labels' }]);
});

test('a chart is drawn in as much detail as its box reads at', () => {
  const two: ChartValue = { type: 'column', categories: ['Spring', 'Summer'], series: [{ name: '2024', values: [5, 9] }, { name: '2025', values: [7, 19] }] };
  const deck = deckOf([freeSlide({ title: 'One', blocks: [{ chart: two }] }), freeSlide({ title: 'Twelve', blocks: Array.from({ length: 12 }, () => ({ chart: two })) })]);
  const { pptx } = renderDeck(template, deck);
  const full = chartsOf(pptx, 1)[0] as string;
  assert.ok(!axesGone(full));
  assert.match(full, /<c:legend>/);
  assert.match(full, /<c:majorGridlines>/);
  assert.match(full, /<c:showVal val="1"\/>/);
  // A small chart keeps the names of its categories and has no gridlines. Its numbers stand at its bars
  // where they have room, in place of a value axis; and twelve like charts say once what their colours
  // stand for, in one legend on the slide, so none of them draws its own.
  const compact = chartsOf(pptx, 2)[0] as string;
  assert.match(compact, /<c:catAx>(?:(?!<\/c:catAx>).)*?<c:delete val="0"\/>/);
  assert.ok(axesGone(compact));
  assert.match(compact, /<c:showVal val="1"\/>/);
  assert.doesNotMatch(compact, /<c:majorGridlines>/);
  assert.doesNotMatch(compact, /<c:legend>/);
  assert.deepEqual(namesOf(slideXml(pptx, 2)).map((shape) => shape.name).filter((name) => /^Legend/.test(name)), ['Legend key 1', 'Legend 1', 'Legend key 2', 'Legend 2']);
  const resolved = freeFor(template, deck, deck.slides[1] as SlideSpec);
  assert.ok(typeof resolved !== 'string');
  assert.deepEqual([...new Set(resolved.plan.pattern.slots.filter((slot) => slot.kind === 'chart').map((slot) => (slot.kind === 'chart' ? slot.detail : undefined)))], ['compact']);
});

test('a crowded slide is made as asked, and the deck is told how small its text is', () => {
  const crowded = freeSlide({ title: 'Sixty cards', blocks: Array.from({ length: 60 }, (_, index) => card(`Card ${index + 1}`, 'A short line of text.')) });
  const issues = validateDeck(template, deckOf([threeCards, crowded]));
  assert.deepEqual(issues.filter((issue) => issue.severity === 'error'), []);
  const told = issues.filter((issue) => /holds 60 blocks, as asked; its smallest text is set at [\d.]+ pt/.test(issue.message));
  assert.equal(told.length, 1);
  // It names the slide in words, not as a slide to rewrite: no rewriting changes how many blocks were asked for.
  assert.equal(told[0]?.slide, undefined);
  assert.match(told[0]?.message ?? '', /^slide 2 /);
});

test('text, a list, a quotation, figures, steps, a note and a table are each drawn in the template\'s styles', () => {
  const slides = [
    freeSlide({ title: 'Why it matters', blocks: [{ text: ['Riders now choose the bike for one trip in three.', 'The docks next to the tram stops empty first.'], heading: 'Mornings' }, { list: ['Open earlier', 'Fix flats in a day', 'Add forty bikes'], heading: 'Next' }] }),
    freeSlide({ blocks: [{ quote: 'I sold my second car in June.', source: 'A rider' }] }),
    freeSlide({ title: 'A season in numbers', blocks: [{ figure: { value: '48k', label: 'rides' } }, { figure: { value: '87%', label: 'would recommend it' } }, { note: 'Dependable docks made it a habit.' }] }),
    freeSlide({ title: 'How a ride starts', blocks: [{ step: { heading: 'Find', text: 'Open the map.' } }, { step: { heading: 'Unlock', text: 'Scan the code.' } }, { step: { heading: 'Ride', text: 'Return to any dock.' } }] }),
    freeSlide({ title: 'Stations', blocks: [{ table: { header: ['Station', 'Docks', 'Rides a day'], rows: [['Tram Square', '24', '410'], ['River Gate', '18', '365']] }, caption: 'The two busiest' }] }),
  ];
  const deck = deckOf(slides);
  assert.deepEqual(validateDeck(template, deck), []);
  const { pptx, warnings } = renderDeck(template, deck);
  assert.deepEqual(warnings, []);
  const first = slideXml(pptx, 1);
  assert.deepEqual(textsOf(first).slice(3), ['Mornings', 'Riders now choose the bike for one trip in three.', 'The docks next to the tram stops empty first.', 'Next', 'Open earlier', 'Fix flats in a day', 'Add forty bikes']);
  // The list has bullets, three of them; the running text has none.
  assert.equal((first.match(/<a:buChar\b/g) ?? []).length, 3);
  assert.deepEqual(textsOf(slideXml(pptx, 2)).slice(2), ['I sold my second car in June.', 'A rider']);
  // Steps are numbered by the code.
  assert.deepEqual(textsOf(slideXml(pptx, 4)).slice(3), ['1', 'Find', 'Open the map.', '2', 'Unlock', 'Scan the code.', '3', 'Ride', 'Return to any dock.']);
  const table = slideXml(pptx, 5);
  assert.equal((table.match(/<a:tr\b/g) ?? []).length, 3);
  assert.ok(textsOf(table).includes('The two busiest'));
});

/** The plan of a slide on the demo template. */
function planOf(slide: SlideSpec, deck: Partial<DeckSpec> = {}, on = template) {
  const resolved = freeFor(on, deck, slide);
  assert.ok(typeof resolved !== 'string', String(resolved));
  const shape = (name: string) => {
    const found = resolved.plan.shapes.find((entry) => entry.name === name);
    assert.ok(found, name);
    return found;
  };
  const slot = (name: string) => {
    const found = resolved.plan.pattern.slots.find((entry) => entry.name === name);
    assert.ok(found, name);
    return found;
  };
  return { ...resolved, shape, slot };
}
/** Points of the demo slide: 960 by 540. */
const W = 12191695 / 12700;
const H = 540;

test('a table is as high as the lines of its cells need, and its text is set smaller until it fits its box', () => {
  const phrase = 'Tram Square station, platform B, by the river gate';
  const tableOf = (text: string) => ({ header: ['Station', 'Platform', 'Docks', 'Rides', 'Notes'], rows: Array.from({ length: 6 }, (_, row) => [text, text, '24', '410', row % 2 ? text : 'Open']) });
  const slide = freeSlide({ title: 'Stations', blocks: [{ table: tableOf(phrase) }, { note: 'Tram Square leads.' }] });
  const plan = planOf(slide);
  const slot = plan.slot('b1Table') as FreeTableSlot;
  // A phrase of 50 characters takes two lines even in the widest its column gets: the text comes down from the template's 15 pt.
  const size = slot.bodyText?.size ?? 0;
  assert.ok(size < 15 && size >= 10, `text at ${size} pt`);
  // The heading row has one line and the others two: each is as high as its lines and the space kept around them.
  const heights = slot.rowHeights ?? [];
  assert.equal(heights.length, 7);
  assert.ok((heights[0] as number) >= 1.5625 * size - 0.01 && (heights[0] as number) < (heights[1] as number));
  for (const height of heights.slice(1)) assert.ok(height >= 2 * 1.25 * size + 7.2 - 0.01, `a row of two lines at ${size} pt is ${height} pt high`);
  // The table's box is the sum of its rows, and ends above the note under it.
  const box = plan.shape('Block 1 table').box;
  assert.ok(Math.abs(box.h * H - heights.reduce((sum, height) => sum + height, 0)) < 0.1);
  assert.ok(box.y + box.h <= (plan.plan.layout.blocks[1]?.box.y ?? 0) + 0.0001, 'the table ends above the note');
  assert.deepEqual(validateDeck(template, deckOf([slide])), []);
  // The file has the rows at those heights, so that PowerPoint has nothing to make higher.
  const xml = slideXml(renderDeck(template, deckOf([slide])).pptx, 1);
  const rows = [...xml.matchAll(/<a:tr h="(\d+)"/g)].map((match) => Number(match[1]));
  assert.equal(rows.length, 7);
  rows.forEach((row, index) => assert.ok(Math.abs(row / 12700 - (heights[index] as number)) < 0.02, `row ${index + 1}`));
  assert.equal(Number(/<p:xfrm><a:off x="\d+" y="\d+"\/><a:ext cx="\d+" cy="(\d+)"\/><\/p:xfrm>/.exec(xml)?.[1]), rows.reduce((sum, row) => sum + row, 0));
  // Cells of one line leave the table as it was: rows of one height, the text at the template's size.
  const plain = planOf(freeSlide({ title: 'Stations', blocks: [{ table: { header: ['Station', 'Docks'], rows: [['Tram Square', '24'], ['River Gate', '18']] } }] })).slot('b1Table') as FreeTableSlot;
  assert.deepEqual([plain.bodyText?.size, plain.rowHeights], [15, undefined]);
});

test('a table that its box cannot hold at the smallest size is told on its slot', () => {
  const told = (table: { header: string[]; rows: string[][] }, more: Partial<DeckSpec> = {}) =>
    validateDeck(template, deckOf([freeSlide({ title: 'Stations', blocks: [{ table }] })], more))
      .filter((issue) => issue.slot !== undefined)
      .map((issue) => `slide ${issue.slide} / ${issue.slot}: ${issue.message}`);
  const stations = (count: number) => ({ header: ['Station', 'Docks'], rows: Array.from({ length: count }, (_, row) => [`Station ${row + 1}`, String(10 + row)]) });
  // Thirty-two rows of one line fit at 6 pt, the smallest a table is set at; sixty would be drawn off the slide.
  assert.deepEqual(told(stations(32)), []);
  assert.deepEqual(told(stations(33)), ['slide 1 / b1Table: the table has 33 rows; the slot fits 32']);
  assert.deepEqual(told(stations(60)), ['slide 1 / b1Table: the table has 60 rows; the slot fits 32']);
  // Rows that are few enough, with cells that take more lines than the box has for them.
  // What a cell may hold is stated at 12 pt: four fifths of the 15 pt these thirteen rows alone would have.
  const long = { header: ['A', 'B', 'C', 'D', 'E', 'F'], rows: Array.from({ length: 12 }, () => Array.from({ length: 6 }, () => 'x'.repeat(200))) };
  assert.deepEqual(told(long), ['slide 1 / b1Table: a cell is 200 characters; cells fit about 20']);
  // A table that fits at that size says nothing, whatever lines its cells take and whatever the deck's
  // text size: what the slot holds is worked out without the deck's wish, and is not cut a second time.
  const phrase = 'Tram Square station, platform B';
  const wrapped = { header: ['Station', 'Platform', 'Docks'], rows: Array.from({ length: 5 }, () => [phrase, phrase, '24']) };
  for (const textScale of [1, 1.25, 1.5]) assert.deepEqual(told(wrapped, { textScale }), [], `at ${textScale}`);
  // Cells that take the text under four fifths are told on the slot of the table, with its slide: the table
  // is still drawn inside its box, in small text, and shorter cells are what makes that text larger.
  const station = 'Tram Square station, platform B, by the river gate and the market hall';
  // Every column holds the long cells: a single column of them would be given the width it asks for (see the test of column widths).
  const cellsOf = (length: number) => ({ header: ['Station', 'Docks', 'Rides', 'Open', 'Notes'], rows: Array.from({ length: 9 }, () => Array.from({ length: 5 }, () => `${station} ${station}`.slice(0, length))) });
  const over = (length: number) => freeSlide({ title: 'Stations', blocks: [{ table: cellsOf(length) }, { note: 'Tram Square leads.' }] });
  const sizeOf = (length: number) => (planOf(over(length)).slot('b1Table') as FreeTableSlot).bodyText?.size;
  // (Each cell keeps the 0.15 inch the template's table keeps free at its sides.)
  assert.deepEqual([8, 24, 31, 60].map(sizeOf), [15, 11.5, 9, 7.5]);
  assert.deepEqual(validateDeck(template, deckOf([over(24)])), []);
  for (const length of [31, 60]) {
    // At 7.5 pt the table's text is under what reads on a projected slide: the deck is told that too, in a note of its own.
    const small = length === 60 ? [{ severity: 'warning', message: 'slide 1 holds 2 blocks, as asked; its smallest text is set at 7.5 pt' }] : [];
    assert.deepEqual(withoutCodes(validateDeck(template, deckOf([over(length)]))), [...small, { severity: 'warning', slide: 1, slot: 'b1Table', message: `a cell is ${length} characters; cells fit about 24` }]);
    const plan = planOf(over(length));
    assert.ok(plan.shape('Block 1 table').box.y + plan.shape('Block 1 table').box.h <= (plan.plan.layout.blocks[1]?.box.y ?? 0) + 0.0001, 'the table still ends above the note');
  }
  // Where the text is small for another reason too (ninety characters take it to the smallest size), the
  // note on the deck that tells the size stands beside the warning on the slot, and no longer alone.
  assert.deepEqual(validateDeck(template, deckOf([over(90)])).map((issue) => `${issue.slot ?? 'deck'}: ${issue.message}`), ['deck: slide 1 holds 2 blocks, as asked; its smallest text is set at 6 pt', 'b1Table: a cell is 90 characters; cells fit about 24']);
  // Before there is content, the slot says what a table is written to in its fields. Its description says
  // what the fields do not, and does not repeat them.
  const planned = freePattern(template, {}, { blocks: ['table'] }, false);
  assert.ok(typeof planned !== 'string');
  const slot = planned.pattern.slots.find((entry) => entry.name === 'b1Table');
  assert.ok(slot?.kind === 'table');
  assert.deepEqual([slot.maxColumns, slot.maxRows, slot.maxCellChars, slot.description], [8, 14, 16, 'A table; a word or a short phrase in each cell. When the brief gives nothing for it, leave it empty']);
});

test('cards and steps too low for their text have it set smaller, inside the card', () => {
  const column = (kind: 'card' | 'step', count: number, text = 'Bikes unlock from five in the morning.') => freeSlide({ title: 'Ways in', blocks: Array.from({ length: count }, (_, index): FreeBlockContent => (kind === 'card' ? { card: { heading: `Point ${index + 1}`, text } } : { step: { heading: `Step ${index + 1}`, text } })), columns: 1 });
  const sizes: number[] = [];
  for (const kind of ['card', 'step'] as const) {
    for (const count of [5, 6, 7]) {
      const slide = column(kind, count);
      const plan = planOf(slide);
      for (let block = 1; block <= count; block++) {
        const panel = plan.shape(`Block ${block} panel`).box;
        const heading = plan.shape(`Block ${block} heading`);
        const text = plan.shape(`Block ${block} text`);
        const label = `${count} ${kind}s, block ${block}`;
        // The text box holds a line of its text, and ends above the card's edge.
        assert.ok(text.box.h * H >= 1.2 * (text.size ?? 0) - 0.06, `${label}: a box ${(text.box.h * H).toFixed(1)} pt high for ${text.size} pt text`);
        assert.ok(text.box.y + text.box.h <= panel.y + panel.h + 0.0002, `${label}: the text ends ${((text.box.y + text.box.h - panel.y - panel.h) * H).toFixed(1)} pt below its card`);
        assert.ok(heading.box.y + heading.box.h <= text.box.y + 0.0001, label);
        assert.ok((heading.size ?? 0) >= (text.size ?? 0), label);
      }
      // Nothing is reported: the text fits where it is set.
      assert.deepEqual(validateDeck(template, deckOf([slide])), [], `${count} ${kind}s`);
      if (kind === 'card') sizes.push(plan.shape('Block 1 text').size ?? 0);
    }
  }
  // Five cards in a column still have room for the template's own size; from six the size follows the height.
  assert.deepEqual(sizes, [18, 14, 12]);
  // Eight in a column would each be under half the least height a card sits well in: that wish is
  // not followed (see the test of a wish the room cannot hold).
  assert.deepEqual(planOf(column('card', 8)).plan.layout.relaxed, { asked: 1, columns: 2 });
  // Text that takes more than the card holds at that size is told on its slot.
  const full = column('card', 7, 'Bikes unlock from five in the morning, for early shifts and the first trams. '.repeat(4).trim());
  const issues = validateDeck(template, deckOf([full]));
  assert.equal(issues.length, 7);
  assert.match(`${issues[0]?.slot}: ${issues[0]?.message}`, /^b1Text: text is 307 characters; the slot fits about \d+$/);
  // A box lower than a line holds no line: a hundred and fifty texts one under the other are told, each on its slot.
  const stack = freeSlide({ title: 'Texts', blocks: Array.from({ length: 150 }, () => ({ text: 'A line of text.' })), columns: 1 });
  const stacked = validateDeck(template, deckOf([stack])).filter((issue) => issue.slot !== undefined);
  assert.equal(stacked.length, 150);
  assert.deepEqual([stacked[0]?.slot, stacked[0]?.message], ['b1Text', 'text is 15 characters; the slot fits about 1']);
});

/** The demo with the lines of its body text set further apart: by a share of the usual, or at a height in points. */
const spacedBody = (spacing: string): Uint8Array => demoWith({ 'ppt/slides/slide2.xml': (xml) => xml.replaceAll('<a:pPr algn="l"><a:spcAft>', `<a:pPr algn="l"><a:lnSpc>${spacing}</a:lnSpc><a:spcAft>`) });

test('a text is counted at the line spacing its template sets: lines further apart leave a card less text, and it stays inside', () => {
  const column = (count: number, text = 'Bikes unlock from five in the morning.') => freeSlide({ title: 'Ways in', blocks: Array.from({ length: count }, (_, index): FreeBlockContent => ({ card: { heading: `Point ${index + 1}`, text } })), columns: 1 });
  assert.equal(freeKit(template)?.styles.body.leading, undefined);
  // A line spacing of 130%: every line of body text is 1.56 times its size high. A height in points: 30 pt, whatever the size.
  for (const [spacing, leading, lineAt, sizes] of [
    ['<a:spcPct val="130000"/>', { percent: 1.3 }, (size: number) => 1.56 * size, [15, 12.5, 10.5]],
    ['<a:spcPts val="3000"/>', { points: 30 }, (size: number) => Math.max(30, 1.2 * size), [12, 7.5, 5]],
  ] as const) {
    const spaced = loadTemplate(spacedBody(spacing), manifest);
    assert.deepEqual(freeKit(spaced)?.styles.body.leading, leading);
    const got: number[] = [];
    // Up to seven cards one under the other: a wish for eight in a column is not followed.
    for (const count of [5, 6, 7]) {
      const slide = column(count);
      const plan = planOf(slide, {}, spaced);
      const size = plan.shape('Block 1 text').size ?? 0;
      got.push(size);
      const fits = plan.plan.layout.blocks.every((block) => {
        const panel = plan.shape(`Block ${block.index} panel`).box;
        const text = plan.shape(`Block ${block.index} text`);
        // A line of the text, as high as the template sets it, ends above the card's edge.
        return text.box.y * H + lineAt(size) <= (panel.y + panel.h) * H + 0.06;
      });
      const told = validateDeck(spaced, deckOf([slide])).filter((issue) => issue.slot !== undefined);
      // Either the line is inside its card, or the deck is told that the card holds none: never neither.
      assert.ok(fits ? told.length === 0 : told.length === count && told[0]?.message === 'text is 38 characters; the slot fits about 1', `${spacing}, ${count} cards at ${size} pt: ${fits ? 'inside' : 'outside'}, ${told.length} told`);
    }
    assert.deepEqual(got, sizes, spacing);
  }
  // The copy keeps the spacing of the shape it is made from: the file says what the plan counted with.
  const spaced = loadTemplate(spacedBody('<a:spcPct val="130000"/>'), manifest);
  const xml = slideXml(renderDeck(spaced, deckOf([column(5)])).pptx, 1);
  assert.ok(/name="Block 1 text"\/>.*?<a:pPr[^>]*><a:lnSpc><a:spcPct val="130000"\/><\/a:lnSpc>.*?sz="1500"/.test(xml), 'the text of a card keeps the line spacing of the body text');
  // Before there is content, the same room is said to hold fewer characters: its lines are counted as high as they are.
  const long = 'Bikes unlock from five in the morning, for early shifts and the first trams. '.repeat(5).trim();
  const budget = (on: typeof template) => {
    const planned = freePattern(on, {}, { blocks: ['text', 'chart'] }, false);
    assert.ok(typeof planned !== 'string');
    const slot = planned.pattern.slots.find((entry) => entry.name === 'b1Text');
    return slot?.kind === 'text' ? (slot.maxChars ?? 0) : 0;
  };
  assert.ok(budget(spaced) < 0.85 * budget(template) && budget(spaced) > 0.65 * budget(template), `${budget(spaced)} of ${budget(template)}`);
  assert.deepEqual(validateDeck(spaced, deckOf([freeSlide({ title: 'Three ways in', blocks: [card('Longer hours', long.slice(0, 150)), card('Safer docks', long.slice(0, 150)), card('Group passes', long.slice(0, 150))] })])), []);
});

test('on a template with wide line spacing the text of a card is fitted as on any other, and a card is as high as the lines its text really takes', () => {
  // The estimate of a line's length is a generous one, made for lines of the usual height: with a line spacing
  // of 130% a text really takes about what the estimate gives it at the usual height. So the length of a line and
  // its height are counted the same way: both by the estimate, and both as PowerPoint sets them. Counting the
  // height as it is and the length by the estimate set card text that fitted a fifth smaller, in a card a third empty.
  const spaced = loadTemplate(spacedBody('<a:spcPct val="130000"/>'), manifest);
  const long = 'Bikes unlock from five in the morning, for early shifts and the first trams. '.repeat(5).trim();
  const three = (text: string) => freeSlide({ title: 'Three ways in', blocks: [card('Longer hours', text), card('Safer docks', text), card('Group passes', text)] });
  const cardOf = (text: string, on: typeof template) => {
    const plan = planOf(three(text), {}, on);
    // A card that grew into room left over is counted as high as its content asked.
    return { high: ((plan.plan.layout.blocks[0]?.box.h ?? 0) - (plan.plan.layout.blocks[0]?.grown ?? 0)) * H, size: plan.shape('Block 1 text').size ?? 0, box: plan.shape('Block 1 text').box.h * H, budget: (plan.slot('b1Text') as FreeTextSlot).maxChars ?? 0 };
  };
  // 150 characters: eight lines by the estimate, 173 pt at the usual height. PowerPoint sets them in six lines of
  // 28 pt (168 pt). The card is as high on both templates, and the text keeps the template's size.
  const [wide, usual] = [cardOf(long.slice(0, 150), spaced), cardOf(long.slice(0, 150), template)];
  assert.deepEqual([wide.size, usual.size, Math.round(wide.high)], [18, 18, Math.round(usual.high)]);
  assert.ok(wide.box >= 6 * 1.56 * 18 - 4.5, `a text box ${wide.box.toFixed(1)} pt high`);
  assert.ok(wide.budget >= 150 && wide.budget <= usual.budget, `${wide.budget} of ${usual.budget}`);
  assert.deepEqual(validateDeck(spaced, deckOf([three(long.slice(0, 150))])), []);
  // A text of one line is where the estimate is short: one line of the usual height is lower than the line the
  // template sets. The card is as high as that line.
  const [oneWide, oneUsual] = [cardOf('Bikes from five.', spaced), cardOf('Bikes from five.', template)];
  assert.ok(Math.abs(oneWide.high - oneUsual.high - (1.56 - 1.2) * 18) < 0.6, `${oneWide.high.toFixed(1)} and ${oneUsual.high.toFixed(1)} pt`);
  // A box holds no more than its lines, as high as they are, hold when PowerPoint sets them: text longer than
  // that is set smaller or reported, however much the estimate would let in at the usual height.
  const beside = (text: string) => freeSlide({ title: 'Rides', blocks: [{ chart }, card('Longer hours', text), card('Safer docks', text), card('Group passes', text)] });
  const low = planOf(beside(long.slice(0, 60)), {}, spaced);
  const lowBox = low.shape('Block 2 text').box.h * H;
  const lowSize = low.shape('Block 2 text').size ?? 0;
  const lines = Math.floor((lowBox + 0.25 * lowSize) / (1.56 * lowSize));
  assert.ok(lines >= 1 && ((low.slot('b2Text') as FreeTextSlot).maxChars ?? 0) <= lines * Math.floor((low.shape('Block 2 text').box.w * W) / (0.42 * lowSize)), `${lines} lines of ${lowSize} pt hold ${(low.slot('b2Text') as FreeTextSlot).maxChars}`);
  // The demo itself has the usual spacing: nothing of this changes a slide on it.
  assert.equal(freeKit(template)?.styles.body.leading, undefined);
});

test('the space between the paragraphs of a free slide is the plan\'s: a copy has none of its source\'s before or after', () => {
  // The demo's list and its caption with 14 pt before every paragraph, as templates have it. PowerPoint draws
  // that space before a first paragraph too, so a caption would sit 14 pt lower than the plan put it.
  const apart = loadTemplate(
    demoWith({
      'ppt/slides/slide7.xml': (xml) => xml.replaceAll('<a:spcAft><a:spcPts val="1200"/></a:spcAft>', '<a:spcBef><a:spcPts val="1400"/></a:spcBef><a:spcAft><a:spcPts val="1200"/></a:spcAft>'),
      'ppt/slides/slide8.xml': (xml) => xml.replace('<a:pPr algn="l"/><a:r><a:rPr sz="1300" b="0" i="1">', '<a:pPr algn="l"><a:spcBef><a:spcPts val="1400"/></a:spcBef><a:spcAft><a:spcPts val="1400"/></a:spcAft></a:pPr><a:r><a:rPr sz="1300" b="0" i="1">'),
    }),
    manifest,
  );
  const slide = freeSlide({ title: 'Rides', blocks: [{ chart, caption: 'Rides by season' }, { list: ['Open earlier', 'Fix flats in a day', 'Add forty bikes'] }, card('Longer hours', 'Bikes unlock from five.')] });
  for (const on of [apart, template]) {
    const plan = planOf(slide, {}, on);
    const xml = slideXml(renderDeck(on, deckOf([slide])).pptx, 1);
    const paragraphs = new Map([...xml.matchAll(/<p:cNvPr id="\d+" name="(Block \d+ (?:caption|list|heading|text))"\/>.*?<p:txBody>(.*?)<\/p:txBody>/g)].map((match) => [match[1] as string, [...(match[2] as string).matchAll(/<a:pPr[^>]*>(.*?)<\/a:pPr>/g)].map((found) => found[1] as string)]));
    const gap = (plan.slot('b2Items') as FreeTextSlot).paragraphGap ?? 0;
    assert.ok(gap > 0);
    // Every paragraph states no space before it. After it there is what the slot says for a list, and none elsewhere.
    for (const [name, found] of paragraphs) {
      const after = name === 'Block 2 list' ? Math.round(gap * 100) : 0;
      assert.ok(found.length > 0, name);
      for (const props of found) assert.ok(props.startsWith(`<a:spcBef><a:spcPts val="0"/></a:spcBef><a:spcAft><a:spcPts val="${after}"/></a:spcAft>`), `${name}: ${props}`);
    }
    assert.equal(paragraphs.get('Block 2 list')?.length, 3);
    assert.ok(['Block 1 caption', 'Block 3 heading', 'Block 3 text'].every((name) => paragraphs.has(name)));
    assert.doesNotMatch(xml, /spcPts val="1400"/);
  }
  // The same slide is planned alike on both: the plan never counted the template's paragraph spacing.
  assert.deepEqual(planOf(slide, {}, apart).plan.layout, planOf(slide).plan.layout);
});

test('a step heading as long as its budget says is a heading of one line', () => {
  for (const count of [4, 6, 12]) {
    const planned = freePattern(template, {}, { blocks: Array<FreeBlockKind>(count).fill('step') }, false);
    assert.ok(typeof planned !== 'string');
    const budget = (name: string) => {
      const slot = planned.pattern.slots.find((entry) => entry.name === name);
      return slot?.kind === 'text' ? (slot.maxChars ?? 0) : 0;
    };
    const steps = (heading: number) => freeSlide({ title: 'How a ride starts', blocks: Array.from({ length: count }, () => ({ step: { heading: 'h'.repeat(heading), text: 'Scan the code on the dock.' } })) });
    const short = planOf(steps(budget('b1Heading') - 3));
    const full = planOf(steps(budget('b1Heading')));
    // The budget was stated for the width the heading has: a heading of that length takes one line and does not
    // take the text's room. Where the text is short and set larger than it was planned at, the heading over it is
    // larger too, and a heading as long as its budget is set up to 15% smaller to stay on one line: the text then
    // keeps within half a point of the size it has under a shorter heading.
    const lineOf = (plan: typeof full) => Math.round(((plan.shape('Block 1 heading').box.h * H) / (1.2 * (plan.shape('Block 1 heading').size ?? 1) + 0.15 * (plan.shape('Block 1 heading').size ?? 1))) * 100) / 100;
    assert.deepEqual([lineOf(full), lineOf(short)], [1, 1], `${count} steps`);
    assert.ok(Math.abs((full.shape('Block 1 text').size ?? 0) - (short.shape('Block 1 text').size ?? 0)) <= 0.5, `${count} steps`);
    assert.deepEqual(validateDeck(template, deckOf([steps(budget('b1Heading'))])).filter((issue) => issue.slot !== undefined), [], `${count} steps`);
  }
});

test('a long text alone is set in two columns that come out alike', () => {
  const sentence = 'Bikes unlock from five in the morning, for early shifts and the first trams. ';
  const long = sentence.repeat(10).trim();
  const plan = planOf(freeSlide({ title: 'A season', blocks: [{ text: long }] }));
  const shape = plan.shape('Block 1 text');
  const slot = plan.slot('b1Text') as FreeTextSlot;
  assert.equal(shape.columns, 2);
  // PowerPoint fills the first column before it starts the second, so the box is as high as half the text
  // needs: eight lines of its sixteen. The text may still take the whole room in each column.
  const room = (plan.kit.band.bottom - plan.kit.band.top) * H;
  assert.ok(Math.abs(shape.box.h * H - 8 * 1.2 * 18) < 0.1, `a box ${(shape.box.h * H).toFixed(1)} pt high`);
  assert.deepEqual([slot.drawn.columns, Math.round(slot.drawn.height ?? 0)], [2, Math.round(room)]);
  assert.ok((slot.maxChars ?? 0) > long.length);
  assert.deepEqual(validateDeck(template, deckOf([freeSlide({ title: 'A season', blocks: [{ text: long }] })])), []);
  // A text of a few sentences keeps one column and the whole room.
  const short = planOf(freeSlide({ title: 'A season', blocks: [{ text: sentence.repeat(4).trim() }] })).shape('Block 1 text');
  assert.equal(short.columns, undefined);
});

test('a figure longer than a figure may be is told on its slot, and does not take every label down with it', () => {
  const figures = (third: string) => freeSlide({ title: 'A season in numbers', blocks: [{ figure: { value: '87%', label: 'would recommend it' } }, { figure: { value: '48k', label: 'rides' } }, { figure: { value: third, label: 'ridden in all' } }] });
  const usual = planOf(figures('2.4M'));
  assert.deepEqual([usual.shape('Block 3 value').size, usual.shape('Block 3 label').size, (usual.slot('b3Value') as FreeTextSlot).maxChars], [56.5, 21.5, 6]);
  assert.deepEqual(validateDeck(template, deckOf([figures('2.4M')])), []);
  // A value does not break, so it is still set as small as it has to be to stay in its box. But its
  // budget is that of a figure at four fifths of the usual size, and the labels keep to that size too.
  const long = planOf(figures('2.4 million rides'));
  assert.deepEqual([long.shape('Block 1 value').size, long.shape('Block 3 value').size, long.shape('Block 1 label').size, (long.slot('b3Value') as FreeTextSlot).maxChars], [22.5, 22.5, 17, 8]);
  assert.deepEqual(withoutCodes(validateDeck(template, deckOf([figures('2.4 million rides')]))), [{ severity: 'warning', slide: 1, slot: 'b3Value', message: 'text is 17 characters; the slot fits about 8' }]);
});

test('larger text for a deck never makes a text of a free slide smaller', () => {
  const texts = ['Bikes unlock from five.', 'Bikes unlock from five in the morning, for early shifts.'];
  for (const kind of ['card', 'step'] as const) {
    for (const count of [3, 4, 5, 6, 8]) {
      for (const text of texts) {
        const slide = freeSlide({ title: 'Ways in', blocks: Array.from({ length: count }, (_, index): FreeBlockContent => (kind === 'card' ? { card: { heading: `Point ${index + 1}`, text } } : { step: { heading: `Open the docks ${index + 1}`, text } })) });
        const at = (textScale: number) => new Map(planOf(slide, { textScale }).plan.shapes.filter((shape) => shape.kind === 'text').map((shape) => [shape.name, shape.size ?? 0]));
        const plain = at(1);
        for (const textScale of [1.25, 1.5]) {
          for (const [name, size] of at(textScale)) assert.ok(size >= (plain.get(name) ?? 0), `${count} ${kind}s at ${textScale}: ${name} is ${size} pt, and ${plain.get(name)} pt at 1`);
        }
      }
    }
  }
  // Where there is room, larger text is larger: three cards have it.
  const three = (textScale: number) => planOf(threeCards, { textScale }).shape('Block 1 text').size ?? 0;
  assert.ok(three(1.25) > three(1));
  // A chart's labels and a table's text are part of it.
  const mixed = freeSlide({ title: 'Stations', blocks: [{ chart, caption: 'Rides' }, { table: { header: ['Station', 'Docks'], rows: Array.from({ length: 9 }, (_, row) => [`Station ${row + 1}`, '24']) }, caption: 'Docks' }] });
  const sizesOf = (textScale: number) => planOf(mixed, { textScale }).plan.pattern.slots.map((slot) => (slot.kind === 'chart' ? (slot.text?.size ?? 0) : slot.kind === 'table' ? (slot.bodyText?.size ?? 0) : 0));
  sizesOf(1.5).forEach((size, index) => assert.ok(size >= (sizesOf(1)[index] as number), `slot ${index}`));
});

test('a slide does not lose a deck\'s larger text for half a point, and one that has no room for it is told', () => {
  // Nine blocks in tiles: their cards need half as much height again as a tile of the usual room has, so the
  // slide is crowded already without a wish, and takes the larger room. Its card headings keep their size, and
  // a larger heading would leave the card text less room than it has without the wish: there is no room for
  // what a deck asks at 1.25 or at 1.5. The deck is told, and no text is ever set smaller than it is without
  // the wish.
  const two: ChartValue = { type: 'column', categories: ['Spring', 'Summer', 'Autumn', 'Winter'], series: [{ name: '2024', values: [10, 14, 12, 18] }, { name: '2025', values: [8, 11, 15, 13] }] };
  const nine = freeSlide({
    title: 'Nine mixed blocks',
    blocks: [
      { chart: two, caption: 'Station 1' },
      { chart: two, caption: 'Station 2' },
      card('Longer hours', 'Bikes unlock from five in the morning, for early shifts and the first trams.'),
      card('Safer docks', 'Every station gets lighting and a camera before the season opens.'),
      card('Group passes', 'One pass for up to five riders, bought once and shared by link.'),
      { figure: { value: '48k', label: 'rides between April and October' } },
      { figure: { value: '87%', label: 'riders who would recommend it' } },
      { text: 'Bikes unlock from five in the morning, for early shifts and the first trams.' },
      { note: 'The bottom line: dependable docks turned a weekend habit into a daily commute.' },
    ],
  });
  assert.equal(planOf(nine).plan.layout.dense, true);
  const sizes = (textScale: number) => new Map(planOf(nine, { textScale }).plan.shapes.filter((shape) => shape.kind === 'text').map((shape) => [shape.name, shape.size ?? 0]));
  const plain = sizes(1);
  for (const [textScale, enlarged, said] of [
    [1.25, 'none', 'no room for it'],
    [1.5, 'none', 'no room for it'],
  ] as const) {
    const resolved = planOf(nine, { textScale });
    const larger = sizes(textScale);
    assert.equal(resolved.enlarged, enlarged, `at ${textScale}`);
    for (const [name, size] of larger) assert.ok(size >= (plain.get(name) ?? 0), `at ${textScale}: ${name} is ${size} pt, and ${plain.get(name)} pt at 1`);
    // With half the wish the bottom line and the captions are larger; with none the slide is as it is without it.
    assert.deepEqual(['Block 9 note', 'Block 1 caption'].map((name) => (larger.get(name) ?? 0) > (plain.get(name) ?? 0)), [false, false], `at ${textScale}`);
    const issues = validateDeck(template, deckOf([nine], { textScale }));
    assert.deepEqual(issues.filter((issue) => /larger text/.test(issue.message)).map((issue) => issue.message), [`slide 1 holds 9 blocks, as asked; the deck asks for larger text, and this slide has ${said}`]);
  }
  // Three cards full of text: larger headings would leave the text under them less room than it needs, so
  // its size would have to come down. The slide is made without the wish, and the deck is told so, as a
  // note without a slide number: no rewriting is asked for.
  const full = 'Bikes unlock from five in the morning for early shifts and the first trams and every station gets lighting and a camera before the season opens so that riders feel safe at any hour Bikes unlock from';
  const filled = (heading: string) => freeSlide({ title: 'Ways in', blocks: Array.from({ length: 3 }, () => card(heading, full)) });
  const kept = planOf(filled('Bikes unlock'), { textScale: 1.5 });
  assert.equal(kept.enlarged, 'none');
  assert.deepEqual(kept.plan.shapes.map((shape) => shape.size), planOf(filled('Bikes unlock')).plan.shapes.map((shape) => shape.size));
  const said = (slide: SlideSpec, textScale: number) => validateDeck(template, deckOf([slide], { textScale })).filter((issue) => /larger text/.test(issue.message));
  assert.deepEqual(withoutCodes(said(filled('Bikes unlock'), 1.5)), [{ severity: 'warning', message: 'slide 1 holds 3 blocks, as asked; the deck asks for larger text, and this slide has no room for it' }]);
  // With a heading of two lines over a text of 160 characters there is room for half of what a deck asks at
  // 1.5: the heading is larger, the text as it was.
  const lesser = freeSlide({ title: 'Ways in', blocks: Array.from({ length: 3 }, () => card('Bikes unlock from five in', full.slice(0, 160))) });
  const halved = planOf(lesser, { textScale: 1.5 });
  assert.deepEqual([halved.enlarged, halved.shape('Block 1 heading').size, planOf(lesser).shape('Block 1 heading').size, halved.shape('Block 1 text').size, planOf(lesser).shape('Block 1 text').size], ['half', 24, 19, 18, 18]);
  assert.deepEqual(withoutCodes(said(lesser, 1.5)), [{ severity: 'warning', message: 'slide 1 holds 3 blocks, as asked; the deck asks for larger text, and this slide has room for half of that' }]);
  // Three cards have the room, and nothing is said.
  assert.equal(planOf(threeCards, { textScale: 1.5 }).enlarged, undefined);
  assert.deepEqual(validateDeck(template, deckOf([threeCards], { textScale: 1.5 })), []);
});

test('photos, graphics and icons of a free slide are placed, and found again in the file', () => {
  const slide = freeSlide({ title: 'The pier', blocks: [{ photo: { asset: 'pier', alt: 'The pier at dawn' }, caption: 'At dawn' }, { card: { heading: 'Hours', text: 'From five.', icon: 'clock' } }, { card: { heading: 'Docks', text: 'Lit at night.', icon: 'shield' } }] });
  const deck = deckOf([slide]);
  const assets = { pier: photo, clock: icon('clock'), shield: icon('shield') };
  assert.deepEqual(validateDeck(template, deck, { assets: Object.keys(assets) }), []);
  // Without the picture, the deck's check says which one is missing.
  assert.match(validateDeck(template, deck, { assets: ['clock', 'shield'] }).map((issue) => `${issue.slot}: ${issue.message}`).join('\n'), /b1Photo: .*pier/);
  const { pptx } = renderDeck(template, deck, { assets });
  const xml = slideXml(pptx, 1);
  // The logo of the frame, the photo and the two icons.
  assert.equal((xml.match(/<p:pic\b/g) ?? []).length, 4);
  assert.match(xml, /descr="The pier at dawn"/);
  const found = extractAssets(template, deck, pptx);
  assert.deepEqual(Object.keys(found).sort(), ['clock', 'pier', 'shield']);
  assert.deepEqual(found.pier, photo);
});

test('a free slide in a deck read from the right is mirrored and set right to left', () => {
  const hebrew = freeSlide({ title: 'שלוש דרכים', blocks: [card('שירות מהיר', 'מענה ללקוח בתוך דקות ספורות.'), card('אבטחה מתקדמת', 'הגנה על כל פעולה.'), card('ליווי אישי', 'בנקאי שמכיר את הלקוח.')] });
  const deck = deckOf([hebrew], { direction: 'rtl', language: 'he' });
  const resolved = freeFor(template, deck, hebrew);
  assert.ok(typeof resolved !== 'string');
  const panels = resolved.plan.shapes.filter((shape) => shape.kind === 'panel').map((shape) => shape.box.x);
  assert.deepEqual(panels, panels.slice().sort((a, b) => b - a), 'block 1 is the rightmost');
  const xml = slideXml(renderDeck(template, deck).pptx, 1);
  const made = xml.slice(xml.indexOf('name="Block 1 heading"'));
  assert.match(made, /<a:pPr[^>]*\brtl="1"/);
  assert.match(made, /<a:pPr[^>]*\balgn="r"/);
});

test('a free slide takes the colours of its deck, and may have its own', () => {
  const figures = freeSlide({ title: 'A season in numbers', blocks: [{ figure: { value: '48k', label: 'rides' } }, { figure: { value: '87%', label: 'would recommend it' } }] });
  const plain = slideXml(renderDeck(template, deckOf([figures])).pptx, 1);
  assert.match(plain, /0E7C7B/i);
  const blue = slideXml(renderDeck(template, deckOf([figures], { colors: { accent: '1565C0' } })).pptx, 1);
  assert.doesNotMatch(blue, /0E7C7B/i);
  assert.match(blue, /1565C0/i);
  const own = renderDeck(template, deckOf([figures, { ...figures, colors: { accent: 'B3261E' } }])).pptx;
  assert.match(slideXml(own, 1), /0E7C7B/i);
  assert.doesNotMatch(slideXml(own, 2), /0E7C7B/i);
});

// ---- Checking -----------------------------------------------------------------------------------

test('a free slide is checked without a template for what needs none', () => {
  const good = deckOf([threeCards]);
  assert.deepEqual(validateDeck(manifest, good), []);
  const missing: SlideSpec = { pattern: 'free', blocks: ['card', 'chart'], slots: { b1Heading: 'Hours', b2Chart: 'not a chart' } };
  const issues = validateDeck(manifest, deckOf([missing]));
  assert.deepEqual(issues.filter((issue) => issue.severity === 'error').map((issue) => issue.slot).sort(), ['b1Text', 'b2Chart']);
  const unknown = validateDeck(manifest, deckOf([{ pattern: 'free', blocks: ['card', 'poem'] as unknown as FreeBlockKind[], slots: {} }]));
  assert.match(unknown[0]?.message ?? '', /"blocks" names "poem"; a free slide is made of text, list, card/);
  const empty = validateDeck(manifest, deckOf([{ pattern: 'free', blocks: [], slots: {} }]));
  assert.match(empty[0]?.message ?? '', /a free slide needs at least one block/);
  assert.throws(() => renderDeck(template, deckOf([missing])), DeckValidationError);
});

test('with the template, text that does not fit its block is told by slot, and a slot of no block is ignored with a warning', () => {
  const long = 'Bikes unlock from five in the morning, for early shifts and the first trams. '.repeat(12);
  const slide = freeSlide({ title: 'Three ways in', blocks: [card('Longer hours', long), card('Safer docks', 'Every station gets a light.'), card('Group passes', 'One pass for five riders.')] });
  const issues = validateDeck(template, deckOf([{ ...slide, slots: { ...slide.slots, b9Text: 'nobody home' } }]));
  assert.deepEqual(issues.filter((issue) => issue.severity === 'error'), []);
  assert.ok(issues.some((issue) => issue.slot === 'b1Text' && /characters/.test(issue.message)), JSON.stringify(issues));
  assert.ok(issues.some((issue) => issue.slot === 'b9Text' && /no such slot/.test(issue.message)), JSON.stringify(issues));
  // An optional part without content is simply not there, in strict layout as in adaptive.
  assert.deepEqual(validateDeck(template, deckOf([freeSlide({ blocks: [{ chart }, { photo: 'pier' }] })]), { layout: 'strict', assets: ['pier'] }), []);
});

test('a value that is no text is told as "expected text" on its slot, with the template and without', () => {
  const cases: Array<[string, SlideSpec, string]> = [
    ['a figure given as a number', { pattern: 'free', blocks: ['figure'], slots: { b1Value: 87 as never, b1Label: 'would recommend it' } }, 'b1Value'],
    ['a list that holds a number', { pattern: 'free', blocks: ['list'], slots: { b1Items: ['Docks', 24] as never } }, 'b1Items'],
    ['a list that holds nothing', { pattern: 'free', blocks: ['list'], slots: { b1Items: ['Docks', null] as never } }, 'b1Items'],
    ['a card text given as an object', { pattern: 'free', blocks: ['card'], slots: { b1Heading: 'Hours', b1Text: { text: 'From five.' } as never } }, 'b1Text'],
    ['a note given as true', { pattern: 'free', blocks: ['note'], slots: { b1Text: true as never } }, 'b1Text'],
    ['a title given as a year', { pattern: 'free', blocks: ['text'], slots: { title: 2026 as never, b1Text: 'One line.' } }, 'title'],
  ];
  for (const [name, slide, slot] of cases) {
    const deck = deckOf([slide]);
    for (const target of [manifest, template]) {
      assert.deepEqual(validateDeck(target, deck).filter((issue) => issue.severity === 'error').map((issue) => `${issue.slot}: ${issue.message}`), [`${slot}: expected text`], name);
    }
    assert.throws(() => renderDeck(template, deck), DeckValidationError, name);
  }
});

test('characters that no file can hold are left out of every part of a free slide', () => {
  const slide = freeSlide({
    title: 'Rides by\u0001 station',
    blocks: [
      { text: 'First line\u000Bsecond line\u0002' },
      { chart: { type: 'column', categories: ['Spring\u000B', 'Sum\u000Cmer'], series: [{ name: 'Rides\u000B2025', values: [7, 19] }] } },
      { table: { header: ['Station', 'Docks\uFFFE'], rows: [['Tram\u000BSquare', '24\uFFFF']] } },
      { photo: { asset: 'pier', alt: 'The pier\u000Bat dawn\uD800' } },
    ],
  });
  const deck = deckOf([slide], { fields: { organizationDepartment: 'River\u0003bend' } });
  assert.deepEqual(validateDeck(template, deck, { assets: ['pier'] }).filter((issue) => issue.severity === 'error'), []);
  const all = files(renderDeck(template, deck, { assets: { pier: photo } }).pptx);
  const sheet = strFromU8(unzipSync(all['ppt/embeddings/Microsoft_Excel_Sheet1.xlsx'] as Uint8Array)['xl/worksheets/sheet1.xml'] as Uint8Array);
  const parts = Object.entries(all).filter(([name]) => /\.(xml|rels)$/.test(name)).map(([name, bytes]): [string, string] => [name, strFromU8(bytes)]);
  // U+FFFD is what half a surrogate pair becomes when it is written as it is.
  for (const [name, text] of [...parts, ['the workbook', sheet] as [string, string]]) assert.doesNotMatch(text, /[\u0000-\u0008\u000B\u000C\u000E-\u001F\uFFFD\uFFFE\uFFFF]/, name);
  const xml = slideXml(renderDeck(template, deck, { assets: { pier: photo } }).pptx, 1);
  const texts = textsOf(xml);
  // A vertical tab is the line break of a text box: a new paragraph in a text, a space in one line.
  for (const text of ['Riverbend', 'Rides by station', 'First line', 'second line', 'Tram Square', 'Docks', '24']) assert.ok(texts.includes(text), `${text}: ${JSON.stringify(texts)}`);
  assert.match(xml, /descr="The pier at dawn"/);
  assert.match(xml, /descr="Column chart of Rides 2025: Spring 7, Summer 19"/);
  const drawn = strFromU8(all['ppt/charts/chart1.xml'] as Uint8Array);
  for (const label of ['<c:v>Spring</c:v>', '<c:v>Summer</c:v>', '<c:v>Rides 2025</c:v>']) assert.ok(drawn.includes(label), label);
  assert.ok(sheet.includes('>Summer</t>') && sheet.includes('>Rides 2025</t>'));
});

test('many notes on a slide are set smaller with their bands, and the deck is told what is small', () => {
  const long = 'Rides grew at every station this season, and the docks next to the tram stops emptied first on every weekday morning of the year.';
  const cards = [card('Longer hours', 'Bikes unlock from five.'), card('Safer docks', 'Every station gets a light.'), card('Group passes', 'One pass for five riders.')];
  for (const count of [2, 4, 6, 9]) {
    const slide = freeSlide({ title: 'What we learned', blocks: [...cards, ...Array.from({ length: count }, () => ({ note: long }))] });
    const plan = planOf(slide);
    // No block ends below the room the slide has, and the cards keep their height.
    for (const block of plan.plan.layout.blocks) assert.ok(block.box.y + block.box.h <= plan.kit.dense.bottom + 0.0006, `${count} notes: block ${block.index} ends at ${block.box.y + block.box.h}`);
    assert.ok((plan.plan.layout.blocks[0]?.box.h ?? 0) > 0.2, `${count} notes`);
    for (let block = 4; block <= 3 + count; block++) {
      const note = plan.shape(`Block ${block} note`);
      const slot = plan.slot(`b${block}Text`) as FreeTextSlot;
      // The text of a note fits the band it was given: its slot holds what was written at the size it is set in.
      assert.ok((note.size ?? 0) <= 17 && (slot.maxChars ?? 0) >= long.length, `${count} notes: ${note.size} pt, budget ${slot.maxChars}`);
      assert.ok(note.box.h * H >= 1.2 * (note.size ?? 0), `${count} notes: a band ${(note.box.h * H).toFixed(1)} pt high for ${note.size} pt text`);
    }
    assert.deepEqual(validateDeck(template, deckOf([slide])).filter((issue) => issue.slot !== undefined), [], `${count} notes`);
  }
  // Two long notes still have the template's own size; twelve are small, and the deck is told so.
  const withNotes = (count: number) => freeSlide({ title: 'What we learned', blocks: [...cards, ...Array.from({ length: count }, () => ({ note: long }))] });
  assert.deepEqual([2, 4, 6, 9, 12].map((count) => planOf(withNotes(count)).shape('Block 4 note').size), [17, 13, 12, 8.5, 6.5]);
  assert.deepEqual(withoutCodes(validateDeck(template, deckOf([withNotes(12)]))), [{ severity: 'warning', message: 'slide 1 holds 15 blocks, as asked; its smallest text is set at 6.5 pt' }]);
  // Twelve notes and nothing else share the whole room.
  const alone = planOf(freeSlide({ title: 'What we learned', blocks: Array.from({ length: 12 }, () => ({ note: 'Rides grew at every station.' })) }));
  assert.ok(Math.max(...alone.plan.layout.blocks.map((block) => block.box.y + block.box.h)) <= alone.kit.dense.bottom + 0.0006);
});

test('how small the notes of a slide are set follows from how much they say: short notes are never smaller than long ones', () => {
  const words = 'Rides grew at every station this season, and the docks next to the tram stops emptied first on every weekday morning of the year, in rain and in sun.';
  const under = (lengths: number[]) => planOf(freeSlide({ title: 'Rides by season', blocks: [{ chart }, ...lengths.map((length) => ({ note: `${words} ${words}`.slice(0, length).trim() }))] }));
  const sizesOf = (lengths: number[]) => lengths.map((_, index) => under(lengths).shape(`Block ${index + 2} note`).size ?? 0);
  // Four notes under a chart have two fifths of the room. Set smaller, a note of 110 characters takes one
  // line where it took two, so it asks for no more than a note of 40: both come out alike. A note of 150
  // takes one line from 11 pt down, where a line holds all of it at the width of a note's heavy face.
  assert.deepEqual([40, 110, 150].map((length) => sizesOf(Array<number>(4).fill(length))[0]), [11.5, 11.5, 11]);
  // More notes are never set larger than fewer, and notes that say less never smaller than notes that say more.
  for (const length of [40, 110, 150]) {
    const byCount = [1, 2, 3, 4, 5, 6, 8, 12].map((count) => sizesOf(Array<number>(count).fill(length))[0] as number);
    byCount.forEach((size, index) => assert.ok(index === 0 || size <= (byCount[index - 1] as number), `${length} characters: ${byCount.join(', ')}`));
  }
  for (const count of [2, 3, 4, 5, 6, 8, 12]) {
    const byLength = [20, 40, 80, 110, 150].map((length) => sizesOf(Array<number>(count).fill(length))[0] as number);
    byLength.forEach((size, index) => assert.ok(index === 0 || size <= (byLength[index - 1] as number), `${count} notes: ${byLength.join(', ')}`));
  }
  // The notes of one slide share one size, whatever each says, and each band is as high as its own lines.
  const mixed = under([30, 110, 240, 60]);
  assert.deepEqual(sizesOf([30, 110, 240, 60]), [10.5, 10.5, 10.5, 10.5]);
  const bands = mixed.plan.layout.blocks.slice(1).map((block) => Math.round(block.box.h * H * 10) / 10);
  assert.ok((bands[2] as number) > 1.3 * (bands[0] as number) && bands[0] === bands[1] && bands[0] === bands[3], bands.join(', '));
  // Space is given up before size: the gap between two notes comes down more than their text does.
  const [first, second] = mixed.plan.layout.blocks.slice(1);
  const gap = ((second?.box.y ?? 0) - (first?.box.y ?? 0) - (first?.box.h ?? 0)) * H;
  assert.ok(gap / (mixed.kit.gap * W) < 10 / 17 && gap > 3, `a gap of ${gap.toFixed(1)} pt`);
  for (const lengths of [[40, 40, 40, 40], [30, 110, 150, 60], Array<number>(12).fill(150)]) {
    const plan = under(lengths);
    for (const block of plan.plan.layout.blocks) assert.ok(block.box.y + block.box.h <= plan.kit.dense.bottom + 0.0006);
    // Every note holds what was written for it.
    assert.deepEqual(validateDeck(template, deckOf([freeSlide({ title: 'Rides by season', blocks: [{ chart }, ...lengths.map((length) => ({ note: words.slice(0, length).trim() }))] })])).filter((issue) => issue.slot !== undefined), [], lengths.join());
  }
});

test('a block that ends below the room of its slide is told, as a note on the deck', () => {
  const deck = deckOf([threeCards]);
  // Nothing lays a block out there. Should a layout do it all the same, the check does not keep quiet.
  const lowered = (slide: SlideSpec) => {
    const resolved = freeFor(template, deck, slide);
    assert.ok(typeof resolved !== 'string');
    const blocks = resolved.plan.layout.blocks.map((block) => (block.index === 1 ? block : { ...block, box: { ...block.box, y: 0.9 } }));
    return { ...resolved, plan: { ...resolved.plan, layout: { ...resolved.plan.layout, blocks } } };
  };
  assert.deepEqual(withoutCodes(checkDeck(manifest, deck, {}, lowered)), [{ severity: 'warning', message: 'slide 1 holds 3 blocks, as asked; 2 of them reach below the room the slide has, and may be cut off at its edge' }]);
  assert.deepEqual(validateDeck(template, deck), []);
});

test('the notes on a crowded slide give sizes in half points and count one block as one', () => {
  const rows = (count: number) => ({ header: ['Station', 'Docks'], rows: Array.from({ length: count }, (_, row) => [`Station ${row + 1}`, '24']) });
  // The demo slide is 959.98 pt wide, so a size that follows the width is no round number before it is written.
  assert.deepEqual(withoutCodes(validateDeck(template, deckOf([freeSlide({ title: 'Stations', blocks: [{ table: rows(30) }] })]))), [{ severity: 'warning', message: 'slide 1 holds 1 block, as asked; its smallest text is set at 6 pt' }]);
  const charts = freeSlide({ title: 'Stations', blocks: Array.from({ length: 200 }, (_, index) => ({ chart, caption: `Station ${index + 1}` })) });
  assert.deepEqual(validateDeck(template, deckOf([charts])).map((issue) => issue.message), ['slide 1 holds 200 blocks, as asked; its smallest text is set at 5 pt', 'slide 1 holds 200 blocks, as asked; 200 charts are drawn small, without axes, legend or labels']);
  const single = freeSlide({ blocks: Array.from({ length: 1 }, () => ({ chart })), columns: 1 });
  const tiny = checkDeck(manifest, deckOf([single]), {}, (slide) => {
    const resolved = freeFor(template, {}, slide);
    assert.ok(typeof resolved !== 'string');
    return { ...resolved, plan: { ...resolved.plan, pattern: { ...resolved.plan.pattern, slots: resolved.plan.pattern.slots.map((slot) => (slot.kind === 'chart' ? { ...slot, detail: 'spark' as const } : slot)) } } };
  });
  assert.deepEqual(tiny.map((issue) => issue.message), ['slide 1 holds 1 block, as asked; one chart is drawn small, without axes, legend or labels']);
});

test('what a slot says to write is never more than the slot holds', () => {
  const mixes: FreeBlockKind[][] = [['figure', 'chart'], ['list', 'text', 'chart', 'photo'], ['text'], ['text', 'text'], ['quote'], ['note', 'card', 'card'], ['figure', 'figure', 'figure', 'note']];
  for (const kind of ['card', 'step', 'text', 'list', 'figure'] as const) for (let count = 1; count <= 12; count++) mixes.push(Array<FreeBlockKind>(count).fill(kind));
  let seen = 0;
  for (const blocks of mixes) {
    for (const textScale of [1, 1.25]) {
      const planned = freePattern(template, { textScale }, { blocks }, false);
      assert.ok(typeof planned !== 'string');
      for (const slot of planned.pattern.slots) {
        const most = /Write \d+ to (\d+) characters/.exec(slot.kind === 'text' ? (slot.description ?? '') : '');
        if (!most || slot.kind !== 'text' || /in all$/.test(slot.description ?? '')) continue;
        seen++;
        assert.ok(Number(most[1]) <= (slot.maxChars ?? 0), `${blocks.join()} at ${textScale}: ${slot.name} holds ${slot.maxChars} and says "${slot.description}"`);
      }
    }
  }
  assert.ok(seen > 150);
  // The two budgets the upper end used to round over: 59 and 79.
  const label = freePattern(template, {}, { blocks: ['figure', 'chart'] }, false);
  assert.ok(typeof label !== 'string');
  const slot = label.pattern.slots.find((entry) => entry.name === 'b1Label');
  assert.deepEqual([slot?.kind === 'text' ? slot.maxChars : 0, slot?.description], [42, 'What the figure counts. Write 20 to 40 characters']);
});

test('with fonts, the text of a free slide is measured in the box it is drawn in', () => {
  const measured = loadTemplate(demoPptx, manifest, { fonts: fontkitMetrics(fontkit, openFonts()) });
  const slide = freeSlide({ title: 'Measured', blocks: [card('Longer hours', 'Bikes unlock from five in the morning.'), { list: ['Open earlier', 'Fix flats in a day'] }, { figure: { value: '87%', label: 'would recommend it' } }, { note: 'Dependable docks made it a habit.' }] });
  const plan = planOf(slide, {}, measured);
  const measure = textMeasurer(measured);
  const widths: Record<string, number> = {};
  for (const slot of plan.plan.pattern.slots) {
    if (slot.kind !== 'text' || slot.name === 'title' || !slot.box) continue;
    const drawn = (slot as FreeTextSlot).drawn;
    const fit = measure.fit(plan.plan.pattern, slot, 'Some words to measure');
    assert.ok(fit, slot.name);
    // The width of a line is the box, less the insets the copy is given and how far its bullet hangs:
    // not the insets and the margin of the template shape it is a copy of.
    const outline = slot.name === 'b4Text' ? 3.2 : 0;
    assert.ok(Math.abs((fit.lineRoom[0] as number) - (slot.box.w * W - 2 * drawn.inset - (drawn.hang ?? 0) - outline)) < 0.6, `${slot.name}: ${fit.lineRoom[0]} pt of ${slot.box.w * W} pt, ${JSON.stringify(drawn)}`);
    widths[slot.name] = Math.round(fit.lineRoom[0] as number);
  }
  assert.deepEqual(widths, { b1Heading: 230, b1Text: 230, b2Items: 212, b3Value: 230, b3Label: 230, b4Text: 846 });
  assert.equal((plan.slot('b2Items') as FreeTextSlot).drawn.hang, 1.1 * 17);
  assert.equal((plan.slot('b3Value') as FreeTextSlot).drawn.wrap, false);
  // A text in two columns is measured in one of them, and has the room of both.
  const long = 'Bikes unlock from five in the morning, for early shifts and the first trams. '.repeat(10).trim();
  const columns = planOf(freeSlide({ title: 'A season', blocks: [{ text: long }] }), {}, measured);
  const fit = measure.fit(columns.plan.pattern, columns.slot('b1Text') as FreeTextSlot, long);
  assert.deepEqual([fit?.fits, Math.round(fit?.lineRoom[0] ?? 0), Math.round(fit?.room ?? 0)], [true, 432, 626]);
  assert.deepEqual(validateDeck(measured, deckOf([slide, freeSlide({ title: 'A season', blocks: [{ text: long }] })])), []);
});

test('the other checks of a deck know a free slide', () => {
  const stated = freeSlide({ title: 'Rides', blocks: [{ chart }, { chart: { ...chart, series: [{ name: '2025', values: [7, 21] }] } }] });
  const deck = deckOf([threeCards, stated], { brief: 'Rides grew from 7 thousand in spring to 19 thousand in summer.' });
  // The second chart shows a number the brief never gave.
  const unstated = unstatedChartValues(manifest, deck);
  assert.deepEqual(unstated.map((issue) => [issue.slide, issue.slot]), [[2, 'b2Chart']]);
  // A free slide is its own layout: there is no other to show it in.
  assert.deepEqual(slideAlternatives(manifest, deck, 1), []);
  // Measured with fonts, a free slide that fits has nothing to say.
  const measured = loadTemplate(demoPptx, manifest, { fonts: fontkitMetrics(fontkit, openFonts()) });
  assert.deepEqual(validateDeck(measured, deckOf([threeCards])), []);
  assert.deepEqual(checkDesign(measured, deckOf([threeCards])).filter((issue) => issue.slide === 1), []);
});

// ---- A studio's settings ---------------------------------------------------------------------

test('a studio names the frame and the styles of free slides, and can turn them off', () => {
  const off = changeManifest(template, [{ op: 'setManifest', free: false }]);
  assert.equal(off.manifest.free, false);
  assert.equal(off.applied[0]?.message, 'The template offers no free slides when a deck is planned.');

  const set = changeManifest(template, [{ op: 'setManifest', free: { base: 'text-only', styles: { heading: { slide: 7, shape: '7' } } } }]);
  assert.deepEqual(set.rejected, []);
  assert.deepEqual(set.manifest.free, { base: 'text-only', styles: { heading: { slide: 7, shape: '7' } } });
  const kit = freeKit(loadTemplate(demoPptx, set.manifest));
  assert.deepEqual([kit?.base, kit?.styles.heading.slide, kit?.styles.heading.shapeId], ['text-only', 7, '7']);
  // Giving the choices back leaves no setting behind.
  const back = changeManifest(loadTemplate(demoPptx, set.manifest), [{ op: 'setManifest', free: { base: null, styles: { heading: null } } }]);
  assert.equal(back.manifest.free, undefined);
  assert.equal(changeManifest(loadTemplate(demoPptx, off.manifest), [{ op: 'setManifest', free: true }]).manifest.free, undefined);

  // What a free slide cannot be made with is refused, with the reason.
  const refused = (free: unknown) => changeManifest(template, [{ op: 'setManifest', free } as never]).rejected[0]?.reason ?? '';
  assert.match(refused({ base: 'two-columns' }), /a free slide cannot stand on "two-columns": the layout "two-columns" has a row of columns, which the software lays out itself; name a layout with a title that leaves room under it and has no row of items, such as "list-with-photo"/);
  assert.match(refused({ base: 'nope' }), /there is no pattern "nope"/);
  assert.match(refused({ styles: { poem: { slide: 6, shape: '7' } } }), /a free slide has no part called "poem"/);
  assert.match(refused({ styles: { heading: { slide: 4, shape: '2' } } }), /is a picture; the heading of a free slide is copied from a text box/);
  assert.match(refused({ styles: { panel: { slide: 6, shape: '7' } } }), /cannot be the panel of a free slide/);
});

test('a manifest whose settings point at nothing is told when it is checked', () => {
  const broken = { ...manifest, colors: { accent: 'nope' }, free: { base: 'gone', styles: { poem: { slide: 1, shape: '2' }, heading: { slide: 99, shape: '2' } } } };
  assert.deepEqual(
    checkManifest({ pkg: template.pkg, manifest: broken }).map((issue) => `${issue.severity}: ${issue.message}`),
    [
      'error: colors.accent is "nope"; a colour is six hex digits such as "0E7C7B"',
      'error: free.base names the pattern "gone", which the manifest does not have',
      'warning: free.styles names "poem"; a free slide is made of body, heading, figure, caption, list, quote, note, panel',
      'error: free.styles.heading points at shape #2 of slide 99, which the template does not have',
    ],
  );
  assert.deepEqual(checkManifest(template), []);
});

test('a manifest that was changed is what free slides are made with, on the package that is open', () => {
  const slide = freeSlide({ title: 'Rides', blocks: [{ chart }, { table: { header: ['Station', 'Docks'], rows: [['Tram Square', '24']] } }] });
  const deck = deckOf([slide]);
  const seriesOf = (pptx: Uint8Array) => [...(chartsOf(pptx, 1)[0] as string).matchAll(/<c:ser>.*?<a:srgbClr val="([0-9A-F]{6})"/g)].map((match) => match[1]);
  const sizeOf = (pptx: Uint8Array) => /<a:tc>.*?\bsz="(\d+)"/.exec(slideXml(pptx, 1))?.[1];
  const opened = loadTemplate(demoPptx, manifest);
  const before = renderDeck(opened, deck).pptx;
  assert.deepEqual([seriesOf(before), sizeOf(before)], [['0E7C7B'], '1500']);
  // A studio changes the colours of the chart and the text size of the table.
  const changed = changeManifest(opened, [
    { op: 'setSlot', pattern: 'chart-native', slot: 'chart', colors: ['AA0000', '0000AA'] },
    { op: 'setSlot', pattern: 'table', slot: 'table', bodyText: { size: 10 } },
  ] as never);
  assert.deepEqual(changed.rejected, []);
  const fresh = renderDeck(loadTemplate(demoPptx, changed.manifest), deck).pptx;
  assert.deepEqual([seriesOf(fresh), sizeOf(fresh)], [['AA0000'], '1000']);
  // The new manifest on the package that is open, as a studio goes on: the same file as after a fresh load.
  assert.deepEqual(renderDeck({ pkg: opened.pkg, manifest: changed.manifest }, deck).pptx, fresh);
  // And on the same template object.
  opened.manifest = changed.manifest;
  assert.deepEqual(renderDeck(opened, deck).pptx, fresh);
  // The frame follows the manifest too, and how many cards a planner is told read well.
  const kept = loadTemplate(demoPptx, manifest);
  assert.equal(freeKit(kept)?.base, 'list-with-photo');
  assert.equal(freeOffer(kept, { figures: true, photos: true, graphics: true, icons: true })?.base, 'list-with-photo');
  kept.manifest = changeManifest(kept, [{ op: 'setManifest', free: { base: 'text-only' } }]).manifest;
  assert.equal(freeKit(kept)?.base, 'text-only');
  assert.equal(freeOffer(kept, { figures: true, photos: true, graphics: true, icons: true })?.base, 'text-only');
});

test('a style of free slides that can no longer be copied is told when the manifest is checked', () => {
  // The heading of free slides is set to the quotation's text. Then the quotation layout is taken out.
  const set = changeManifest(template, [{ op: 'setManifest', free: { styles: { heading: { slide: 11, shape: '6' } } } }]);
  assert.deepEqual([set.rejected, set.issues], [[], []]);
  const without = changeManifest(loadTemplate(demoPptx, set.manifest), [{ op: 'removePattern', pattern: 'quote' }]);
  assert.deepEqual(without.rejected, []);
  assert.deepEqual(without.manifest.free, { styles: { heading: { slide: 11, shape: '6' } } });
  const loaded = loadTemplate(demoPptx, without.manifest);
  // The setting is passed over when a slide is made, so the check says that it is.
  assert.deepEqual([freeKit(loaded)?.styles.heading.slide, freeKit(loaded)?.styles.heading.shapeId], [5, '10']);
  const told = [{ severity: 'warning', message: "free.styles.heading points at shape #6 of slide 11, which no layout of the manifest uses as a text that can be copied; the software's own choice is used" }];
  assert.deepEqual(checkManifest(loaded), told);
  assert.deepEqual(without.issues, told);
  // Styles that are used say nothing, a panel among them.
  const good = changeManifest(template, [{ op: 'setManifest', free: { styles: { heading: { slide: 7, shape: '7' }, panel: { slide: 5, shape: '7' } } } }]);
  assert.deepEqual([good.rejected, good.issues], [[], []]);
  // A manifest written by hand may hold anything: it is reported, and nothing throws.
  const broken = { pkg: template.pkg, manifest: { ...manifest, free: { styles: { heading: null, body: 'the grey one' } } } as unknown as TemplateManifest };
  assert.deepEqual(
    checkManifest(broken).map((issue) => `${issue.severity}: ${issue.message}`),
    ['error: free.styles.heading is null; it names a shape of the template, such as { "slide": 7, "shape": "7" }', 'error: free.styles.body is "the grey one"; it names a shape of the template, such as { "slide": 7, "shape": "7" }'],
  );
  assert.equal(freeKit(broken)?.styles.heading.shapeId, '10');
  assert.deepEqual(validateDeck(broken, deckOf([threeCards])), []);
});

test('running text that is all in placeholders cannot be copied, and the reason says so', () => {
  // The demo with the content texts of every slide made body placeholders that keep their position.
  const keep = /^(Title|Organization|Page number)$/;
  let index = 20;
  const held = (xml: string) => xml.replace(/<p:cNvPr id="(\d+)" name="([^"]*)"\/><p:cNvSpPr txBox="1"\/><p:nvPr\/>/g, (whole, id: string, name: string) => (keep.test(name) ? whole : `<p:cNvPr id="${id}" name="${name}"/><p:cNvSpPr/><p:nvPr><p:ph type="body" idx="${index++}"/></p:nvPr>`));
  const slides = Object.fromEntries(Array.from({ length: 13 }, (_, number) => [`ppt/slides/slide${number + 1}.xml`, held]));
  const placed = loadTemplate(demoWith(slides), manifest);
  const reason = freeKitOrReason(placed);
  assert.equal(reason, "the template's running text is in placeholders, which a free slide cannot copy: there is no text style to take unless some of it is set in text boxes of their own");
  assert.deepEqual(validateDeck(placed, deckOf([threeCards])), [{ severity: 'error', slide: 1, message: `a free slide cannot be made from this template: ${reason}` }]);
  // One placeholder among text boxes: it cannot be named as a style, and a studio is told why.
  const one = loadTemplate(demoWith({ 'ppt/slides/slide11.xml': held }), manifest);
  assert.ok(typeof freeKitOrReason(one) !== 'string');
  assert.equal(
    changeManifest(one, [{ op: 'setManifest', free: { styles: { quote: { slide: 11, shape: '6' } } } }]).rejected[0]?.reason,
    'shape #6 of slide 11 cannot be the quote of a free slide: it is a placeholder, which takes its look from the slide layout, and a free slide copies a text box that states its own',
  );
});

test('an instruction to a studio can set the colours and the free slides of a template', async () => {
  const answer = {
    note: '',
    changes: [
      { op: 'setManifest', pattern: '', shape: '', slot: '', values: [{ property: 'accent', value: '#0e7c7b' }, { property: 'freeBase', value: 'text-only' }] },
      { op: 'setManifest', pattern: '', shape: '', slot: '', values: [{ property: 'keepColors', value: 'true' }] },
      { op: 'setManifest', pattern: '', shape: '', slot: '', values: [{ property: 'freeSlides', value: 'false' }] },
    ],
  };
  const requests: JsonRequest[] = [];
  const llm: LlmProvider = {
    generateJson: async <T>(request: JsonRequest) => {
      requests.push(request);
      return answer as T;
    },
  };
  const result = await editManifest(template, 'Our colour is 0E7C7B, stand free slides on the text slide; then lock the colours and turn free slides off.', { llm });
  assert.deepEqual(result.rejected, []);
  assert.deepEqual(result.applied.map((entry) => entry.message), [
    'The template has 0E7C7B as its main colour, for decks that ask for other colours, lays out free slides on the frame of "text-only".',
    'The template keeps its own colours in every deck.',
    'The template offers no free slides when a deck is planned.',
  ]);
  assert.deepEqual([result.manifest.colors, result.manifest.free], [false, false]);
  // The model is told of the settings it may change.
  assert.match(requests[0]?.prompt ?? '', /keepColors \("true" keeps every deck in the template's own colours/);
  assert.match(requests[0]?.prompt ?? '', /freeBase \(the id of the pattern whose frame such slides stand on/);
});

// ---- What a planner is told ---------------------------------------------------------------------

test('a planner is offered the kinds of block a deck can fill, and how many read well', () => {
  const all = freeOffer(template, { figures: true, photos: true, graphics: true, icons: true });
  assert.deepEqual(all, { kinds: ['text', 'list', 'card', 'figure', 'step', 'phase', 'milestone', 'option', 'quote', 'note', 'chart', 'table', 'photo', 'graphic'], comfortable: 3, base: 'list-with-photo' });
  // Without numbers in the brief there is nothing to chart; without pictures, none to show.
  assert.deepEqual(freeOffer(template, { figures: false, photos: false, graphics: false, icons: false })?.kinds, ['text', 'list', 'card', 'step', 'phase', 'milestone', 'option', 'quote', 'note']);
  assert.deepEqual(freeOffer(template, { figures: true, photos: true, graphics: true, icons: true, textOnly: true })?.kinds, ['text', 'list', 'quote', 'note']);
  assert.equal(freeOffer(loadTemplate(demoPptx, { ...manifest, free: false }), { figures: true, photos: true, graphics: true, icons: true }), undefined);
  for (const kind of all?.kinds ?? []) assert.ok(FREE_KIND_NOTES[kind].length > 5, kind);
});

test('a planned free slide is written against slots that say how much each holds', () => {
  const planned = freePattern(template, {}, { blocks: ['card', 'card', 'card'] }, true);
  assert.ok(typeof planned !== 'string');
  assert.equal(planned.base, 'list-with-photo');
  const slots = planned.pattern.slots.filter((slot) => slot.name.startsWith('b'));
  assert.deepEqual(slots.map((slot) => `${slot.name}${slot.empty === 'remove' ? '?' : ''}`), ['b1Icon?', 'b1Heading', 'b1Text', 'b2Icon?', 'b2Heading', 'b2Text', 'b3Icon?', 'b3Heading', 'b3Text']);
  // A card is planned as high as a text of the usual length needs, and its slot says what that box holds.
  const text = slots.find((slot) => slot.name === 'b1Text');
  assert.ok(text?.kind === 'text');
  assert.deepEqual([text.maxChars, text.description], [78, 'Write 40 to 70 characters']);
  // Without an icon set no card is offered one, and like cards hold the same.
  const bare = freePattern(template, {}, { blocks: ['card', 'card', 'card'] }, false);
  assert.ok(typeof bare !== 'string');
  assert.ok(!bare.pattern.slots.some((slot) => slot.name.endsWith('Icon')));
  // Charts of a slide that has several are offered a caption each. What reads at their size is in the
  // fields of the slot, and the description does not say it a second time.
  const charts = freePattern(template, {}, { blocks: Array<FreeBlockKind>(12).fill('chart') }, false);
  assert.ok(typeof charts !== 'string');
  const chartSlot = charts.pattern.slots.find((slot) => slot.name === 'b12Chart');
  assert.ok(chartSlot?.kind === 'chart');
  assert.deepEqual([chartSlot.detail, chartSlot.maxCategories, chartSlot.maxSeries, chartSlot.description], ['compact', 6, 2, 'A small chart. When the brief gives no numbers for it, leave it empty']);
  assert.ok(charts.pattern.slots.some((slot) => slot.name === 'b12Caption' && slot.empty === 'remove'));
  // The slide's own frame and arrangement are followed.
  const wide = freePattern(template, {}, { blocks: ['card', 'card', 'card', 'card'], arrange: { columns: 2 }, base: 'text-only' }, false);
  assert.ok(typeof wide !== 'string');
  assert.equal(wide.base, 'text-only');
});

test('blocks are named to a model as runs of like blocks, and read back', () => {
  assert.deepEqual(blockRuns(['chart', 'chart', 'note', 'chart']), [{ kind: 'chart', number: 2, first: 1 }, { kind: 'note', number: 1, first: 3 }, { kind: 'chart', number: 1, first: 4 }]);
  assert.deepEqual(blocksOfRuns([{ kind: 'chart', number: 2 }, { kind: 'poem', number: 1 }, { kind: 'note', number: 1.7 }, { kind: 'card', number: 0 }, { kind: 'card', number: '2' }]), ['chart', 'chart', 'note', 'card', 'card']);
  assert.deepEqual(blocksOfRuns(blockRuns(['text', 'card', 'card', 'card'])), ['text', 'card', 'card', 'card']);
  // How many blocks a run has is its `count`; a run without one is read by its `number`.
  assert.deepEqual(blocksOfRuns([{ kind: 'chart', count: 3, number: 9 }, { kind: 'note', number: 1 }, { kind: 'card', count: 0, number: 2 }, { kind: 'step', count: 'two' }]), ['chart', 'chart', 'chart', 'note', 'card', 'card']);
  // Whatever the numbers say, no more than two thousand blocks are written out.
  assert.equal(MOST_BLOCKS, 2000);
  assert.equal(blocksOfRuns([{ kind: 'chart', count: 1e9 }]).length, 2000);
  assert.equal(blocksOfRuns([{ kind: 'chart', count: Infinity }]).length, 2000);
  const mixed = blocksOfRuns([{ kind: 'chart', count: 1500 }, { kind: 'note', number: 1500 }, { kind: 'card', count: 5 }]);
  assert.deepEqual([mixed.length, mixed.filter((kind) => kind === 'note').length, mixed.includes('card')], [2000, 500, false]);
});

// ---- What a slot says of itself -------------------------------------------------------------

/** Every kind of block once, and the kinds that take a caption twice, as a slide is planned. */
const everyKind: FreeBlockKind[] = ['text', 'list', 'card', 'figure', 'step', 'quote', 'note', 'chart', 'table', 'photo', 'graphic', 'chart', 'table', 'photo', 'graphic'];

test('a slot says its limits once, in its fields: its description names none of them and ends without a full stop', () => {
  const described: string[] = [];
  const lists = { one: 0, several: 0 };
  for (const blocks of [everyKind, ['chart'], Array<FreeBlockKind>(12).fill('chart'), Array<FreeBlockKind>(200).fill('chart'), ['list'], ['table'], ['card', 'card', 'card'], ['quote'], ['note', 'card'], ['figure', 'figure', 'figure'], ['text'], Array<FreeBlockKind>(24).fill('quote'), Array<FreeBlockKind>(16).fill('card'), Array<FreeBlockKind>(9).fill('list'), Array<FreeBlockKind>(16).fill('list'), Array<FreeBlockKind>(30).fill('list')] as FreeBlockKind[][]) {
    for (const textScale of [1, 1.25]) {
      const planned = freePattern(template, { textScale }, { blocks }, true);
      assert.ok(typeof planned !== 'string');
      for (const slot of planned.pattern.slots.filter((entry) => /^b\d/.test(entry.name))) {
        const said = slot.description;
        if (said === undefined) continue;
        described.push(said);
        const label = `${blocks.length} ${blocks[0]}, ${slot.name}: "${said}"`;
        assert.doesNotMatch(said, /[.\s]$/, label);
        // No limit that a field states: characters, paragraphs, categories, series, columns, rows.
        assert.doesNotMatch(said, /at most|up to \d|\d+ (categor|series|columns|rows|items|paragraphs)|read well here/i, label);
        // A text says how much to write only as a target, and only where there is room for one.
        if (/characters/.test(said)) assert.match(said, /(^|\. )Write \d+ to \d+ characters( in all)?$/, label);
        assert.doesNotMatch(said, /\b1 (characters|blocks|items|points)\b/, label);
        // A list whose slot holds one paragraph is one point, not one in each item, and its target is not "in all".
        if (slot.kind === 'text' && /Items$/.test(slot.name)) {
          const single = slot.maxParagraphs === 1;
          lists[single ? 'one' : 'several']++;
          assert.match(said, single ? /^One short point(\. Write \d+ to \d+ characters)?$/ : /^One short point in each item(\. Write \d+ to \d+ characters in all)?$/, label);
        }
      }
    }
  }
  // What a part is told, where it is told anything besides its limits: every description is one of these.
  const told = new Set(described.map((said) => said.replace(/\d+/g, 'N')));
  const sentences = [
    'A chart. When the brief gives no numbers for it, leave it empty',
    'A heading of a few words',
    'A heading of a few words. Give one to every text of the slide, or to none',
    'A small chart. When the brief gives no numbers for it, leave it empty',
    'A table; a word or a short phrase in each cell. When the brief gives nothing for it, leave it empty',
    'An icon that suits the card. Give one to every card of the slide, or to none',
    'One short point',
    'One short point. Write N to N characters',
    'One short point in each item',
    'One short point in each item. Write N to N characters in all',
    'The figure itself, such as N% or N.NM. When the brief gives no number for it, leave it empty',
    'The one line to take away. Write N to N characters',
    'The quotation, word for word, without quotation marks',
    'The quotation, word for word, without quotation marks. Write N to N characters',
    'What it shows, in a few words. Give one to every chart of the slide, or to none',
    'What it shows, in a few words. Give one to every graphic of the slide, or to none',
    'What it shows, in a few words. Give one to every photo of the slide, or to none',
    'What it shows, in a few words. Give one to every table of the slide, or to none',
    'What the figure counts',
    'What the figure counts. Write N to N characters',
    'Who said it, when the brief says who. Never make a name up',
    'Write N to N characters',
  ];
  for (const said of told) assert.ok(sentences.includes(said), said);
  assert.ok(told.size >= 16, [...told].join(' | '));
  assert.ok(lists.one > 0 && lists.several > 0, JSON.stringify(lists));
  // A count of one is singular in the sentences the checks build.
  const fit = { fits: false, lines: 1, capacity: 0, overflow: 1, height: 20, room: 10, widthRatio: 1, size: 12, sizes: [12], scale: 1, fitsChars: 1, paragraphs: [['x']], widths: [[5]], lineRoom: [10], sampleLines: 0, approximate: false };
  assert.equal(overflowMessage(fit), 'text takes 1 line at 12 pt; the box holds 0 (1 line too many). Text of about 1 character fits.');
  assert.match(validateDeck(template, deckOf([freeSlide({ title: 'Sixty', blocks: Array.from({ length: 1 }, () => ({ table: { header: ['A', 'B'], rows: Array.from({ length: 60 }, () => ['1', '2']) } })) })])).map((issue) => issue.message).join(' | '), /slide 1 holds 1 block, as asked/);
});

test('a chart of a free slide states how many categories and series read at its size, as soft limits', () => {
  const limits = (blocks: FreeBlockKind[]) => {
    const planned = freePattern(template, {}, { blocks }, false);
    assert.ok(typeof planned !== 'string');
    const slot = planned.pattern.slots.find((entry) => entry.name === 'b1Chart');
    assert.ok(slot?.kind === 'chart');
    return [slot.detail ?? 'full', slot.maxCategories, slot.maxSeries];
  };
  assert.deepEqual(limits(['chart']), ['full', 8, 3]);
  assert.deepEqual(limits(Array<FreeBlockKind>(12).fill('chart')), ['compact', 6, 2]);
  assert.deepEqual(limits(Array<FreeBlockKind>(200).fill('chart')), ['spark', 12, 1]);
  // With content the slot states the same, and a chart over it is drawn and told on its slot, like a chart of any layout.
  const many: ChartValue = { type: 'column', categories: ['1', '2', '3', '4', '5', '6', '7', '8', '9', '10'], series: ['a', 'b', 'c', 'd'].map((name) => ({ name, values: [1, 2, 3, 4, 5, 6, 7, 8, 9, 10] })) };
  const deck = deckOf([freeSlide({ title: 'Rides', blocks: [{ chart: many }] })]);
  assert.deepEqual(withoutCodes(validateDeck(template, deck)), [
    { severity: 'warning', slide: 1, slot: 'b1Chart', message: 'the chart has 10 categories; the slot fits 8' },
    { severity: 'warning', slide: 1, slot: 'b1Chart', message: 'the chart has 4 series; the slot fits 3' },
  ]);
  assert.equal(chartsOf(renderDeck(template, deck).pptx, 1).length, 1);
  assert.deepEqual(validateDeck(template, deckOf([freeSlide({ title: 'Rides', blocks: [{ chart }] })])), []);
});

test('a caption is planned for photos and graphics as for charts and tables, and the parts a slide has can be counted in', () => {
  const names = (blocks: FreeBlockKind[], withParts?: string[]) => {
    const planned = freePattern(template, {}, { blocks }, false, withParts);
    assert.ok(typeof planned !== 'string');
    return planned.pattern.slots.map((slot) => slot.name).filter((name) => /^b\d/.test(name));
  };
  // Two or more of a kind are told apart by their captions; one alone has none.
  for (const [kind, part] of [['photo', 'Photo'], ['graphic', 'Graphic']] as const) {
    assert.deepEqual(names([kind, kind]), [`b1${part}`, 'b1Caption', `b2${part}`, 'b2Caption'], kind);
    assert.deepEqual(names([kind, 'text']), [`b1${part}`, 'b2Heading', 'b2Text'], kind);
  }
  assert.deepEqual(names(['chart', 'chart']), ['b1Caption', 'b1Chart', 'b2Caption', 'b2Chart']);
  assert.deepEqual(names(['chart', 'photo', 'table']), ['b1Chart', 'b2Photo', 'b3Table']);
  // A change to a deck names the optional parts the slide has: they are offered again, whatever the rules say.
  assert.deepEqual(names(['photo', 'text'], ['b1Caption']), ['b1Photo', 'b1Caption', 'b2Heading', 'b2Text']);
  assert.deepEqual(names(['text']), ['b1Text']);
  assert.deepEqual(names(['text'], ['b1Heading']), ['b1Heading', 'b1Text']);
  assert.deepEqual(names(['chart', 'note'], ['b1Caption', 'b9Caption', 'b2Text']), ['b1Caption', 'b1Chart', 'b2Text']);
  // A card has an icon when every card of the slide has one: a name for one card alone does not give it one.
  assert.deepEqual(names(['card', 'card'], ['b1Icon']), ['b1Heading', 'b1Text', 'b2Heading', 'b2Text']);
  assert.deepEqual(names(['card', 'card'], ['b1Icon', 'b2Icon']), ['b1Icon', 'b1Heading', 'b1Text', 'b2Icon', 'b2Heading', 'b2Text']);
  // The room a counted part takes is counted too: a chart with a caption has a lower chart.
  const boxOf = (withParts?: string[]) => {
    const planned = freePattern(template, {}, { blocks: ['chart', 'note'] }, false, withParts);
    assert.ok(typeof planned !== 'string');
    const slot = planned.pattern.slots.find((entry) => entry.name === 'b1Chart');
    return slot?.kind === 'chart' ? (slot.box?.h ?? 0) : 0;
  };
  assert.ok(boxOf(['b1Caption']) < boxOf(), `${boxOf(['b1Caption'])} and ${boxOf()}`);
});

// ---- What a slide is told before it is written -----------------------------------------------

const WORDS = 'river dock bike morning station tram season riders bridge market signal repair member weekend commute harbour lantern ticket meadow copper'.split(' ');
/** Ordinary words, exactly so many characters of them. */
function wordsOf(length: number, seed = 0): string {
  let out = '';
  for (let index = seed; out.length < length; index++) out += (out ? ' ' : '') + (WORDS[(index * 7 + seed) % WORDS.length] as string);
  out = out.slice(0, Math.max(1, length));
  return out.endsWith(' ') ? `${out.slice(0, -1)}x` : out;
}
/** A planned slide with every slot written: its texts to a share of their budgets, in ordinary words. */
function writtenTo(planned: { pattern: TemplateManifest['patterns'][number] }, blocks: FreeBlockKind[], share: number): SlideSpec {
  const slots: SlideSpec['slots'] = { title: 'A title' };
  for (const slot of planned.pattern.slots) {
    if (!/^b\d/.test(slot.name)) continue;
    if (slot.kind === 'text') {
      const length = Math.max(1, Math.round((slot.maxChars ?? 1) * share));
      if (/Items$/.test(slot.name)) {
        const points = Math.min(4, slot.maxParagraphs ?? 1);
        const each = Math.floor(length / points);
        slots[slot.name] = Array.from({ length: points }, (_, index) => wordsOf(index === 0 ? length - each * (points - 1) : each, index));
      } else slots[slot.name] = /Value$/.test(slot.name) ? '87%' : wordsOf(length, slot.name.length);
    } else if (slot.kind === 'chart') slots[slot.name] = chart;
    else if (slot.kind === 'table') slots[slot.name] = { header: ['Station', 'Docks'], rows: [['Tram Square', '24']] };
    else if (slot.kind === 'image') slots[slot.name] = { asset: slot.imageKind === 'icon' ? 'clock' : 'pier' };
  }
  return { pattern: 'free', blocks, slots };
}

test('a planned slide is set out for texts of the usual length, and each slot is told what its box holds then', () => {
  const planned = (blocks: FreeBlockKind[], on = template) => {
    const made = freePattern(on, {}, { blocks }, false);
    assert.ok(typeof made !== 'string');
    return made;
  };
  const told = (kind: FreeBlockKind, part: string, count: number, on = template): number => {
    const slot = planned(Array<FreeBlockKind>(count).fill(kind), on).pattern.slots.find((entry) => entry.name === `b1${part}`);
    return slot?.kind === 'text' ? (slot.maxChars ?? 0) : 0;
  };
  // Cards. A card is planned as high as a text of about 70 characters needs, not as high as its room: three
  // cards are told 78 characters where they were told 195. Nine, twelve and sixteen cards were told 21, 17
  // and 17, at the template's own size, where the slide that is laid out takes 54, 42 and 22 at the size its
  // text is then set at. A heading keeps its size when its text is set smaller.
  assert.deepEqual([1, 2, 3, 4, 6, 9, 12, 16].map((count) => told('card', 'Text', count)), [144, 129, 78, 85, 95, 54, 42, 22]);
  for (const count of [3, 9, 12, 16]) {
    // The slide written with texts of 70 characters, or of its budget where that is less, is that slide: the
    // same sizes, and slots that hold the same. (Text longer than its budget is drawn smaller where PowerPoint
    // would not hold it, and may take the larger room.)
    const blocks = Array<FreeBlockKind>(count).fill('card');
    const before = planned(blocks);
    const told = before.pattern.slots.find((entry) => entry.name === 'b1Text');
    const laid = planOf(freeSlide({ title: 'A title', blocks: Array.from({ length: count }, () => card('A heading', wordsOf(Math.min(70, told?.kind === 'text' ? (told.maxChars ?? 70) : 70)))) }));
    const [budget, text] = [told, laid.slot('b1Text')];
    assert.ok(budget?.kind === 'text' && text.kind === 'text');
    assert.deepEqual([text.fontSize, text.maxChars], [budget.fontSize, budget.maxChars], `${count} cards`);
  }
  // No body text is told less than a dozen characters, however many blocks the slide has: twenty-four
  // quotations, whose boxes hold no line, were told 1. A note is never told more than a note alone.
  for (const [kind, part] of [['card', 'Text'], ['step', 'Text'], ['text', 'Text'], ['list', 'Items'], ['quote', 'Quote'], ['note', 'Text']] as const) {
    const budgets = Array.from({ length: 24 }, (_, index) => told(kind, part, index + 1));
    assert.ok(Math.min(...budgets) >= 12, `${kind}: ${budgets.join(' ')}`);
    if (kind === 'note') assert.ok(Math.max(...budgets) <= (budgets[0] as number), budgets.join(' '));
  }
  assert.deepEqual([told('quote', 'Quote', 24), told('text', 'Text', 24), told('note', 'Text', 1), told('note', 'Text', 24)], [17, 12, 153, 153]);
  // Running text and a quotation alone are told what the slide holds at the size a text of the usual length gets there.
  const alone = planned(['text']).pattern.slots.find((entry) => entry.name === 'b1Text');
  assert.ok(alone?.kind === 'text');
  assert.deepEqual([alone.fontSize, alone.maxChars, alone.description], [22, 584, 'Write 290 to 550 characters']);
  // Text written to its budget, and to nine tenths of it, is never reported on its slot: every kind, from one
  // block to twenty-four, on the demo and on the demo with a line spacing of 130%. The one report that stays is
  // the honest one, for a box that holds no line of its text at all.
  const spaced = loadTemplate(spacedBody('<a:spcPct val="130000"/>'), manifest);
  const kinds: FreeBlockKind[] = ['text', 'list', 'card', 'figure', 'step', 'quote', 'note', 'chart', 'table', 'photo', 'graphic'];
  const mixes: FreeBlockKind[][] = [['chart', 'card', 'card', 'card'], ['figure', 'figure', 'figure', 'chart', 'chart'], ['text', 'list'], ['note', 'card', 'card', 'card'], ['chart', 'chart', 'card', 'card', 'card', 'figure', 'figure', 'text', 'note'], kinds];
  for (const kind of kinds) for (const count of [1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 12, 16, 20, 24]) mixes.push(Array<FreeBlockKind>(count).fill(kind));
  let written = 0;
  for (const on of [template, spaced]) {
    for (const blocks of mixes) {
      const before = planned(blocks, on);
      for (const share of [1, 0.9]) {
        const slide = writtenTo(before, blocks, share);
        const reported = validateDeck(on, deckOf([slide]), { assets: ['clock', 'pier'] }).filter((issue) => issue.slot !== undefined && !/the slot fits about 1$/.test(issue.message));
        assert.deepEqual(reported, [], `${blocks.length} of ${[...new Set(blocks)].join(', ')} at ${share}`);
        written += Object.keys(slide.slots).length;
      }
    }
  }
  assert.ok(written > 10000);
});

// ---- A wish the room cannot hold -------------------------------------------------------------

test('a wish for more blocks side by side than the room holds is relaxed to the nearest number that fits, and the deck is told', () => {
  const cards = (count: number) => Array<FreeBlockKind>(count).fill('card');
  // Twelve cards on one line would each be half as wide as a card sits well in: five fit, so there are five on a line.
  const twelve = arrangeBlocks(cards(12), { frame, rtl: false, arrange: { columns: 12 } });
  assert.deepEqual([twelve.relaxed, twelve.columns], [{ asked: 12, columns: 5 }, 5]);
  assert.ok(twelve.blocks.every((block) => block.box.w >= 0.15), 'every card has the least width a card sits well in');
  assertTidy(twelve, 12, 'twelve cards, twelve asked for on a line');
  // A wish that fits is followed, and says nothing: five cards on a line, and a number larger than the blocks there are.
  for (const columns of [5, 9]) assert.equal(arrangeBlocks(cards(5), { frame, rtl: false, arrange: { columns } }).relaxed, undefined);
  assert.equal(arrangeBlocks(cards(6), { frame, rtl: false, arrange: { columns: 6 } }).relaxed?.columns, 5);
  // Too few on a line is a wish the room cannot hold either. Blocks may be low, down to half the least height
  // their kind sits well in: seven cards one under the other are, eight are not, and stand in two columns.
  assert.equal(arrangeBlocks(cards(7), { frame, rtl: false, arrange: { columns: 1 } }).relaxed, undefined);
  const eight = arrangeBlocks(cards(8), { frame, rtl: false, arrange: { columns: 1 } });
  assert.deepEqual([eight.relaxed, eight.columns], [{ asked: 1, columns: 2 }, 2]);
  assertTidy(eight, 8, 'eight cards, one asked for on a line');
  // The nearest number is taken, fewer before more: fourteen texts asked for in six columns stand in four.
  assert.deepEqual(arrangeBlocks(Array<FreeBlockKind>(14).fill('text'), { frame, rtl: false, arrange: { columns: 6 } }).relaxed, { asked: 6, columns: 4 });
  // A lead stays where it was asked for, and the blocks beside it are relaxed among themselves.
  const beside = arrangeBlocks(['chart', ...cards(8)], { frame, rtl: false, arrange: { columns: 1, lead: 'left' } });
  assert.deepEqual([beside.relaxed, beside.blocks[0]?.lead, beside.blocks[0]?.box.x], [{ asked: 1, columns: 2 }, true, frame.x]);
  // When no number of columns helps, the slide holds too many blocks: it is made as asked, as before.
  for (const [kinds, columns] of [[Array<FreeBlockKind>(400).fill('text'), 400], [Array<FreeBlockKind>(300).fill('chart'), 1], [cards(200), 3]] as const) {
    const crowded = arrangeBlocks(kinds, { frame, rtl: false, arrange: { columns } });
    assert.deepEqual([crowded.relaxed, crowded.columns, crowded.dense], [undefined, Math.min(columns, kinds.length), true], `${kinds.length} blocks, ${columns} on a line`);
  }
  // The deck is told in a note of its own, without a slide number: no rewriting of text changes how a slide is set out.
  const deck = (count: number, columns: number) => deckOf([freeSlide({ title: 'Ways in', blocks: Array.from({ length: count }, (_, index) => card(`Point ${index + 1}`, 'Bikes unlock from five.')), columns })]);
  assert.deepEqual(withoutCodes(validateDeck(template, deck(12, 12))), [{ severity: 'warning', message: 'slide 1: 12 side by side were asked for; the slide holds 5' }]);
  assert.deepEqual(withoutCodes(validateDeck(template, deck(8, 1))), [{ severity: 'warning', message: 'slide 1: 1 side by side was asked for; the slide holds 2' }]);
  assert.deepEqual(validateDeck(template, deck(5, 5)), []);
  // The slide is rendered as it was checked.
  const xml = slideXml(renderDeck(template, deck(12, 12)).pptx, 1);
  const tops = new Set([...xml.matchAll(/name="Block \d+ panel"\/>.*?<a:off x="\d+" y="(\d+)"\/>/g)].map((match) => match[1]));
  assert.equal(tops.size, 3, 'twelve cards on three lines of five, five and two');
});

/** Whether two boxes share more than a hair of the slide. */
const overlap = (a: { x: number; y: number; w: number; h: number }, b: { x: number; y: number; w: number; h: number }) => !(a.x + a.w <= b.x + 0.0008 || b.x + b.w <= a.x + 0.0008 || a.y + a.h <= b.y + 0.0008 || b.y + b.h <= a.y + 0.0008);

test('no shape of a block lies outside its block, however crowded the slide or whatever was asked for', () => {
  const kinds: FreeBlockKind[] = ['text', 'list', 'card', 'figure', 'step', 'quote', 'note', 'chart', 'table', 'photo', 'graphic'];
  const content = (kind: FreeBlockKind, index: number, full: boolean): FreeBlockContent => {
    const heading = full ? { heading: wordsOf(14, index) } : {};
    const caption = full ? { caption: wordsOf(16, index) } : {};
    switch (kind) {
      case 'text':
        return { text: wordsOf(180, index), ...heading };
      case 'list':
        return { list: [wordsOf(40, index), wordsOf(40, index + 1), wordsOf(40, index + 2)], ...heading };
      case 'card':
        return { card: { heading: wordsOf(12, index), text: wordsOf(70, index), ...(full ? { icon: 'clock' } : {}) } };
      case 'figure':
        return { figure: { value: '87%', label: wordsOf(26, index) } };
      case 'step':
        return { step: { heading: wordsOf(12, index), text: wordsOf(70, index) } };
      case 'quote':
        return { quote: wordsOf(110, index), ...(full ? { source: wordsOf(22, index) } : {}) };
      case 'note':
        return { note: wordsOf(80, index) };
      case 'chart':
        return { chart, ...caption };
      case 'table':
        return { table: { header: ['Station', 'Docks'], rows: [['Tram Square', '24'], ['River Gate', '18']] }, ...caption };
      case 'photo':
        return { photo: 'pier', ...caption };
      default:
        return { graphic: 'pier', ...caption };
    }
  };
  const slides: Array<[string, SlideSpec, Partial<DeckSpec>]> = [];
  // Slides of one kind, far past what a slide holds, with every optional part and at larger text.
  for (const kind of kinds) {
    for (const [count, textScale] of [[30, 1], [60, 1.5], [200, 1], [400, 1.25]] as const) {
      slides.push([`${count} ${kind}s at ${textScale}`, freeSlide({ title: 'A title', blocks: Array.from({ length: count }, (_, index) => content(kind, index, true)) }), { textScale }]);
    }
    for (const columns of [1, 2, 7]) slides.push([`20 ${kind}s, ${columns} on a line`, freeSlide({ blocks: Array.from({ length: 20 }, (_, index) => content(kind, index, false)), columns }), {}]);
  }
  // Mixes, with and without wishes. A fixed sequence, so that a failure can be found again.
  let seed = 11;
  const next = () => (seed = (seed * 1103515245 + 12345) % 2147483648) / 2147483648;
  for (let round = 0; round < 240; round++) {
    const count = 1 + Math.floor(next() * (round % 5 === 0 ? 120 : 14));
    const blocks = Array.from({ length: count }, (_, index) => content(kinds[Math.floor(next() * kinds.length)] as FreeBlockKind, index, next() < 0.5));
    const more: { columns?: number; lead?: 'left' | 'right' | 'top' | 'none' } = {};
    if (next() < 0.4) more.columns = 1 + Math.floor(next() * 6);
    if (next() < 0.3) more.lead = (['left', 'right', 'top', 'none'] as const)[Math.floor(next() * 4)] as 'left';
    slides.push([`mix ${round}: ${blocks.map((block) => Object.keys(block)[0]).join(',')} ${JSON.stringify(more)}`, freeSlide({ ...(next() < 0.8 ? { title: 'A title' } : {}), blocks, ...more }), next() < 0.3 ? { textScale: 1.25 } : {}]);
  }
  let shapes = 0;
  for (const [label, slide, more] of slides) {
    for (const direction of ['ltr', 'rtl'] as const) {
      const resolved = freeFor(template, { ...more, direction }, slide);
      assert.ok(typeof resolved !== 'string', label);
      const placeOf = new Map(resolved.plan.layout.blocks.map((block) => [block.index, block.box]));
      const solid = resolved.plan.shapes.filter((shape) => shape.kind !== 'rule');
      const blockOf = (name: string) => Number(/^Block (\d+) /.exec(name)?.[1] ?? 0);
      for (const shape of solid) {
        const block = placeOf.get(blockOf(shape.name));
        assert.ok(block, `${label}: ${shape.name}`);
        const { x, y, w, h } = shape.box;
        assert.ok([x, y, w, h].every(Number.isFinite) && w >= 0 && h >= 0, `${label}: ${shape.name} ${JSON.stringify(shape.box)}`);
        assert.ok(x >= block.x - 0.0008 && x + w <= block.x + block.w + 0.0008 && y >= block.y - 0.0008 && y + h <= block.y + block.h + 0.0008, `${label} (${direction}): ${shape.name} ${JSON.stringify(shape.box)} is outside its block ${JSON.stringify(block)}`);
        shapes++;
      }
      // So no shape of one block lies on a shape of another: the blocks themselves stand apart.
      const laid = resolved.plan.layout.blocks;
      for (let a = 0; a < laid.length; a++) {
        const first = laid[a] as (typeof laid)[number];
        for (let b = a + 1; b < Math.min(laid.length, a + 40); b++) assert.ok(!overlap(first.box, (laid[b] as (typeof laid)[number]).box), `${label} (${direction}): blocks ${first.index} and ${laid[b]?.index} overlap`);
      }
    }
  }
  assert.ok(shapes > 50000, String(shapes));
  // A slide whose blocks have their sizes is not touched by this: its shapes are where they were planned.
  const three = planOf(threeCards);
  assert.deepEqual(three.shape('Block 1 text').box, three.slot('b1Text').kind === 'text' ? (three.slot('b1Text') as FreeTextSlot).box : undefined);
  // A caption and a chart share a block too low for both half and half, and a text whose box holds no line says so.
  const low = planOf(freeSlide({ title: 'A title', blocks: Array.from({ length: 400 }, (_, index) => ({ chart, caption: `Station ${index + 1}` })) }));
  const [caption, drawn, block] = [low.shape('Block 1 caption').box, low.shape('Block 1 chart').box, low.plan.layout.blocks[0]?.box];
  assert.ok(block && caption.y + caption.h <= drawn.y + 0.0002 && drawn.y + drawn.h <= block.y + block.h + 0.0002 && drawn.h >= 0.4 * block.h, JSON.stringify([caption, drawn, block]));
});

test('blocks that are lower than their room sit down it by degrees: no height at which a slide jumps', () => {
  // Low blocks sit two fifths of what they leave down their room. That share falls off evenly as they
  // leave less, from three tenths of the room left over down to one tenth, where it used to end at once
  // when they took three fifths of the room: four cards 60.4% of the room high stood pressed under the
  // title over an empty 40%, and cards 59% high 16% of the room lower.
  assert.equal(settle(0.2, 1), 0.4 * 0.8);
  assert.ok(Math.abs(settle(0.6, 1) - 0.4 * 0.4) < 1e-12);
  assert.deepEqual([settle(0.9, 1), settle(0.95, 1), settle(1, 1), settle(1.2, 1), settle(0.5, 0)], [0, 0, 0, 0, 0]);
  let before = settle(0, 1);
  for (let used = 0.005; used <= 1.0001; used += 0.005) {
    const now = settle(used, 1);
    assert.ok(now <= before + 1e-12 && before - now < 0.006, `at ${used.toFixed(3)}: ${before.toFixed(4)} then ${now.toFixed(4)}`);
    // The blocks never end lower for being higher, and never leave their room.
    assert.ok(used + now <= 1 + 1e-9 && used + now >= used - 0.005 + before - 1e-9, `at ${used.toFixed(3)}`);
    before = now;
  }
  // Three cards, a little higher each time: their top moves up by degrees. Lengths in slide widths. The
  // cards first grow into what is left over, by up to 35%, so less is left to share above and below.
  const tops = [0.19, 0.195, 0.2, 0.205, 0.21, 0.25, 0.3].map((high) => arrangeBlocks(['card', 'card', 'card'], { frame, rtl: false, needed: () => high }).blocks[0]?.box.y ?? 0);
  assert.deepEqual(tops, [0.3282, 0.3209, 0.3145, 0.3091, 0.3047, 0.3004, 0.3]);
  tops.slice(0, 4).forEach((top, index) => assert.ok(top - (tops[index + 1] as number) < 0.01, tops.join(', ')));
  // What holds running text or a table starts at the top of the room, as before, and so does a slide that fills it.
  assert.equal(arrangeBlocks(['text', 'text'], { frame, rtl: false }).blocks[0]?.box.y, frame.top);
  assert.equal(arrangeBlocks(['chart', 'chart'], { frame, rtl: false }).blocks[0]?.box.y, frame.top);
});

// ---- Larger text, small text, lists --------------------------------------------------------------

test('a deck that asks for larger text keeps it where it fits no longer, and the renderer never shrinks a text under its size without the wish', () => {
  // A chart beside three cards. At the deck's own size the card text is 14.5 pt and fits.
  const two: ChartValue = { type: 'column', categories: ['Spring', 'Summer', 'Autumn', 'Winter'], series: [{ name: '2024', values: [10, 14, 12, 18] }, { name: '2025', values: [8, 11, 15, 13] }] };
  const slide = freeSlide({
    title: 'A chart and three cards',
    blocks: [
      { chart: two },
      card('Longer hours', 'Bikes unlock from five in the morning, for early shifts and trams.'),
      card('Safer docks', 'Every station gets lighting and a camera before the season opens.'),
      card('Group passes', 'One pass for up to five riders, bought once and shared by link.'),
    ],
  });
  const sizeOf = (plan: ReturnType<typeof planOf>) => plan.shape('Block 2 text').size ?? 0;
  assert.equal(sizeOf(planOf(slide)), 14.5);
  assert.deepEqual(validateDeck(template, deckOf([slide])), []);
  // With larger text the slide keeps it: no text is smaller than without the wish, so nothing is said of the
  // text size. The card text is then too long for its box, and that is reported on its slot, like text over
  // any budget. The larger text is not bought with smaller text; it can be bought with text that fits no longer.
  // Such text is drawn at the size at which it stays in its card (17 pt at 1.25, 15 at 1.5, under the larger
  // heading), never under its size without the wish.
  for (const [textScale, size] of [[1.25, 17], [1.5, 15]] as const) {
    const larger = planOf(slide, { textScale });
    assert.deepEqual([larger.enlarged, sizeOf(larger)], [undefined, size], `at ${textScale}`);
    const issues = validateDeck(template, deckOf([slide], { textScale }));
    assert.deepEqual(issues.filter((issue) => /larger text/.test(issue.message)), [], `at ${textScale}`);
    assert.deepEqual(issues.map((issue) => issue.slot), ['b2Text', 'b3Text', 'b4Text'], `at ${textScale}`);
    assert.match(issues[0]?.message ?? '', /^text is 66 characters; the slot fits about \d+$/);
  }
  // A template that lets the renderer set text smaller to fit (`shrinkText`), loaded with fonts. The renderer
  // used to take the card text down to 14 pt at 1.25 and 1.5: under the 14.5 pt it has without the wish. The
  // plan now draws it at a size its card holds, never under 14.5 pt, and the renderer has nothing to set smaller.
  const fonts = fontkitMetrics(fontkit, openFonts());
  const shrinking = loadTemplate(demoPptx, { ...manifest, shrinkText: { to: 0.8 } }, { fonts });
  const drawn = (textScale: number) => {
    const deck = deckOf([slide], textScale === 1 ? {} : { textScale });
    const { pptx, shrunk } = renderDeck(shrinking, deck);
    const xml = slideXml(pptx, 1);
    const sizes = [2, 3, 4].map((block) => Number(new RegExp(`name="Block ${block} text"/>.*?<a:rPr[^>]*\\bsz="(\\d+)"`).exec(xml)?.[1]) / 100);
    return { sizes, shrunk, issues: validateDeck(shrinking, deck) };
  };
  const plain = drawn(1);
  assert.deepEqual([plain.sizes, plain.shrunk], [[14.5, 14.5, 14.5], []]);
  for (const textScale of [1.25, 1.5]) {
    const larger = drawn(textScale);
    // The plan has drawn the text at a size its card holds, so the renderer sets nothing smaller.
    assert.deepEqual(larger.sizes, textScale === 1.25 ? [17, 17, 17] : [15, 15, 15], `at ${textScale}`);
    assert.deepEqual(larger.shrunk, [], `at ${textScale}`);
    // The heading keeps the larger size the deck asked for.
    assert.ok(Number(/name="Block 2 heading"\/>.*?<a:rPr[^>]*\bsz="(\d+)"/.exec(slideXml(renderDeck(shrinking, deckOf([slide], { textScale })).pptx, 1))?.[1]) / 100 > 19, `at ${textScale}`);
  }
  // The floor is the size the text is drawn at without the wish: a text that is set smaller there too may be set as small here.
  const long = freeSlide({ title: 'Three ways in', blocks: [card('Longer hours', wordsOf(200)), card('Safer docks', wordsOf(200, 1)), card('Group passes', wordsOf(200, 2))] });
  const sizesOf = (textScale: number) => {
    const xml = slideXml(renderDeck(shrinking, deckOf([long], textScale === 1 ? {} : { textScale })).pptx, 1);
    return [1, 2, 3].map((block) => Number(new RegExp(`name="Block ${block} text"/>.*?<a:rPr[^>]*\\bsz="(\\d+)"`).exec(xml)?.[1]) / 100);
  };
  const at1 = sizesOf(1);
  for (const textScale of [1.25, 1.5]) sizesOf(textScale).forEach((size, index) => assert.ok(size >= (at1[index] as number), `at ${textScale}: ${size} pt, and ${at1[index]} pt at 1`));
  // A slide of a deck without the wish, and a slide of a layout, are shrunk as the manifest says, as before.
  assert.equal(freeTextFloor(undefined, { name: 'body', kind: 'text', shapeId: '3' }, ['x'], shrinking.manifest, {}, textMeasurer(shrinking)), undefined);
  const resolved = freeFor(shrinking, {}, slide);
  assert.ok(typeof resolved !== 'string');
  assert.equal(resolved.plain, undefined);
});

test('a slide whose smallest text is under 9 pt is told so', () => {
  // On a slide 13.33 inches wide. Under that nobody reads a projected slide; the note used to start under 7 pt.
  const words = 'Rides grew at every station this season, and the docks next to the tram stops emptied first on every weekday morning of the year, in rain and sun.';
  const notes = (count: number) => freeSlide({ title: 'Rides', blocks: [{ chart }, ...Array.from({ length: count }, (_, index) => ({ note: index % 2 ? words : words.slice(0, 55) }))] });
  const six = planOf(notes(6));
  assert.deepEqual([six.plan.smallest, six.plan.comfortable], [8, 9]);
  assert.deepEqual(withoutCodes(validateDeck(template, deckOf([notes(6)]))), [{ severity: 'warning', message: 'slide 1 holds 7 blocks, as asked; its smallest text is set at 8 pt' }]);
  // Text at 9 pt and over says nothing: three notes under a chart are set at 12.5 pt.
  assert.ok(planOf(notes(3)).plan.smallest >= 9);
  assert.deepEqual(validateDeck(template, deckOf([notes(3)])), []);
  // The size follows the slide's width, as the other sizes do: a slide 10 inches wide is told under 7 pt.
  const narrow = freeFor(loadTemplate(buildFixturePptx(), fixtureManifest), {}, freeSlide({ title: 'Three things', blocks: [{ text: 'A short text.' }] }));
  assert.ok(typeof narrow !== 'string');
  assert.equal(narrow.plan.comfortable, Math.round(Math.max(5, (9 * narrow.kit.slideWidth) / 12700 / 960) * 2) / 2);
  assert.ok(narrow.plan.comfortable < 9);
});

test('a list is fitted point by point: each point in whole lines behind its bullet, with the space between two of them', () => {
  // Six points of 60 characters beside a text. A point takes two lines in the width its bullet leaves, so the
  // list takes twelve lines and five gaps, where its 360 characters alone come to nine lines. At 18 pt
  // PowerPoint set it 26 pt higher than its box, and nothing was said. It is set at 16.5 pt with the
  // usual space of 8 between its points, or, the space between them giving way first (0.6.9), at
  // 17 pt with 7.
  const beside = (points: number) => freeSlide({ title: 'A list of six', blocks: [{ list: Array.from({ length: points }, (_, index) => wordsOf(60, index)), heading: wordsOf(12) }, { text: wordsOf(260), heading: wordsOf(12, 3) }] });
  const six = planOf(beside(6));
  const list = six.shape('Block 1 list');
  const slot = six.slot('b1Items') as FreeTextSlot;
  assert.deepEqual([list.size, slot.paragraphGap], [17, 7]);
  // Twelve lines of 1.2 times the size and five gaps end inside the box, its last line within the
  // part of a line that may hang under it (a quarter of the size, at most 3 pt).
  assert.ok(12 * 1.2 * 17 + 5 * 7 <= list.box.h * H + Math.min(3, 0.25 * 17) + 0.1, `a box ${(list.box.h * H).toFixed(1)} pt high`);
  assert.ok(12 * 1.2 * 18 + 5 * 7 > list.box.h * H + 4.5, 'at 18 pt it would not, even with the gaps closed up');
  // A line of the list is counted in the width its bullet leaves.
  assert.equal(slot.drawn.hang, 1.1 * 17);
  assert.deepEqual(validateDeck(template, deckOf([beside(6)])), []);
  // Measured with the template's fonts it fits too: the count and the measure agree.
  const measured = loadTemplate(demoPptx, manifest, { fonts: fontkitMetrics(fontkit, openFonts()) });
  const fit = textMeasurer(measured).fit(planOf(beside(6), {}, measured).plan.pattern, planOf(beside(6), {}, measured).slot('b1Items') as FreeTextSlot, beside(6).slots.b1Items as string[]);
  assert.ok(fit?.fits && fit.lines >= 10 && fit.lines <= 12, JSON.stringify([fit?.fits, fit?.lines]));
  // The text beside it is not set smaller for the list's sake: unlike texts do not share a size.
  assert.equal(six.shape('Block 2 text').size, 18);
  // Eight such points do not fit at four fifths of the size either: that is reported on the slot, by the lines the points take.
  const eight = validateDeck(template, deckOf([beside(8)]));
  assert.equal(eight.length, 1);
  assert.match(`${eight[0]?.slot}: ${eight[0]?.message}`, /^b1Items: text is 480 characters; the slot fits about \d+$/);
  assert.ok(Number(/about (\d+)$/.exec(eight[0]?.message ?? '')?.[1]) < 480);
  // A list that fits holds its points: four short points are not told to be three because the text was set larger.
  const short = freeSlide({ title: 'Lists', blocks: Array.from({ length: 5 }, () => ({ list: ['Open earlier each day', 'Fix flats in one day', 'Add forty more bikes', 'Train the volunteers'] })) });
  assert.deepEqual(validateDeck(template, deckOf([short])), []);
  assert.ok(((planOf(short).slot('b1Items') as FreeTextSlot).maxParagraphs ?? 0) >= 4);
});

// ---- What is the package's interface, and what is not ------------------------------------------

test('what the engine keeps to itself is not exported by the package, and a frame can be asked why it is refused', async () => {
  const api: Record<string, unknown> = await import('../src/index.ts');
  // Working parts of the engine: they change with it, and nothing outside it is to lean on them.
  for (const name of ['FREE_FRAME_ROOM', 'isTextValue', 'plainText', 'freeKit', 'freeKitOrReason', 'freeFor', 'freePlan', 'arrangeBlocks']) assert.equal(name in api, false, name);
  // Why a layout cannot be the frame of free slides, as a sentence, or undefined when it can.
  assert.equal(freeBaseRefusal(template, 'list-with-photo'), undefined);
  assert.equal(freeBaseRefusal(template, 'text-only'), undefined);
  assert.equal(freeBaseRefusal(template, 'gone'), 'the template has no layout "gone"');
  assert.equal(freeBaseRefusal(template, 'two-columns'), 'the layout "two-columns" has a row of columns, which the software lays out itself');
  assert.equal(freeBaseRefusal(template, 'cover'), 'the layout "cover" has its title low on the slide: blocks would get 22% of the slide\'s height, and a frame leaves them 45% or more');
});

// ---- How a free slide looks ----------------------------------------------------------------------

/** What the value axis of a chart part says: where it ends, the step between its lines, whether it is shown. */
const axisOf = (xml: string) => ({ max: /<c:max val="([^"]+)"\/>/.exec(xml)?.[1], step: /<c:majorUnit val="([^"]+)"\/>/.exec(xml)?.[1], shown: !axesGone(xml), numbers: /<c:showVal val="1"\/>/.test(xml) });
const quarters = (values: number[], more: Partial<ChartValue> = {}): ChartValue => ({ type: 'column', categories: ['Q1', 'Q2', 'Q3', 'Q4'], series: [{ name: 'Rides', values }], ...more });
const chartSlide = (charts: ChartValue[], more: FreeBlockContent[] = [], title = 'Stations') => freeSlide({ title, blocks: [...charts.map((value, index): FreeBlockContent => ({ chart: value, caption: `Station ${index + 1}` })), ...more] });
const detailsOf = (slide: SlideSpec, on = template) => planOf(slide, {}, on).plan.pattern.slots.flatMap((slot) => (slot.kind === 'chart' ? [slot.detail ?? 'full'] : []));

test('like charts of a slide share one value axis, and a bare chart uses the height of its frame', () => {
  const stations = [[12, 18, 23, 14], [15, 22, 29, 18], [40, 61, 75, 70], [7, 13, 17, 9], [11, 19, 24, 13], [14, 20, 26, 16]];
  // Six small charts whose largest values are 17 to 75: one scale, so that a bar of one can be read against a
  // bar of the next. Their numbers stand at the bars, so no axis is drawn, and the longest bar of them all ends
  // where its plot ends.
  const six = chartsOf(renderDeck(template, deckOf([chartSlide(stations.map((values) => quarters(values)))])).pptx, 1).map(axisOf);
  assert.deepEqual([...new Set(six.map((axis) => JSON.stringify(axis)))], [JSON.stringify({ max: '75', step: undefined, shown: false, numbers: true })]);
  // Two full charts: lines at round steps, the same in both, and an end at the half step over the largest value
  // of the two (45 for 44, not 50), so that the longest bar uses as much of its plot as it can.
  const two = chartsOf(renderDeck(template, deckOf([chartSlide([quarters([33, 38, 38, 36]), quarters([38, 41, 39, 44])])])).pptx, 1).map(axisOf);
  assert.deepEqual(two, [{ max: '45', step: '10', shown: true, numbers: true }, { max: '45', step: '10', shown: true, numbers: true }]);
  // More than five times apart nothing is forced: each chart has the round end its own values ask for.
  const apart = chartsOf(renderDeck(template, deckOf([chartSlide([quarters([33, 38, 38, 36]), quarters([214, 238, 262, 291])])])).pptx, 1).map(axisOf);
  assert.deepEqual(apart.map((axis) => axis.max), ['40', '300']);
  // Charts of another type, another format or another number of series are no group: each has the round axis
  // its own values ask for. With a line among them no chart has numbers at its bars, so every one shows its
  // axis. A chart with a value under nothing starts at the whole step under its lowest value; a pie has no axis.
  const mixed = [quarters([10, 20, 30, 40]), quarters([40, 50, 60, 70], { type: 'bar' }), quarters([5, 6, 7, 8], { format: 'percent' }), quarters([5, -6, 7, 8]), quarters([1, 2, 3, 4], { type: 'line' }), quarters([1, 2, 3, 4], { type: 'pie' })];
  const mixedCharts = chartsOf(renderDeck(template, deckOf([chartSlide(mixed)])).pptx, 1);
  assert.deepEqual(mixedCharts.map((xml) => axisOf(xml).max), ['40', '70', '8', '10', '4', undefined]);
  assert.deepEqual(mixedCharts.map((xml) => /<c:min val="(-?[\d.]+)"\/>/.exec(xml)?.[1]), ['0', '0', '0', '-10', '0', undefined]);
  assert.ok(mixedCharts.slice(0, 5).every((xml) => axisOf(xml).shown && !axisOf(xml).numbers));
  // Line charts that are a group share the scale too, from nothing up.
  const lines = chartsOf(renderDeck(template, deckOf([chartSlide([quarters([10, 20, 30, 40], { type: 'line' }), quarters([40, 50, 60, 70], { type: 'line' })])])).pptx, 1);
  assert.deepEqual(lines.map((xml) => [axisOf(xml).max, /<c:min val="0"\/>/.test(xml)]), [['70', true], ['70', true]]);
  // Thirty bare charts: every one ends at the largest value among them, so the longest bar of the slide touches the top of its frame.
  const thirty = chartsOf(renderDeck(template, deckOf([freeSlide({ title: 'Stations', blocks: Array.from({ length: 30 }, (_, index) => ({ chart: quarters([10 + index, 12, 14, 9]) })) })])).pptx, 1).map(axisOf);
  assert.deepEqual([...new Set(thirty.map((axis) => `${axis.max} ${axis.shown} ${axis.numbers}`))], ['39 false false']);
});

test('a small chart keeps the names of its categories and shows its numbers or its value axis; charts are bare only among many or where nothing reads', () => {
  const eight = chartSlide(Array.from({ length: 8 }, (_, index) => quarters([12 + index, 18, 23, 14])), [{ note: 'Every station peaked in the third quarter.' }]);
  // Eight charts in three rows over a bottom line were bare bars: they are small charts now, with names and numbers.
  assert.deepEqual([...new Set(detailsOf({ ...eight, arrange: { columns: 3 } }))], ['compact']);
  const parts = chartsOf(renderDeck(template, deckOf([{ ...eight, arrange: { columns: 3 } }])).pptx, 1);
  for (const xml of parts) {
    assert.match(xml, /<c:catAx>(?:(?!<\/c:catAx>).)*?<c:delete val="0"\/>/);
    assert.ok(axisOf(xml).numbers && !axisOf(xml).shown);
    assert.match(xml, /<a:bodyPr rot="0" vert="horz"\/>/);
  }
  // The deck is no longer told that they are drawn bare.
  assert.deepEqual(validateDeck(template, deckOf([{ ...eight, arrange: { columns: 3 } }])).filter((issue) => /drawn small/.test(issue.message)), []);
  // Numbers that have no room at the bars are left to a value axis of few steps: 0, 1,000 and 2,000.
  const year: ChartValue = { type: 'column', categories: ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'], series: [{ name: 'Rides', values: [1210, 1340, 1500, 1620, 1710, 1800, 1930, 1990, 1820, 1640, 1400, 1250] }] };
  const crowded = chartsOf(renderDeck(template, deckOf([chartSlide(Array.from({ length: 6 }, () => year))])).pptx, 1).map(axisOf);
  assert.deepEqual([...new Set(crowded.map((axis) => JSON.stringify(axis)))], [JSON.stringify({ max: '2000', step: '1000', shown: true, numbers: false })]);
  // Charts keep their names while their text is 7 pt or more; where it would be smaller they are drawn bare,
  // and from twenty-five on they are a texture, drawn bare whatever their size.
  const many = (count: number) => freeSlide({ title: 'Stations', blocks: Array.from({ length: count }, () => ({ chart: quarters([12, 18, 23, 14]) })) });
  assert.deepEqual([[...new Set(detailsOf(many(10)))], [...new Set(detailsOf(many(24)))], [...new Set(detailsOf(many(25)))]], [['compact'], ['spark'], ['spark']]);
  assert.ok(planOf(many(10)).plan.pattern.slots.every((slot) => slot.kind !== 'chart' || (slot.text?.size ?? 0) >= 7));
  // The same before there is content: what a slot says of its chart does not change when the chart is written.
  for (const count of [1, 3, 8, 12, 24, 25]) {
    const planned = freePattern(template, {}, { blocks: Array<FreeBlockKind>(count).fill('chart') }, false);
    assert.ok(typeof planned !== 'string');
    assert.deepEqual(planned.pattern.slots.flatMap((slot) => (slot.kind === 'chart' ? [slot.detail ?? 'full'] : [])), detailsOf(many(count)), `${count} charts`);
  }
  // On a tile among other blocks a low chart is a picture of its numbers: it stays bare.
  const tiles = freeSlide({ title: 'Mixed', blocks: [{ chart: quarters([1, 2, 3, 4]), caption: 'A' }, { chart: quarters([1, 2, 3, 4]), caption: 'B' }, card('One', 'A line.'), card('Two', 'A line.'), card('Three', 'A line.'), { figure: { value: '48k', label: 'rides' } }, { figure: { value: '87%', label: 'riders' } }, { text: 'A sentence.' }, { note: 'The bottom line.' }] });
  assert.equal(planOf(tiles).plan.layout.tiles, true);
  assert.deepEqual(detailsOf(tiles), ['spark', 'spark']);
});

test('a slide of like charts that show the same series has one legend, and none in any chart', () => {
  const pair = (values: number[], other: number[], more: Partial<ChartValue> = {}): ChartValue => ({ type: 'column', categories: ['Q1', 'Q2', 'Q3', 'Q4'], series: [{ name: '2024', values }, { name: '2025', values: other }], ...more });
  const four = chartSlide([pair([10, 14, 12, 18], [8, 11, 15, 13]), pair([12, 15, 11, 19], [9, 12, 16, 14]), pair([7, 9, 8, 12], [6, 8, 10, 9]), pair([14, 18, 16, 22], [11, 13, 19, 17], { type: 'bar' })]);
  const plan = planOf(four);
  const legend = plan.plan.shapes.filter((shape) => /^Legend/.test(shape.name));
  assert.deepEqual(legend.map((shape) => [shape.name, shape.kind, shape.kind === 'key' ? shape.color : shape.text]), [['Legend key 1', 'key', '0E7C7B'], ['Legend 1', 'text', '2024'], ['Legend key 2', 'key', 'F4A259'], ['Legend 2', 'text', '2025']]);
  // It stands on a strip of its own at the top of the room, from the edge the title starts at; the charts start under it.
  const size = legend[1]?.size ?? 0;
  const top = Math.min(...plan.plan.layout.blocks.map((block) => block.box.y));
  assert.ok(Math.abs(top - (plan.kit.band.top + (2 * size) / H)) < 0.0002, `the charts start at ${top}`);
  assert.equal(legend[0]?.box.x, plan.kit.area.x);
  assert.ok((legend[1]?.box.y ?? 0) >= plan.kit.band.top - 0.0001 && (legend[1]?.box.y ?? 0) + (legend[1]?.box.h ?? 0) <= top + 0.0001);
  for (const slot of plan.plan.pattern.slots) if (slot.kind === 'chart') assert.equal(slot.drawn?.legend, false);
  const { pptx } = renderDeck(template, deckOf([four]));
  for (const xml of chartsOf(pptx, 1)) assert.doesNotMatch(xml, /<c:legend>/);
  assert.deepEqual(textsOf(slideXml(pptx, 1)).filter((text) => /^20\d\d$/.test(text)), ['2024', '2025']);
  assert.match(slideXml(pptx, 1), /name="Legend key 1"\/>.*?<a:solidFill><a:srgbClr val="0E7C7B"\/><\/a:solidFill><a:ln><a:noFill\/><\/a:ln>/);
  // Read from the right, the legend starts at the right edge of the room.
  const mirrored = planOf(four, { direction: 'rtl' }).plan.shapes.find((shape) => shape.name === 'Legend key 1');
  assert.ok(Math.abs((mirrored?.box.x ?? 0) + (mirrored?.box.w ?? 0) - (plan.kit.area.x + plan.kit.area.w)) < 0.0002);
  // Charts that do not all show the same series, a pie among them, or other blocks on the slide: each chart keeps its own legend.
  const own = (slide: SlideSpec) => {
    const made = planOf(slide);
    return [made.plan.shapes.some((shape) => /^Legend/.test(shape.name)), chartsOf(renderDeck(template, deckOf([slide])).pptx, 1).filter((xml) => /<c:legend>/.test(xml)).length];
  };
  assert.deepEqual(own(chartSlide([pair([1, 2, 3, 4], [2, 3, 4, 5]), { ...pair([1, 2, 3, 4], [2, 3, 4, 5]), series: [{ name: '2023', values: [1, 2, 3, 4] }, { name: '2025', values: [2, 3, 4, 5] }] }])), [false, 2]);
  assert.deepEqual(own(chartSlide([pair([1, 2, 3, 4], [2, 3, 4, 5]), pair([1, 2, 3, 4], [2, 3, 4, 5]), quarters([40, 30, 20, 10], { type: 'pie' })])), [false, 3]);
  assert.deepEqual(own(chartSlide([pair([1, 2, 3, 4], [2, 3, 4, 5]), pair([1, 2, 3, 4], [2, 3, 4, 5])], [card('Longer hours', 'Bikes unlock from five.')])), [false, 2]);
  assert.deepEqual(own(chartSlide([pair([1, 2, 3, 4], [2, 3, 4, 5]), pair([1, 2, 3, 4], [2, 3, 4, 5])], [{ note: 'Both grew.' }])), [true, 0]);
  // Before there is content no strip is kept.
  assert.ok(typeof freePattern(template, {}, { blocks: ['chart', 'chart'] }, false) !== 'string');
});

test('a figure is as large as its card and the number of figures allow, and sits with its label in the middle of the card', () => {
  const labels = ['active users of the app', 'customers who would recommend us', 'minutes to answer, on average', 'branches across the country', 'operations done in digital channels', 'new customers this year'];
  const figures = (count: number) => freeSlide({ title: 'The year in numbers', blocks: Array.from({ length: count }, (_, index): FreeBlockContent => ({ figure: { value: ['350K', '87%', '2.5', '112', '78%', '24K'][index % 6] as string, label: labels[index % 6] as string } })) });
  // The fewer figures, the larger: one has up to two and a half times the title, six have the template's own 40 pt.
  // Three on one line are three quarters as high as wide, so that they are not set smaller than four in two lines.
  const sizes = [1, 2, 3, 4, 6, 12].map((count) => [planOf(figures(count)).shape('Block 1 value').size, planOf(figures(count)).shape('Block 1 label').size]);
  assert.deepEqual(sizes, [[85, 22.5], [69, 22.5], [56.5, 21.5], [48.5, 18.5], [40, 16], [28.5, 14]]);
  // Fewer figures are never set smaller than more in the same room: from one to eight, and ten, twelve, sixteen
  // and twenty-four. (Nine stand three by three in lower cards than ten in two lines of five, and are smaller.)
  const byCount = [1, 2, 3, 4, 5, 6, 7, 8, 10, 12, 16, 24].map((count) => planOf(figures(count)).shape('Block 1 value').size ?? 0);
  byCount.forEach((size, index) => assert.ok(index === 0 || size <= (byCount[index - 1] as number), byCount.join(', ')));
  for (const count of [2, 3, 4, 6, 12]) {
    const plan = planOf(figures(count));
    assert.deepEqual(validateDeck(template, deckOf([figures(count)])).filter((issue) => issue.slot !== undefined), [], `${count} figures`);
    for (const block of plan.plan.layout.blocks) {
      const [value, label, panel] = [plan.shape(`Block ${block.index} value`), plan.shape(`Block ${block.index} label`), plan.shape(`Block ${block.index} panel`)];
      // The value's box is one line of the value high, and the label's lines follow it at a fifth of its size at most.
      assert.ok(Math.abs(value.box.h * H - 1.2 * (value.size ?? 0)) < 0.06, `${count} figures`);
      assert.ok(label.box.y - (value.box.y + value.box.h) <= (0.2 * (value.size ?? 0)) / H + 0.0002);
      // Both lie inside the card, and the pair is as far from its top as from its foot, give or take the one line PowerPoint may not use.
      const above = (value.box.y - panel.box.y) * H;
      const below = (panel.box.y + panel.box.h - label.box.y - label.box.h) * H;
      assert.ok(above > 0 && below > 0 && Math.abs(above - below) <= 1.3 * (label.size ?? 0) + 0.5, `${count} figures: ${above.toFixed(1)} pt above, ${below.toFixed(1)} pt below`);
    }
  }
  // A figure that is the slide's only block is the slide: no panel, and it starts where the title starts.
  const one = planOf(figures(1));
  assert.equal(one.plan.shapes.some((shape) => shape.kind === 'panel'), false);
  assert.equal(one.shape('Block 1 value').box.x, one.kit.area.x);
  assert.ok(one.shape('Block 1 value').box.y > one.kit.band.top + 0.03);
  // A label is never set under a size a person reads where its card has the room: it follows its value.
  const [three] = [planOf(figures(3))];
  assert.ok((three.shape('Block 2 label').size ?? 0) >= 16);
  // And it is never larger than the value over it, however long a value somebody writes.
  const sentence = planOf(freeSlide({ title: 'The year in numbers', blocks: [{ figure: { value: 'A value that is a whole sentence and has no business being a figure at all, ninety characters', label: 'rides' } }] }));
  assert.ok((sentence.shape('Block 1 label').size ?? 0) <= (sentence.shape('Block 1 value').size ?? 0));
});

test('a short last line: cards, steps and figures spread or stand in the middle, charts start where the lines above start, text shares the width', () => {
  const xs = (kinds: FreeBlockKind[], arrange?: { columns: number }) => arrangeBlocks(kinds, { frame, rtl: false, ...(arrange ? { arrange } : {}) }).blocks.map((block) => [block.box.x, block.box.w]);
  const right = frame.x + frame.w;
  // Five cards are three and two: the two share the whole width, and their outer edges stand under the outer edges above.
  const five = xs(Array<FreeBlockKind>(5).fill('card'));
  assert.deepEqual([five[3]?.[0], Math.round(((five[4]?.[0] ?? 0) + (five[4]?.[1] ?? 0)) * 10000) / 10000, five[3]?.[1]], [frame.x, right, 0.45]);
  assert.deepEqual(xs(Array<FreeBlockKind>(5).fill('figure')).slice(3).map((box) => box[1]), [0.45, 0.45]);
  // Steps too: the two of a short last line take the whole width, in their order, and no hole stands beside them.
  const steps = xs(Array<FreeBlockKind>(5).fill('step'), { columns: 3 });
  assert.deepEqual([steps[3]?.[0], Math.round(((steps[4]?.[0] ?? 0) + (steps[4]?.[1] ?? 0)) * 10000) / 10000], [frame.x, right]);
  // A last line that holds fewer than half as many stands in the middle, as wide as the others: one card under three.
  const four = xs(Array<FreeBlockKind>(4).fill('card'), { columns: 3 });
  assert.deepEqual([four[3]?.[1], Math.round(((four[3]?.[0] ?? 0) + (four[3]?.[1] ?? 0) / 2) * 1000) / 1000], [four[0]?.[1], Math.round((frame.x + frame.w / 2) * 1000) / 1000]);
  // Charts and pictures are a raster: the fifth stands under the first, as wide as the first.
  for (const kind of ['chart', 'photo', 'graphic'] as const) {
    const line = xs(Array<FreeBlockKind>(5).fill(kind), { columns: 3 });
    assert.deepEqual([line[3], line[4]], [line[0], line[1]], kind);
  }
  // Running text, lists, quotations and tables have no box whose edge would have to line up: a last one takes the whole width.
  for (const kind of ['text', 'list', 'quote', 'table'] as const) assert.deepEqual(xs(Array<FreeBlockKind>(3).fill(kind), { columns: 2 })[2], [frame.x, frame.w], kind);
  // Read from the right everything is mirrored: the fifth chart stands under the first, at the right.
  const mirrored = arrangeBlocks(Array<FreeBlockKind>(5).fill('chart'), { frame, rtl: true, arrange: { columns: 3 } }).blocks;
  assert.equal(mirrored[3]?.box.x, mirrored[0]?.box.x);
  assert.ok((mirrored[0]?.box.x ?? 0) > (mirrored[1]?.box.x ?? 1));
});

test('how many stand on a line: few cards never five, figures and charts so that the room is used', () => {
  const columns = (kind: FreeBlockKind, count: number) => arrangeBlocks(Array<FreeBlockKind>(count).fill(kind), { frame, rtl: false }).columns;
  // Up to eight cards or steps never stand more than four on a line: a fifth makes every text small.
  assert.deepEqual([1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 12, 20].map((count) => columns('card', count)), [1, 2, 3, 4, 3, 3, 4, 4, 3, 5, 4, 5]);
  assert.deepEqual([5, 7, 8].map((count) => columns('step', count)), [3, 4, 4]);
  // A wish for five on a line is still followed.
  assert.equal(arrangeBlocks(Array<FreeBlockKind>(5).fill('card'), { frame, rtl: false, arrange: { columns: 5 } }).columns, 5);
  // Twelve figures stood six and six over an empty lower half: four by three uses the room.
  const twelve = arrangeBlocks(Array<FreeBlockKind>(12).fill('figure'), { frame, rtl: false });
  assert.equal(twelve.columns, 4);
  assert.ok(Math.max(...twelve.blocks.map((block) => block.box.y + block.box.h)) > frame.bottom - 0.001);
  // The arrangements that suited their blocks keep them.
  assert.deepEqual([2, 3, 4, 6, 8].map((count) => columns('figure', count)), [2, 3, 2, 3, 4]);
  assert.deepEqual([2, 3, 4, 6, 8, 12].map((count) => columns('chart', count)), [2, 3, 2, 3, 4, 4]);
  // Three charts on a line may be a little higher than wide, and end where the room ends.
  const three = arrangeBlocks(['chart', 'chart', 'chart'], { frame, rtl: false });
  assert.deepEqual(three.blocks.map((block) => [block.box.y, Math.round((block.box.y + block.box.h) * 1000) / 1000]), [[0.3, 0.88], [0.3, 0.88], [0.3, 0.88]]);
});

test('cards and steps grow into room that is left over, and hold what they held', () => {
  // Three cards whose content asks for 0.2 of the slide's width in height grow by 35%, and say by how much.
  const grown = arrangeBlocks(['card', 'card', 'card'], { frame, rtl: false, needed: () => 0.2 });
  const asked = Math.round(0.2 * frame.aspect * 10000) / 10000;
  assert.deepEqual(grown.blocks.map((block) => [Math.round(block.box.h * 10000) / 10000, block.grown]), Array(3).fill([Math.round(1.35 * 0.2 * frame.aspect * 10000) / 10000, Math.round(0.35 * 0.2 * frame.aspect * 10000) / 10000]));
  assert.ok((grown.blocks[0]?.box.h ?? 0) > asked);
  // Lying cards grow by 15%; a card is never made higher than the template draws one of its width; two lines grow alike, the second moving down with the first.
  const lying = arrangeBlocks(['card', 'card'], { frame, rtl: false, needed: () => 0.1 });
  assert.equal(lying.blocks[0]?.grown, Math.round(0.15 * 0.1 * frame.aspect * 10000) / 10000);
  const flat = arrangeBlocks(['card', 'card', 'card'], { frame: { ...frame, card: 0.45 }, rtl: false, needed: () => 0.12 });
  assert.equal(Math.round((flat.blocks[0]?.box.h ?? 0) * 10000) / 10000, Math.round(0.45 * 0.2925 * frame.aspect * 10000) / 10000);
  const six = arrangeBlocks(Array<FreeBlockKind>(6).fill('card'), { frame, rtl: false, needed: () => 0.08 });
  assert.equal(new Set(six.blocks.map((block) => block.grown)).size, 1);
  assertTidy(six, 6, 'six cards that grew');
  // Cards that fill their room, cards beside a main block, and cards on a template that draws no panel behind them stay as they are.
  assert.equal(arrangeBlocks(Array<FreeBlockKind>(6).fill('card'), { frame, rtl: false, needed: () => 0.2 }).blocks[0]?.grown, undefined);
  assert.equal(arrangeBlocks(['chart', 'card', 'card', 'card'], { frame, rtl: false }).blocks[1]?.grown, undefined);
  assert.equal(arrangeBlocks(['card', 'card', 'card'], { frame: { ...frame, card: 0 }, rtl: false, needed: () => 0.12 }).blocks[0]?.grown, undefined);
  // A card grows only while what is written on it, from its top to the foot of its text, fills 60% of it:
  // three cards of a short line each do not grow at all, and the slide settles them in its room.
  assert.deepEqual(planOf(threeCards).plan.layout.blocks.map((entry) => entry.grown), [undefined, undefined, undefined]);
  const fuller = freeSlide({ title: 'Three ways in', blocks: [card('Longer hours', 'Bikes unlock from five in the morning, for early shifts and the first trams.'), card('Safer docks', 'Every station gets lighting and a camera before the season opens.'), card('Group passes', 'One pass for up to five riders, bought once and shared by link.')] });
  // On the slide the panel is as high as the card grew, and heading and text are set out as in the card that was asked for.
  const plan = planOf(fuller);
  const block = plan.plan.layout.blocks[0];
  assert.ok((block?.grown ?? 0) > 0.02);
  // What is written still reaches 60% down the card that grew, as PowerPoint sets its lines (at about 0.42 of its size a character).
  const textShape = plan.shape('Block 1 text');
  const lines = Math.ceil('Bikes unlock from five in the morning, for early shifts and the first trams.'.length / Math.floor((textShape.box.w * W) / (0.42 * (textShape.size ?? 1))));
  const foot = textShape.box.y + (lines * 1.2 * (textShape.size ?? 0)) / H;
  assert.ok(foot - (block?.box.y ?? 0) >= 0.6 * (block?.box.h ?? 0) - 0.002, `written to ${(foot - (block?.box.y ?? 0)).toFixed(4)} of ${block?.box.h}`);
  assert.equal(plan.shape('Block 1 panel').box.h, block?.box.h);
  assert.ok(plan.shape('Block 1 text').box.y + plan.shape('Block 1 text').box.h <= (block?.box.y ?? 0) + (block?.box.h ?? 0) - (block?.grown ?? 0) + 0.0001);
  // So a slot holds the same whether the card has room to grow or not: three cards alone, and the same three over two long bottom lines that take the room.
  const line = 'Rides grew at every station this season, and the docks next to the tram stops emptied first on every weekday morning of the year, in rain and in sun.';
  const over = planOf(freeSlide({ title: 'Three ways in', blocks: [...freeContent(fuller).blocks, { note: line }, { note: line }] }));
  assert.ok((over.plan.layout.blocks[0]?.grown ?? 0) < (block?.grown ?? 0));
  assert.deepEqual([(plan.slot('b1Text') as FreeTextSlot).maxChars, plan.shape('Block 1 text').size, plan.shape('Block 1 heading').size], [(over.slot('b1Text') as FreeTextSlot).maxChars, over.shape('Block 1 text').size, over.shape('Block 1 heading').size]);
});

test('the bottom line closes the blocks over it, tiles of a line are equally high, and blocks beside a main block end where it ends', () => {
  // A note that is not first stands one gap under the lowest block, never under a hole that the blocks leave.
  // Cards, figures and charts that are lower than their room sit a little down it with their note as one group;
  // a list hangs from the top line with its note under it. What the group leaves is over it and under it.
  for (const kinds of [['card', 'card', 'card', 'note'], ['figure', 'figure', 'note'], ['list', 'note'], ['chart', 'chart', 'note']] as FreeBlockKind[][]) {
    const layout = arrangeBlocks(kinds, { frame, rtl: false, needed: () => 0.1 });
    const note = layout.blocks[layout.blocks.length - 1];
    const rest = layout.blocks.slice(0, -1);
    const lowest = rest.reduce((most, block) => Math.max(most, block.box.y + block.box.h), 0);
    assert.ok(Math.abs((note?.box.y ?? 0) - lowest - frame.gap * frame.aspect) < 0.0002, `${kinds.join()}: ${note?.box.y} under ${lowest}`);
    assert.ok((note?.box.y ?? 0) + (note?.box.h ?? 0) < frame.bottom - 0.001, kinds.join());
    assert.equal(Math.abs((rest[0]?.box.y ?? 0) - frame.top) < 0.0002, kinds[0] === 'list', kinds.join());
    assertTidy(layout, kinds.length, kinds.join());
  }
  // Blocks that fill their room have their note at its foot.
  const filled = arrangeBlocks(['chart', 'chart', 'note'], { frame, rtl: false });
  assert.ok(Math.abs((filled.blocks[2]?.box.y ?? 0) + (filled.blocks[2]?.box.h ?? 0) - frame.bottom) < 0.0002);
  // Tiles of one line are equally high, whatever they hold.
  const tiles = arrangeBlocks(['photo', 'chart', 'text', 'figure', 'card', 'list'], { frame, rtl: false });
  assert.equal(tiles.tiles, true);
  for (const y of new Set(tiles.blocks.map((block) => block.box.y))) assert.equal(new Set(tiles.blocks.filter((block) => block.box.y === y).map((block) => block.box.h)).size, 1);
  // A short last line of tiles shares the width.
  const five = arrangeBlocks(['figure', 'figure', 'figure', 'chart', 'table'], { frame, rtl: false });
  assert.deepEqual([five.tiles, five.blocks[3]?.box.x, Math.round(((five.blocks[4]?.box.x ?? 0) + (five.blocks[4]?.box.w ?? 0)) * 10000) / 10000], [true, frame.x, frame.x + frame.w]);
  // Three figures beside a photo are a column that ends where the photo ends, not two and one; four are two by two.
  const beside = arrangeBlocks(['photo', 'figure', 'figure', 'figure'], { frame, rtl: false });
  const lead = beside.blocks[0];
  assert.equal(new Set(beside.blocks.slice(1).map((block) => block.box.x)).size, 1);
  assert.ok(Math.abs((beside.blocks[3]?.box.y ?? 0) + (beside.blocks[3]?.box.h ?? 0) - ((lead?.box.y ?? 0) + (lead?.box.h ?? 0))) < 0.0002);
  const fourBeside = arrangeBlocks(['chart', 'figure', 'figure', 'figure', 'figure'], { frame, rtl: false });
  assert.equal(new Set(fourBeside.blocks.slice(1).map((block) => block.box.x)).size, 2);
  assert.ok(Math.abs((fourBeside.blocks[4]?.box.y ?? 0) + (fourBeside.blocks[4]?.box.h ?? 0) - ((fourBeside.blocks[0]?.box.y ?? 0) + (fourBeside.blocks[0]?.box.h ?? 0))) < 0.0002);
  // A text over cards sits a little down the room with them; a slide of text alone hangs from the top line.
  assert.ok((arrangeBlocks(['text', 'card', 'card', 'card'], { frame, rtl: false, needed: (_index, kind) => (kind === 'card' ? 0.05 : undefined) }).blocks[0]?.box.y ?? 0) >= frame.top);
  assert.equal(arrangeBlocks(['text', 'list'], { frame, rtl: false }).blocks[0]?.box.y, frame.top);
});

test('a slide takes the larger room only when that gives a block more room', () => {
  const kinds: FreeBlockKind[] = ['text', 'list', 'card', 'figure', 'step', 'quote', 'chart', 'table', 'photo', 'graphic'];
  const usualOnly: LayoutFrame = { ...frame, denseTop: frame.top, denseBottom: frame.bottom };
  let seed = 11;
  const next = () => (seed = (seed * 1103515245 + 12345) % 2147483648) / 2147483648;
  let dense = 0;
  for (let round = 0; round < 400; round++) {
    const count = 1 + Math.floor(next() * 30);
    const uniform = round % 2 === 0;
    const kind = kinds[Math.floor(next() * kinds.length)] as FreeBlockKind;
    const blocks = Array.from({ length: count }, () => (uniform ? kind : (kinds[Math.floor(next() * kinds.length)] as FreeBlockKind)));
    const wish = round % 5 === 0 ? { columns: 1 + Math.floor(next() * 8) } : undefined;
    const options = { rtl: false, needed: () => 0.06, ...(wish ? { arrange: wish } : {}) };
    const [got, usual] = [arrangeBlocks(blocks, { frame, ...options }), arrangeBlocks(blocks, { frame: usualOnly, ...options })];
    if (!got.dense) {
      assert.deepEqual(boxes(got), boxes(usual), `round ${round}: ${blocks.join()}`);
      continue;
    }
    dense++;
    const larger = got.tiles !== usual.tiles || got.columns !== usual.columns || got.blocks.some((block, index) => block.box.w > (usual.blocks[index]?.box.w ?? 0) + 0.0001 || block.box.h > (usual.blocks[index]?.box.h ?? 0) + 0.0001);
    assert.ok(larger, `round ${round}: ${blocks.join()} moved up for nothing`);
  }
  assert.ok(dense > 40, `${dense} slides took the larger room`);
});

test('a subtitle is the title\'s line only where the free slide shows something on it', () => {
  // The demo's text layouts have a line named subtitle between the title and the body. A free slide has no
  // subtitle of its own, and its frame draws nothing on that line: the blocks start on it, where the layouts'
  // first line starts, as though it were named otherwise.
  const renamed: TemplateManifest = { ...manifest, patterns: manifest.patterns.map((pattern) => ({ ...pattern, slots: pattern.slots.map((slot) => (slot.name === 'subtitle' ? { ...slot, name: 'lede' } : slot)) })) };
  const [plain, on] = [freeKit(template, 'text-only'), freeKit(loadTemplate(demoPptx, renamed), 'text-only')];
  assert.deepEqual([plain?.base, on?.base], ['text-only', 'text-only']);
  assert.deepEqual(plain?.band, on?.band);
  // Where that line holds a text of the deck that stays on every slide, the line is the title's: the blocks
  // start under it, and the room ends lower by what that takes, no further than the layouts' own content.
  const kept: TemplateManifest = { ...manifest, patterns: manifest.patterns.map((pattern) => ({ ...pattern, slots: pattern.slots.map((slot) => (slot.name === 'subtitle' && slot.kind === 'text' ? { ...slot, deckField: 'organizationDepartment' } : slot)) })) };
  const under = freeKit(loadTemplate(demoPptx, kept), 'text-only');
  assert.ok((under?.band.top ?? 0) > (on?.band.top ?? 1) + 0.05, `${under?.band.top} against ${on?.band.top}`);
  assert.ok((under?.band.top ?? 0) <= (under?.title?.box.y ?? 0) + (under?.title?.box.h ?? 0) + 0.09 + 0.0001);
  assert.ok((under?.band.bottom ?? 0) >= (on?.band.bottom ?? 1));
  // The layouts the demo's usual frame is read from have no subtitle: nothing changes there.
  assert.deepEqual(freeKit(template)?.band, freeKit(loadTemplate(demoPptx, renamed))?.band);
});

test('the room is made wider at the side reading ends where it leaves clearly more margin than at the side reading starts', () => {
  const area = { x: 0.08, y: 0.15, w: 0.88, h: 0.8 };
  const title = { x: 0.08, y: 0.15, w: 0.88, h: 0.07 };
  const shape = (id: string, box: { x: number; y: number; w: number; h: number }) => ({ id, name: id, kind: 'shape', box }) as unknown as Parameters<typeof widened>[2][number];
  const frameShapes = [shape('rule', { x: 0.045, y: 0.07, w: 0.91, h: 0.002 }), shape('logo', { x: 0.045, y: 0.03, w: 0.05, h: 0.03 })];
  // Read from the right, the margin at the right is 4% and at the left 8%: the room reaches as far left as the rule over it does.
  assert.deepEqual(widened(area, 'rtl', frameShapes, title), { x: 0.045, y: 0.15, w: 0.915, h: 0.8 });
  // With nothing above to say how far the frame reaches, the margin at the start side is the measure.
  assert.deepEqual(widened(area, 'rtl', [], title), { x: 0.04, y: 0.15, w: 0.92, h: 0.8 });
  // Read from the left, the same area is not widened: its margins at the right is the smaller one already.
  assert.deepEqual(widened(area, 'ltr', frameShapes, title), area);
  // A difference under a hundredth of the slide is left alone, and so is an area with something that stays in the strip.
  assert.deepEqual(widened({ x: 0.045, y: 0.15, w: 0.915, h: 0.8 }, 'rtl', frameShapes, title), { x: 0.045, y: 0.15, w: 0.915, h: 0.8 });
  assert.deepEqual(widened(area, 'rtl', [...frameShapes, shape('band', { x: 0.03, y: 0.3, w: 0.04, h: 0.5 })], title), area);
  // The demo's margins are alike: its room is its content area.
  assert.deepEqual([freeKit(template)?.area.x, freeKit(template)?.area.w], [0.0375, 0.9225]);
});

test('the source of a quotation is the small line right under it; a quotation made from a heading is plain', () => {
  const quote = 'I sold my second car in June. The bike gets me to the tram in six minutes, and I have not missed it once.';
  const alone = planOf(freeSlide({ title: 'A rider', blocks: [{ quote, source: 'A rider, in the end-of-season survey' }] }));
  const [text, source] = [alone.shape('Block 1 quote'), alone.shape('Block 1 source')];
  // Half the quotation's size, never over running text; in the template's accent, the colour of its big numbers.
  assert.deepEqual([text.size, source.size, source.color, alone.kit.styles.figure.color], [30.5, 15, '0E7C7B', '0E7C7B']);
  // The quotation's text stands at the foot of its box, so the source follows its last line at half its size, however many lines PowerPoint sets.
  assert.equal(text.anchor, 'b');
  assert.ok(Math.abs((source.box.y - text.box.y - text.box.h) * H - 0.5 * (text.size ?? 0)) < 0.06);
  assert.match(slideXml(renderDeck(template, deckOf([freeSlide({ title: 'A rider', blocks: [{ quote, source: 'A rider' }] })])).pptx, 1), /name="Block 1 quote".*?anchor="b"/);
  // Quotations side by side have boxes of one height, and their sources stand on one line. Each has its text at
  // the foot of its box, so every source stands as far under its own last line, however long the quotation.
  const three = planOf(freeSlide({ title: 'Riders', blocks: [{ quote: 'The answer came before I expected it, and I did not have to ask twice.', source: 'A resident' }, { quote: 'One form.', source: 'A shop owner' }, { quote: 'I never had to take a day off for it.', source: 'A contractor' }] }));
  assert.equal(three.shape('Block 1 quote').box.y, three.shape('Block 2 quote').box.y);
  assert.deepEqual([three.shape('Block 1 quote').anchor, three.shape('Block 2 quote').anchor, three.shape('Block 3 quote').anchor], ['b', 'b', 'b']);
  assert.equal(three.shape('Block 1 quote').box.h, three.shape('Block 2 quote').box.h);
  assert.equal(three.shape('Block 1 source').box.y, three.shape('Block 2 source').box.y);
  // A template with no quotation of its own has one made from its heading: light, in the colour of running text,
  // slanted when it is read from the left and upright when it is read from the right, whose letters have no slant.
  const without = only(...manifest.patterns.map((pattern) => pattern.id).filter((id) => id !== 'quote'));
  const derived = freeKit(loadTemplate(demoPptx, without))?.styles;
  assert.deepEqual([derived?.quote.shapeId, derived?.quote.set], [derived?.heading.shapeId, { italic: true, bold: false, color: derived?.body.color, size: 29 }]);
  assert.equal(freeKit(loadTemplate(demoPptx, { ...without, direction: 'rtl' }))?.styles.quote.set?.italic, false);
});

test('a note that fits one line is one line high, and a note is no louder than a heading', () => {
  const short = 'The bottom line: dependable docks turned a weekend habit into a daily commute.';
  const slide = (note: string) => freeSlide({ title: 'Rides', blocks: [{ chart }, { note }] });
  const band = (note: string) => planOf(slide(note)).plan.layout.blocks[1]?.box.h ?? 0;
  // 78 characters fit one line of 17 pt across the slide: the band is one line and the space around it, where it was two lines.
  const one = planOf(slide(short));
  assert.ok(Math.abs(band(short) * H - (1.2 * 17 + 2 * 0.6 * 1.2 * 17)) < 0.3, `${(band(short) * H).toFixed(1)} pt`);
  assert.equal((one.slot('b2Text') as FreeTextSlot).maxChars, Math.floor((one.shape('Block 2 note').box.w * W - 2 * 0.019 * W) / (0.55 * 17)));
  // A note that does not fit one line has two, counted with the loss at the break, as before.
  assert.ok(band(`${short} ${short}`) > 1.4 * band(short));
  for (const note of [short, `${short} ${short}`]) assert.deepEqual(validateDeck(template, deckOf([slide(note)])), []);
  // The demo's note is set at 17 pt under headings of 19. A template whose note text is larger has it at nine tenths of a heading, and never under running text.
  const loud = loadTemplate(demoWith({ 'ppt/slides/slide9.xml': (xml) => xml.replace(/(<p:cNvPr id="8" .*?<a:rPr[^>]*? sz=")\d+(")/s, '$12800$2') }), manifest);
  assert.equal(freeKit(loud)?.styles.note?.size, 28);
  assert.equal(planOf(slide(short), {}, loud).shape('Block 2 note').size, 18);
});

test('a text that can take several lines has them at least as far apart as its size asks, whatever its source', () => {
  // The demo's quotation and its card heading with their lines set at 70%, as a template may set a heading of one line.
  const tight = (xml: string) => xml.replace(/<a:pPr([^>]*)>/g, '<a:pPr$1><a:lnSpc><a:spcPct val="70000"/></a:lnSpc>').replace(/<a:pPr([^>]*)\/>/g, '<a:pPr$1><a:lnSpc><a:spcPct val="70000"/></a:lnSpc></a:pPr>');
  const squeezed = loadTemplate(demoWith({ 'ppt/slides/slide11.xml': tight, 'ppt/slides/slide5.xml': tight }), manifest);
  const slide = freeSlide({ title: 'Riders', blocks: [{ quote: 'I sold my second car in June. The bike gets me to the tram in six minutes.', source: 'A rider' }, card('Longer hours', 'Bikes unlock from five.'), card('A heading that is long enough to take two lines', 'Every station gets a light.')] });
  const xml = slideXml(renderDeck(squeezed, deckOf([slide])).pptx, 1);
  const spacingOf = (name: string) => new RegExp(`name="${name}".*?<a:lnSpc><a:spcPct val="(\\d+)"/>`).exec(xml)?.[1];
  // The quotation and the headings of a slide on which a heading takes two lines are set at 100%.
  assert.deepEqual([spacingOf('Block 1 quote'), spacingOf('Block 2 heading'), spacingOf('Block 3 heading')], ['100000', '100000', '100000']);
  // A heading of one line keeps the spacing of its source, and with it its place.
  const single = slideXml(renderDeck(squeezed, deckOf([freeSlide({ title: 'Ways in', blocks: [card('Longer hours', 'Bikes unlock from five.'), card('Safer docks', 'Every station gets a light.')] })])).pptx, 1);
  assert.equal(new RegExp('name="Block 1 heading".*?<a:lnSpc><a:spcPct val="(\\d+)"/>').exec(single)?.[1], '70000');
});

test('a heading of three lines has its three lines, or two at a smaller size, and the text starts under them', () => {
  const texts = ['Residents reach every service through one phone number and one website.', 'Budgets and contracts are published every month.', 'A pothole is filled within three days.', 'Six night lines run until two.'];
  const cards = (heading: string) => freeSlide({ title: 'Ways in', blocks: [card(heading, texts[0] as string), card('Open data', texts[1] as string), card('Street repairs', texts[2] as string), card('Night buses', texts[3] as string)] });
  const linesOf = (plan: ReturnType<typeof planOf>) => Math.round((plan.shape('Block 1 heading').box.h * H - 0.15 * (plan.shape('Block 1 heading').size ?? 0)) / (1.2 * (plan.shape('Block 1 heading').size ?? 1)));
  const [one, two, three] = [planOf(cards('Front door')), planOf(cards('A single front door now')), planOf(cards('A single front door for every city service'))];
  assert.deepEqual([linesOf(one), linesOf(two), linesOf(three)], [1, 2, 3]);
  // Every card of the slide keeps the lines of the longest heading, so that the texts start on one line under them.
  for (const plan of [one, two, three]) {
    assert.equal(new Set(plan.plan.layout.blocks.map((block) => plan.shape(`Block ${block.index} text`).box.y)).size, 1);
    assert.ok(plan.shape('Block 1 text').box.y >= plan.shape('Block 1 heading').box.y + plan.shape('Block 1 heading').box.h - 0.0001);
  }
  // A heading that takes two lines once it is set up to 15% smaller is set so, and keeps to two.
  const nearly = planOf(cards('Residents reach all of the services'));
  assert.deepEqual([linesOf(nearly), (nearly.shape('Block 1 heading').size ?? 0) < (two.shape('Block 1 heading').size ?? 0), (nearly.shape('Block 1 heading').size ?? 0) >= 0.85 * (two.shape('Block 1 heading').size ?? 0) - 0.01], [2, true, true]);
  // What a heading may be is still two lines: a longer one is told on its slot, though it is drawn without lying on its text.
  assert.deepEqual(validateDeck(template, deckOf([cards('A single front door for every city service')])).map((issue) => issue.slot), ['b1Heading']);
});

test('a text alone keeps a line of the same length under a heading, and pictures and their captions start where reading starts', () => {
  const para = 'Two years ago the city answered a request for a building permit in thirty one working days. The file went from planning to fire safety to utilities, and each office started only when the one before it had finished.';
  const [bare, headed] = [planOf(freeSlide({ title: 'What changed', blocks: [{ text: para }] })).shape('Block 1 text'), planOf(freeSlide({ title: 'What changed', blocks: [{ text: para, heading: 'What changed' }] })).shape('Block 1 text')];
  // Sixty-six characters to a line at the size the text is set at: 22 pt alone, 18 pt under a heading, which the line follows.
  assert.deepEqual([bare.size, headed.size], [22, 18]);
  assert.ok(Math.abs(bare.box.w * W - 66 * 0.55 * 22) < 0.6 && Math.abs(headed.box.w * W - 66 * 0.55 * 18) < 0.6, `${bare.box.w * W} and ${headed.box.w * W} pt`);
  // A photo fills its box by being cut, up to 3.2 times as wide as high: one photo alone is as wide as the room.
  const one = planOf(freeSlide({ title: 'The pier', blocks: [{ photo: 'pier', caption: 'At dawn' }] }));
  assert.deepEqual([one.shape('Block 1 photo').box.x, one.shape('Block 1 photo').box.w], [one.kit.area.x, one.kit.area.w]);
  // Its caption starts under its corner and is as wide as it is, also where the photo is cut narrower than its box.
  const strip = planOf(freeSlide({ title: 'The pier', blocks: [{ chart }, { chart }, { chart }, { photo: 'pier', caption: 'At dawn' }], columns: 3 }));
  for (const plan of [one, strip]) {
    const index = plan.plan.layout.blocks.find((block) => block.kind === 'photo')?.index;
    const [photo, caption] = [plan.shape(`Block ${index} photo`), plan.shape(`Block ${index} caption`)];
    assert.deepEqual([caption.box.x, caption.box.w, caption.align], [photo.box.x, photo.box.w, 'start']);
    assert.ok(photo.box.w * W <= 3.2 * photo.box.h * H + 0.6);
  }
  const wide = planOf(freeSlide({ title: 'The pier', blocks: [{ photo: 'pier', caption: 'At dawn' }, { note: 'One line.' }, { note: 'Another line.' }, { note: 'A third line.' }] }));
  const [cut, block] = [wide.shape('Block 1 photo'), wide.plan.layout.blocks[0]];
  assert.ok(cut.box.w < (block?.box.w ?? 0) - 0.01 && cut.box.x === block?.box.x, 'a photo cut narrower stands at the edge reading starts at');
  const mirrored = planOf(freeSlide({ title: 'The pier', blocks: [{ photo: 'pier', caption: 'At dawn' }, { note: 'One line.' }, { note: 'Another line.' }, { note: 'A third line.' }] }), { direction: 'rtl' });
  assert.ok(Math.abs(mirrored.shape('Block 1 photo').box.x + mirrored.shape('Block 1 photo').box.w - ((mirrored.plan.layout.blocks[0]?.box.x ?? 0) + (mirrored.plan.layout.blocks[0]?.box.w ?? 0))) < 0.0002);
  // On a slide of far too many blocks a photo is small, or cut to a sliver: its caption keeps the width of a short line, as far as its block has it.
  const crowd = planOf(freeSlide({ title: 'The pier', blocks: Array.from({ length: 40 }, (): FreeBlockContent => ({ photo: 'pier', caption: 'At dawn' })) }));
  const [small, under] = [crowd.shape('Block 7 photo'), crowd.shape('Block 7 caption')];
  assert.ok(under.box.w >= small.box.w - 0.0001 && under.box.w * W >= Math.min((crowd.plan.layout.blocks[6]?.box.w ?? 0) * W, 12 * 0.55 * (under.size ?? 0)) - 0.6, `a caption ${(under.box.w * W).toFixed(1)} pt wide under a photo ${(small.box.w * W).toFixed(1)} pt wide`);
  // A graphic is shown whole in the middle of its box: its caption stands under its middle.
  assert.equal(planOf(freeSlide({ title: 'The map', blocks: [{ graphic: 'map', caption: 'Stations' }] })).shape('Block 1 caption').align, 'ctr');
});

test('the columns of a table are as wide as what is written in them asks, and a column of numbers ends on one edge', () => {
  const phrase = 'Tram Square station, platform B, by the river gate';
  const table = { header: ['Station', 'Docks', 'Rides a day', 'Change'], rows: [[phrase, '24', '1,410', '+4.2%'], ['River Gate', '18', '365', '-1.5%'], ['Market Hall', '16', '290', '+0.3%']] };
  const slide = freeSlide({ title: 'Stations', blocks: [{ table }] });
  const slot = planOf(slide).slot('b1Table') as FreeTableSlot;
  // One column of long names among three of short numbers: it has the width, the text stays at the template's size and no row is higher than another.
  const widths = slot.columnWidths ?? [];
  assert.deepEqual([slot.bodyText?.size, slot.rowHeights, widths.length], [15, undefined, 4]);
  assert.ok((widths[0] as number) > 2.5 * (widths[2] as number) && (widths[2] as number) > (widths[1] as number), widths.join(', '));
  assert.ok(Math.abs(widths.reduce((sum, width) => sum + width, 0) - slot.box.w * W) < 1);
  // The numbers end on one edge, their headings with them; the names start where reading starts.
  assert.deepEqual(slot.columnAligns, [undefined, 'r', 'r', 'r']);
  assert.deepEqual(validateDeck(template, deckOf([slide])), []);
  const xml = slideXml(renderDeck(template, deckOf([slide])).pptx, 1);
  const columns = [...xml.matchAll(/<a:gridCol w="(\d+)"/g)].map((match) => Number(match[1]) / 12700);
  columns.forEach((width, index) => assert.ok(Math.abs(width - (widths[index] as number)) < 0.6, `column ${index + 1}: ${width} pt`));
  const firstRow = /<a:tr\b.*?<\/a:tr>/s.exec(xml)?.[0] ?? '';
  assert.deepEqual([...firstRow.matchAll(/<a:pPr[^>]*algn="(\w)"/g)].map((match) => match[1]), ['l', 'r', 'r', 'r']);
  assert.equal((xml.match(/<a:tcPr[^>]*anchor="ctr"/g) ?? []).length, 16);
  // Read from the right every column ends on the right already: nothing is set apart.
  assert.equal((planOf(slide, { direction: 'rtl' }).slot('b1Table') as FreeTableSlot).columnAligns, undefined);
  // Cells of one length leave the columns alike, and a table before it is written has no widths to state.
  const even = planOf(freeSlide({ title: 'Stations', blocks: [{ table: { header: ['Aaaa', 'Bbbb'], rows: [['cccc', 'dddd']] } }] })).slot('b1Table') as FreeTableSlot;
  assert.deepEqual([even.columnWidths, even.columnAligns], [undefined, undefined]);
  // A column of long cells among short ones used to set the whole table smaller: it stays at the template's size.
  const station = 'Tram Square station, platform B, by the river gate and the market hall';
  const oneLong = (length: number) => freeSlide({ title: 'Stations', blocks: [{ table: { header: ['Station', 'Docks', 'Rides', 'Open', 'Notes'], rows: Array.from({ length: 9 }, (_, row) => [`Station ${row + 1}`, `${station} ${station}`.slice(0, length), '24', '410', 'Open']) } }, { note: 'Tram Square leads.' }] });
  assert.deepEqual([8, 24, 31, 60].map((length) => (planOf(oneLong(length)).slot('b1Table') as FreeTableSlot).bodyText?.size), [15, 15, 15, 15]);
  for (const length of [31, 60]) assert.deepEqual(validateDeck(template, deckOf([oneLong(length)])), []);
});

test('a title of two lines has its blocks start under its second line', () => {
  const long = 'Services review 2025: what residents asked for, what the city delivered and what comes next';
  // Three charts fill their room, so they start where it starts.
  const cards = (title: string) => freeSlide({ title, blocks: [{ chart }, { chart }, { chart }] });
  const [short, two] = [planOf(cards('Three ways in')), planOf(cards(long))];
  const topOf = (plan: ReturnType<typeof planOf>) => Math.min(...plan.plan.layout.blocks.map((block) => block.box.y));
  const title = two.kit.title;
  assert.ok(title);
  // Ninety characters of 34 pt take two lines of the title's width: the blocks start under them, as far under
  // the second line as they start under a title of one line. The room starts lower by the line the title has more.
  const under = title.box.y + (2 * 1.2 * title.size + 0.5 * two.kit.gap * W) / H;
  const dropped = two.kit.band.top + (1.2 * title.size) / H;
  assert.ok(dropped > under && Math.abs(topOf(two) - dropped) < 0.0006, `blocks from ${topOf(two)}, the title's lines end at ${under}`);
  assert.equal(topOf(short), short.kit.band.top);
  // The title itself is still told that it is longer than a title may be.
  assert.deepEqual(validateDeck(template, deckOf([cards(long)])).map((issue) => issue.slot), ['title']);
  // Before there is content a title is one line.
  const planned = freeFor(template, {}, { blocks: ['chart', 'chart', 'chart'], slots: {} }, { icons: false });
  assert.ok(typeof planned !== 'string');
  assert.equal(Math.min(...planned.plan.layout.blocks.map((block) => block.box.y)), planned.kit.band.top);
});

// ---- The second mending round of the design pass ------------------------------------------------

test('a card heading keeps its size when its text is set smaller, and card text is never larger than the running text of its slide', () => {
  const long = 'Bikes unlock from five in the morning for early shifts and the first trams, and every station gets lighting and a camera before the season opens so that riders feel safe.';
  const six = (text: string) => freeSlide({ title: 'Six', blocks: Array.from({ length: 6 }, (_, index) => card(`Point ${index + 1}`, text)) });
  const short = planOf(six('Bikes unlock from five.'));
  const full = planOf(six(long.slice(0, 140)));
  // The long text is set smaller; its heading is not: a heading of a slide is not set smaller because its text is long.
  assert.ok((full.shape('Block 1 text').size ?? 0) < (short.shape('Block 1 text').size ?? 0));
  assert.equal(full.shape('Block 1 heading').size, short.shape('Block 1 heading').size);
  // Short card text is set larger under a large heading, but not larger than the running text of the slide.
  const loud = loadTemplate(demoPptx, { ...manifest, free: { styles: { heading: { slide: 11, shape: '6' } } } });
  const cards = [card('Longer hours', 'Bikes unlock from five.'), card('Safer docks', 'Every station gets a light.'), card('Group passes', 'One pass for five riders.')];
  const alone = planOf(freeSlide({ title: 'Three ways in', blocks: cards }), {}, loud);
  const led = planOf(freeSlide({ title: 'Three ways in', blocks: [{ text: 'Three changes this year made the bikes a daily habit for commuters across the city.' }, ...cards] }), {}, loud);
  assert.deepEqual([alone.shape('Block 1 text').size, led.shape('Block 2 text').size, led.shape('Block 1 text').size], [22.5, 18, 18]);
});

test('card text that its box would not hold as PowerPoint sets it is drawn smaller, inside its card, and still told', () => {
  const long = 'Bikes unlock from five in the morning for early shifts and the first trams, and every station gets lighting and a camera before the season opens so that riders feel safe.';
  const slide = freeSlide({ title: 'Twelve', blocks: Array.from({ length: 12 }, (_, index) => card(`Point ${index + 1}`, long)) });
  const plan = planOf(slide);
  // The text is far over its budget: the slide takes the larger room, where its cards are higher, and the text is
  // drawn at the size at which it stays in its card, never under 9 pt.
  assert.equal(plan.plan.layout.dense, true);
  const text = plan.shape('Block 1 text');
  const panel = plan.shape('Block 1 panel');
  assert.ok((text.size ?? 0) >= 9 && text.box.y + text.box.h <= panel.box.y + panel.box.h + 0.0001);
  // Its budget is stated at the size it was fitted at, so that the deck is told the text is too long.
  const issues = validateDeck(template, deckOf([slide])).filter((issue) => issue.slot === 'b1Text');
  assert.match(issues[0]?.message ?? '', /^text is 170 characters; the slot fits about \d+$/);
  // The box of a card text is drawn as high as its last line may hang into the card's padding, so that what
  // PowerPoint sets stays inside its own box too, and it still ends inside the card.
  const three = planOf(threeCards);
  const box = three.shape('Block 1 text').box;
  assert.ok(box.y + box.h <= three.shape('Block 1 panel').box.y + three.shape('Block 1 panel').box.h + 0.0001);
});

test('a note closes the cards over it, and the group sits a little down the room', () => {
  const slide = freeSlide({ title: 'Three cards', blocks: [card('Point 1', 'A short line.'), card('Point 2', 'A short line.'), card('Point 3', 'A short line.'), { note: 'The bottom line: one shared file cut the wait from thirty days to nine.' }] });
  const { plan } = planOf(slide);
  const [first, , third, note] = plan.layout.blocks;
  const gap = (note?.box.y ?? 0) - ((third?.box.y ?? 0) + (third?.box.h ?? 0));
  // One gap between the cards and their note, and what they leave of the room over the cards and under the note.
  assert.ok(Math.abs(gap - 0.0225 * (16 / 9)) < 0.002, `${gap}`);
  assert.ok((first?.box.y ?? 0) > 0.3 && (note?.box.y ?? 0) + (note?.box.h ?? 0) < 0.88);
});

test('a table keeps a number with its unit, and its rows are as high as the lines PowerPoint sets them in', () => {
  const quarters = (values: number[]): ChartValue => ({ type: 'column', categories: ['Q1', 'Q2', 'Q3', 'Q4'], series: [{ name: 'Rides', values }] });
  const table = { header: ['Channel', 'Calls a month', 'Answered in', 'Solved at once', 'Rating'], rows: [['App', '41,200', '1.2 minutes', '89%', '91%'], ['Chat', '18,700', '2.4 minutes', '82%', '88%'], ['Phone', '12,300', '3.1 minutes', '77%', '84%']] };
  const plan = planOf(freeSlide({ title: 'Channels', blocks: [{ table }, { chart: quarters([33, 38, 38, 44]) }] }));
  const slot = plan.slot('b1Table') as FreeTableSlot;
  const size = slot.bodyText?.size ?? 0;
  // The column of times is as wide as "2.4 minutes" with the cell's margins, so the number does not stand over its unit.
  assert.ok((slot.columnWidths?.[2] ?? 0) >= 0.45 * 11 * size + 2 * 7.2, JSON.stringify(slot.columnWidths));
  // Rows set in as many lines are equally high, and a table beside a chart that fills its block takes some of
  // the height it leaves: up to two fifths more than its rows need.
  const rows = slot.rowHeights ?? [];
  assert.equal(new Set(rows.slice(1)).size, 1, JSON.stringify(rows));
  // A table alone hangs from the top line at its rows' own height.
  const alone = planOf(freeSlide({ title: 'Channels', blocks: [{ table }] })).slot('b1Table') as FreeTableSlot;
  assert.ok(alone.box.h < slot.box.h);
  // A column whose every cell starts with a number is aligned on its last characters, as a column of numbers is.
  const days = planOf(freeSlide({ title: 'Waits', blocks: [{ table: { header: ['Clinic', 'Visits', 'Average wait'], rows: [['Harbour', '4,120', '6 days'], ['Old Town', '3,480', '9 days']] } }] })).slot('b1Table') as FreeTableSlot;
  assert.deepEqual(days.columnAligns, [undefined, 'r', 'r']);
});

test('like full charts have numbers where they hold them, charts of one detail share one size, and low charts take the larger room', () => {
  const months = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];
  const quarters = (values: number[], more: Partial<ChartValue> = {}): ChartValue => ({ type: 'column', categories: ['Q1', 'Q2', 'Q3', 'Q4'], series: [{ name: 'Rides', values }], ...more });
  const chartSlot = (plan: ReturnType<typeof planOf>, name: string) => {
    const slot = plan.slot(name);
    assert.ok(slot.kind === 'chart');
    return slot;
  };
  // A column chart of twelve months of millions cannot hold its numbers at its columns; the bar chart of shares
  // beside it can, and keeps them: both are full charts with an axis, and they are not alike.
  const two = planOf(freeSlide({ title: 'Deposits', blocks: [{ chart: { type: 'column', categories: months, series: [{ name: 'Deposits', values: months.map((_, index) => 1250000 + 35000 * index) }] } }, { chart: quarters([47, 22, 13, 18], { type: 'bar', format: 'percent', categories: ['Savings', 'Loans', 'Cards', 'Other'] }) }] }));
  assert.deepEqual(['b1Chart', 'b2Chart'].map((name) => chartSlot(two, name).drawn?.numbers), [false, true]);
  assert.deepEqual(chartSlot(two, 'b1Chart').drawn?.axis, { max: 2000000, step: 1000000 });
  // A bar chart that names 24 categories sets its text smaller; the column and line charts beside it, drawn in the
  // same detail, set theirs at that size too.
  const names = Array.from({ length: 24 }, (_, index) => `R${index + 1}`);
  const four = planOf(freeSlide({ title: 'Four', blocks: [{ chart: { type: 'bar', categories: names, series: [{ name: 'Rides', values: names.map((_, index) => 10 + index) }] } }, { chart: quarters([3, 4, 5, 6]) }, { chart: quarters([3, 4, 5, 6]) }, { chart: quarters([3, 4, 5, 6], { type: 'line' }) }, { note: 'A bottom line for the four charts.' }] }));
  assert.deepEqual([...new Set(four.plan.pattern.slots.flatMap((slot) => (slot.kind === 'chart' ? [slot.text?.size] : [])))], [8]);
  // Eight charts over a bottom line are lower than 13% of the slide's width in the usual room: they take the
  // larger room, as ten do. Without the note they have the height, and stay where they are.
  const eight = (note: boolean) => freeSlide({ title: 'Stations', blocks: [...Array.from({ length: 8 }, (_, index): FreeBlockContent => ({ chart: quarters([12 + index, 18, 23, 14]), caption: `Station ${index + 1}` })), ...(note ? [{ note: 'Every station peaked in the third quarter.' }] : [])] });
  assert.deepEqual([planOf(eight(true)).plan.layout.dense, planOf(eight(false)).plan.layout.dense], [true, false]);
  // A caption that PowerPoint sets on one line has a strip of one line, and the chart starts a fifth of its size
  // under it. Its slot holds it: it is not reported.
  const calls = freeSlide({ title: 'Calls', blocks: [{ chart: quarters([12, 18, 23, 14]), caption: 'Calls to the service line (thousands)' }, { chart: quarters([12, 18, 23, 14]), caption: 'Station 2' }, { chart: quarters([12, 18, 23, 14]), caption: 'Station 3' }, { chart: quarters([12, 18, 23, 14]), caption: 'Station 4' }], columns: 4 });
  const captioned = planOf(calls);
  const caption = captioned.shape('Block 1 caption');
  assert.ok(Math.abs(caption.box.h * H - 1.25 * (caption.size ?? 0)) < 0.1, `${caption.box.h * H}`);
  assert.ok(Math.abs((captioned.shape('Block 1 chart').box.y - caption.box.y) * H - 1.45 * (caption.size ?? 0)) < 0.1);
  assert.deepEqual(validateDeck(template, deckOf([calls])).filter((issue) => issue.slot === 'b1Caption'), []);
});
