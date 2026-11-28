// 0.9.0: a free slide whose blocks leave much of its room free sets its text larger, one note
// alone is a statement, and notes alone sit down their room.
import assert from 'node:assert/strict';
import { readFileSync } from 'node:fs';
import { dirname, join } from 'node:path';
import { test } from 'node:test';
import { fileURLToPath } from 'node:url';
import { freeContent, freeSlide } from '../src/free/blocks.ts';
import type { FreeBlockContent } from '../src/free/blocks.ts';
import { arrangeBlocks } from '../src/free/layout.ts';
import type { FreeLayout, LayoutFrame } from '../src/free/layout.ts';
import { freeFor } from '../src/free/resolve.ts';
import type { FreeResolved } from '../src/free/resolve.ts';
import { freeLine } from '../src/plan/free.ts';
import { loadTemplate, validateDeck } from '../src/index.ts';
import type { DeckSpec, SlideSpec, TemplateManifest } from '../src/index.ts';

const root = join(dirname(fileURLToPath(import.meta.url)), '..');
const manifest = JSON.parse(readFileSync(join(root, 'templates', 'demo.manifest.json'), 'utf8')) as TemplateManifest;
const template = loadTemplate(new Uint8Array(readFileSync(join(root, 'templates', 'demo.pptx'))), manifest);

const planOf = (slide: SlideSpec, deck: Pick<DeckSpec, 'textScale'> = {}): FreeResolved['plan'] => {
  const made = freeFor(template, deck, slide);
  assert.notEqual(typeof made, 'string', String(made));
  return (made as FreeResolved).plan;
};
const sizesOf = (plan: FreeResolved['plan']) => new Map(plan.shapes.filter((shape) => shape.kind === 'text').map((shape) => [shape.name, shape.size ?? 0]));
const usedOf = (plan: FreeResolved['plan']) => Math.max(...plan.layout.blocks.map((block) => block.box.y + block.box.h)) - Math.min(...plan.layout.blocks.map((block) => block.box.y));
const card = (heading: string, text: string): FreeBlockContent => ({ card: { heading, text } });
const three = [card('Longer hours', 'Bikes unlock from five.'), card('Safer docks', 'Every station gets a light.'), card('Group passes', 'One pass for five riders.')];
const kept = (slide: SlideSpec): SlideSpec => ({ ...slide, arrange: { ...slide.arrange, grow: false } });

test('a slide whose blocks leave much of its room free sets its text larger, and `grow: false` keeps it as it was', () => {
  const slide = freeSlide({ title: 'Three ways in', blocks: three });
  const [grown, plain] = [planOf(slide), planOf(kept(slide))];
  const [now, before] = [sizesOf(grown), sizesOf(plain)];
  // Every text is at least as large as it was, some are larger, and they grew together.
  const ratios = [...before].map(([name, size]) => (now.get(name) ?? 0) / size);
  assert.ok(Math.min(...ratios) >= 1 && Math.max(...ratios) > 1.1, ratios.join(' '));
  assert.ok(Math.max(...ratios) <= 1.3 * Math.min(...ratios) + 1e-6, 'no text grew far more than another');
  assert.ok(Math.max(...ratios) <= 1.5 + 1e-6, 'by no more than half');
  // The blocks are set out as they were, three on a line, and take more of the room.
  assert.equal(grown.layout.columns, plain.layout.columns);
  assert.equal(new Set(grown.layout.blocks.map((block) => block.box.y)).size, 1);
  assert.ok(usedOf(grown) > usedOf(plain));
  // The title stays the largest text of the slide.
  assert.ok(Math.max(...now.values()) < 34);
  // What is written still fits: the deck's check has nothing to say of either slide.
  assert.deepEqual(validateDeck(template, { version: 1, fields: { organizationDepartment: 'Riverbend' }, slides: [slide] }, { layout: 'adaptive' }), []);
  // The wish is the deck's to state, and it comes back from the slide's content.
  assert.equal(freeSlide({ blocks: three, grow: false }).arrange?.grow, false);
  assert.equal(freeContent(kept(slide)).grow, false);
  assert.equal(slide.arrange, undefined);
});

test('a slide that fills its room, a slide of running text and a slide with a chart keep the template\'s sizes', () => {
  const same = (slide: SlideSpec, why: string) => assert.deepEqual(JSON.parse(JSON.stringify(planOf(slide))), JSON.parse(JSON.stringify(planOf(kept(slide)))), why);
  // Full: nine cards of two lines each take most of the room.
  const long = 'Bikes unlock from five in the morning, for early shifts and the first trams.';
  same(freeSlide({ title: 'Nine ways in', blocks: Array.from({ length: 9 }, (_, index) => card(`Point ${index + 1}`, long)) }), 'a full slide');
  // Running text, lists and quotations alone hang from the top line, and a text alone has rules of its own.
  same(freeSlide({ title: 'Why', blocks: [{ text: 'A station within five minutes of every home.' }] }), 'a text alone');
  same(freeSlide({ title: 'Why', blocks: [{ list: ['One', 'Two', 'Three'] }] }), 'a list alone');
  same(freeSlide({ title: 'Said', blocks: [{ quote: 'I sold my second car.', source: 'Dana' }] }), 'a quotation alone');
  // A chart, a table and a picture take the room there is.
  const chart = { type: 'column' as const, categories: ['Spring', 'Summer'], series: [{ name: '2025', values: [7, 19] }] };
  same(freeSlide({ title: 'Rides', blocks: [{ chart }, card('Longer hours', 'Bikes unlock from five.'), card('Safer docks', 'Every station gets a light.')] }), 'a slide with a chart');
  // A slide that is planned for a writer is planned at the template's own sizes: what a writer is told always fits.
  const planned = (grow?: false) => freeFor(template, {}, { blocks: ['card', 'card', 'card'], slots: {}, ...(grow === false ? { arrange: { grow } } : {}) }, { icons: false });
  assert.deepEqual(JSON.parse(JSON.stringify(planned())), JSON.parse(JSON.stringify(planned(false))));
});

test('phases, milestones and figures grow where their width holds larger text, and a one-line heading stays one line', () => {
  const timeline = freeSlide({
    title: 'The road to 2027',
    blocks: [
      { milestone: { when: 'Q1 2026', heading: 'North Bridge', text: 'Forty docks beside the tram stop.' } },
      { milestone: { when: 'Q3 2026', heading: 'E-bikes arrive', text: 'A first fleet of 120.' } },
      { milestone: { when: 'Q1 2027', heading: 'One pass', text: 'Bikes and trams on one ticket.' } },
    ],
  });
  const [grown, plain] = [planOf(timeline), planOf(kept(timeline))];
  assert.ok((sizesOf(grown).get('Block 1 heading') ?? 0) > (sizesOf(plain).get('Block 1 heading') ?? 0));
  // No heading that took one line is given two: a line kept free under it would stay empty.
  for (const shape of grown.shapes) if (shape.kind === 'text' && / (heading|when)$/.test(shape.name)) assert.notEqual(shape.lines, 'several', shape.name);
  // Figures on a band of colour grow too, and the band with them.
  const figures = freeSlide({ title: 'A year', tone: 'accent', blocks: [{ figure: { value: '+34%', label: 'Monthly pass membership growth' } }, { figure: { value: '11', label: 'New stations opened citywide' } }, { figure: { value: '2 sec', label: 'Unlock time down from 9 seconds' } }] });
  const [band, small] = [planOf(figures), planOf(kept(figures))];
  assert.ok((sizesOf(band).get('Block 1 label') ?? 0) > (sizesOf(small).get('Block 1 label') ?? 0));
  const ground = (plan: FreeResolved['plan']) => plan.shapes.find((shape) => shape.name === 'Ground')?.box.h ?? 0;
  assert.ok(ground(band) > ground(small));
});

test('a deck that asks for larger text never gets smaller text than one that does not', () => {
  for (const blocks of [three, [card('Hours', 'From five.'), card('Docks', 'Lit.')], Array.from({ length: 5 }, (_, index) => card(`Point ${index + 1}`, 'Bikes unlock from five in the morning.'))]) {
    const slide = freeSlide({ title: 'Ways in', blocks });
    const without = sizesOf(planOf(slide));
    for (const textScale of [1.25, 1.5]) {
      for (const [name, size] of sizesOf(planOf(slide, { textScale }))) assert.ok(size >= (without.get(name) ?? 0) - 1e-6, `${blocks.length} cards at ${textScale}: ${name} is ${size} pt, and ${without.get(name)} pt without the wish`);
    }
  }
});

test('one note alone is a statement: set large, in up to four lines, and sitting in its room', () => {
  const statement = freeSlide({ title: 'Our commitment', blocks: [{ note: 'A station within five minutes of every home across Riverbend by the end of 2027.' }] });
  const plan = planOf(statement);
  const note = plan.shapes.find((shape) => shape.name === 'Block 1 note');
  // Far larger than a bottom line under other blocks, and no larger than seven tenths of the title.
  const under = planOf(kept(freeSlide({ title: 'Ways in', blocks: [...three, { note: 'A station within five minutes of every home across Riverbend by the end of 2027.' }] })));
  const bottomLine = under.shapes.find((shape) => / note$/.test(shape.name));
  assert.ok((note?.size ?? 0) >= 1.3 * (bottomLine?.size ?? 0), `${note?.size} and ${bottomLine?.size} pt`);
  assert.ok((note?.size ?? 0) <= 0.7 * 34 + 0.5);
  // It does not hang under the title: it sits down its room, with more of the empty room under it than over it.
  const block = plan.layout.blocks[0]?.box as { y: number; h: number };
  const two = planOf(freeSlide({ title: 'Ways in', blocks: [...three, { note: 'Short.' }] }));
  const roomTop = Math.min(...two.layout.blocks.map((entry) => entry.box.y));
  assert.ok(block.y > roomTop + 0.05, `at ${block.y}, the room starts near ${roomTop}`);
  assert.ok(block.y + block.h < 0.8);
  assert.deepEqual(validateDeck(template, { version: 1, fields: { organizationDepartment: 'Riverbend' }, slides: [statement] }, { layout: 'adaptive' }), []);
  // Several notes alone are bottom lines still, at a bottom line's size, and sit down their room together.
  const several = planOf(freeSlide({ title: 'Three things', blocks: [{ note: 'One.' }, { note: 'Two.' }, { note: 'Three.' }] }));
  assert.ok((several.shapes.find((shape) => shape.name === 'Block 1 note')?.size ?? 0) < (note?.size ?? 0));
  assert.ok((several.layout.blocks[0]?.box.y ?? 0) > roomTop + 0.02);
});

test('a planner is told that a row of small parts alone leaves the slide half empty', () => {
  const line = freeLine(['text', 'card', 'figure', 'phase', 'milestone', 'note'], { comfortable: 3 });
  assert.match(line, /A slide of one row of figures, phases or milestones alone leaves its lower half empty: give it a note under the row/);
  // Said only of the kinds on offer, and not at all where there is no note to give or no such part.
  assert.match(freeLine(['text', 'card', 'phase', 'note'], { comfortable: 3 }), /one row of phases alone/);
  assert.doesNotMatch(freeLine(['text', 'card', 'note'], { comfortable: 3 }), /lower half empty/);
  assert.doesNotMatch(freeLine(['text', 'figure', 'phase'], { comfortable: 3 }), /lower half empty/);
  assert.doesNotMatch(freeLine(['text', 'figure', 'note'], { closed: true }), /lower half empty/);
});

test('where the usual room ends well short of the page, a low row sits down the room to the page\'s foot', () => {
  // The room of a template's own layouts, and the larger one a full slide may take: from 0.3 to 0.95 of the slide.
  const usual: LayoutFrame = { x: 0.04, w: 0.92, top: 0.32, bottom: 0.88, denseTop: 0.3, denseBottom: 0.95, gap: 0.02, aspect: 16 / 9 };
  const row = (bottom: number, kinds: Array<'card' | 'note'> = ['card', 'card', 'card']): FreeLayout => arrangeBlocks(kinds, { frame: { ...usual, bottom }, rtl: false, needed: () => 0.06 });
  const first = (layout: FreeLayout) => layout.blocks[0]?.box as { y: number; h: number; w: number };
  // Two fifths of what a row leaves is above it. Where the usual room ends near the page's foot, that is counted in the usual room.
  const share = (layout: FreeLayout, bottom: number) => (first(layout).y - usual.top) / (bottom - usual.top - first(layout).h);
  assert.ok(Math.abs(share(row(0.88), 0.88) - 0.4) < 0.005, String(share(row(0.88), 0.88)));
  assert.ok(Math.abs(share(row(0.86), 0.86) - 0.4) < 0.005);
  // Where it ends a tenth of the slide or more above it (a template whose own layouts use the upper part of
  // the page), the row is settled in the room down to the page's foot: it stood over an empty lower third.
  for (const bottom of [0.85, 0.8, 0.77]) {
    const short = row(bottom);
    assert.ok(Math.abs(share(short, usual.denseBottom) - 0.4) < 0.005, `to ${bottom}: ${share(short, usual.denseBottom)}`);
    assert.ok(first(short).y > first(row(0.88)).y, 'lower than in a room that ends near the foot');
    // It is moved, not set out anew: the blocks have the size they had, and the slide is not one that took the larger room.
    assert.equal(short.dense, false);
    assert.equal(first(short).h, first(row(0.88)).h);
    assert.equal(first(short).w, first(row(0.88)).w);
    assert.ok(first(short).y + first(short).h <= usual.denseBottom);
  }
  // Notes alone sit down the same room.
  assert.ok(first(row(0.77, ['note'])).y > first(row(0.88, ['note'])).y);
});
