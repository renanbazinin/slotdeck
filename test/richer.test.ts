// 0.8.0: phrases set strong inside a text, three kinds of block that are drawn (a process, a
// timeline, a comparison), and a band of colour behind the blocks of a slide.
import assert from 'node:assert/strict';
import { readFileSync } from 'node:fs';
import { dirname, join } from 'node:path';
import { test } from 'node:test';
import { fileURLToPath } from 'node:url';
import { strFromU8, unzipSync } from 'fflate';
import { freeTone, markedParagraphs, plainText, toParagraphs } from '../src/deck/spec.ts';
import { describeBlocks, freeContent, freeSlide } from '../src/free/blocks.ts';
import { arrangeBlocks } from '../src/free/layout.ts';
import type { LayoutFrame } from '../src/free/layout.ts';
import { freeOffer } from '../src/free/offer.ts';
import { freeStrongColor } from '../src/free/plan.ts';
import { freeFor } from '../src/free/resolve.ts';
import { calmTones, drawnOnOneSlide, readShows } from '../src/plan/free.ts';
import { STRONG_RULE, systemPrompt } from '../src/plan/write.ts';
import { loadTemplate, planOutline, renderDeck, validateDeck } from '../src/index.ts';
import type { DeckSpec, FreeBlockKind, JsonRequest, LlmProvider, SlideSpec, TemplateManifest } from '../src/index.ts';

const root = join(dirname(fileURLToPath(import.meta.url)), '..');
const manifest = JSON.parse(readFileSync(join(root, 'templates', 'demo.manifest.json'), 'utf8')) as TemplateManifest;
const demoPptx = new Uint8Array(readFileSync(join(root, 'templates', 'demo.pptx')));
const template = loadTemplate(demoPptx, manifest);

const deckOf = (slides: SlideSpec[], more: Partial<DeckSpec> = {}): DeckSpec => ({ version: 1, fields: { organizationDepartment: 'Riverbend' }, ...more, slides });
const slideXml = (pptx: Uint8Array, number: number) => strFromU8(unzipSync(pptx)[`ppt/slides/slide${number}.xml`] as Uint8Array);
/** The shapes of a slide by name, each with its XML. */
const shapesOf = (xml: string): Map<string, string> => new Map([...xml.matchAll(/<p:(sp|pic|graphicFrame)\b.*?<\/p:\1>/gs)].map((match) => [/<p:cNvPr id="\d+" name="([^"]*)"/.exec(match[0])?.[1] ?? '', match[0]]));
/** The runs of a shape: their text, whether they are bold, and their colour where they state one. */
const runsOf = (shape: string) =>
  [...shape.matchAll(/<a:r>(.*?)<\/a:r>/gs)].map((match) => ({
    text: /<a:t>([^<]*)<\/a:t>/.exec(match[1] as string)?.[1] ?? '',
    bold: /<a:rPr[^>]*\bb="1"/.test(match[1] as string),
    color: /<a:rPr[^>]*>.*?<a:solidFill><a:srgbClr val="([0-9A-F]{6})"/s.exec(match[1] as string)?.[1],
  }));
const offOf = (shape: string) => ({ x: Number(/<a:off x="(-?\d+)"/.exec(shape)?.[1]), y: Number(/<a:off x="-?\d+" y="(-?\d+)"/.exec(shape)?.[1]), cx: Number(/<a:ext cx="(\d+)"/.exec(shape)?.[1]), cy: Number(/<a:ext cx="\d+" cy="(\d+)"/.exec(shape)?.[1]) });

// ---- Phrases set strong ---------------------------------------------------------------------------

test('a phrase between two pairs of asterisks is taken out of its marks, and counted as the slide shows it', () => {
  assert.deepEqual(markedParagraphs('**Faster:** passes grew **34%** this year'), [{ text: 'Faster: passes grew 34% this year', strong: [[0, 7], [20, 23]] }]);
  assert.deepEqual(toParagraphs('**Faster:** passes grew **34%** this year'), ['Faster: passes grew 34% this year']);
  // Paragraph by paragraph, in a string of several lines and in a list of strings.
  assert.deepEqual(markedParagraphs('One **two**\nthree').map((paragraph) => paragraph.strong), [[[4, 7]], []]);
  assert.deepEqual(markedParagraphs(['**a** b', 'c **d e**']).map((paragraph) => paragraph.text), ['a b', 'c d e']);
  // Asterisks that are no such pair stay as they are written: a footnote mark, a product of two numbers, marks with a space inside.
  for (const plain of ['Price** and cost', '2 ** 3 = 8', '** not this **', 'a * b', '****']) assert.deepEqual(markedParagraphs(plain), [{ text: plain, strong: [] }], plain);
  // A cell of a table, a label of a chart and the text read out for a picture have no bold: the marks go, the words stay.
  assert.equal(plainText('**34%** more'), '34% more');
});

test('a strong phrase is a bold run of its own, and the paragraph is otherwise as it was', () => {
  const plain = renderDeck(template, deckOf([freeSlide({ title: 'This year', blocks: [{ list: ['Faster unlock: two seconds', 'Passes grew 34% this year'] }] })]), { layout: 'adaptive' });
  const strong = renderDeck(template, deckOf([freeSlide({ title: 'This year', blocks: [{ list: ['**Faster unlock:** two seconds', 'Passes grew **34%** this year'] }] })]), { layout: 'adaptive' });
  const list = shapesOf(slideXml(strong.pptx, 1)).get('Block 1 list') as string;
  const runs = runsOf(list);
  assert.deepEqual(runs.map((run) => [run.text, run.bold]), [['Faster unlock:', true], [' two seconds', false], ['Passes grew ', false], ['34%', true], [' this year', false]]);
  // On a free slide a strong phrase with a number in it takes the template's accent; one of words alone is bold and no more.
  const accent = freeStrongColor((freeFor(template, {}, { blocks: ['list'], slots: {} }) as Exclude<ReturnType<typeof freeFor>, string>).kit);
  assert.equal(accent, '0E7C7B');
  assert.equal(runs[3]?.color, accent);
  assert.notEqual(runs[0]?.color, accent);
  assert.equal(runs[1]?.color, runs[0]?.color, 'the words of a phrase keep the colour of their text');
  // Nothing but the runs differs: the same boxes, the same sizes, the same text.
  const texts = (xml: string) => [...xml.matchAll(/<a:t>([^<]*)<\/a:t>/g)].map((match) => match[1]).join('');
  assert.equal(texts(slideXml(strong.pptx, 1)), texts(slideXml(plain.pptx, 1)));
  assert.deepEqual([...shapesOf(slideXml(strong.pptx, 1))].map(([name, xml]) => [name, offOf(xml)]), [...shapesOf(slideXml(plain.pptx, 1))].map(([name, xml]) => [name, offOf(xml)]));
  assert.deepEqual(validateDeck(template, deckOf([freeSlide({ title: 'This year', blocks: [{ list: ['**Faster unlock:** two seconds'] }] })]), { layout: 'adaptive' }), []);
});

test('a strong phrase is set on a slide of a layout too, in bold alone, and in a right-to-left paragraph', () => {
  const deck = deckOf([{ pattern: 'text-only', slots: { title: 'Notes', body: ['**Faster unlock:** the app opens a dock in two seconds.', 'זמן ההמתנה ירד ל-**4 דקות** בממוצע.'] } }]);
  assert.deepEqual(validateDeck(template, deck, { layout: 'adaptive' }).filter((issue) => issue.severity === 'error'), []);
  const xml = slideXml(renderDeck(template, deck, { layout: 'adaptive' }).pptx, 1);
  const body = [...shapesOf(xml).values()].find((shape) => shape.includes('Faster unlock:')) as string;
  const runs = runsOf(body);
  assert.deepEqual(runs.filter((run) => run.bold).map((run) => run.text), ['Faster unlock:', '4 דקות']);
  // No colour is given on a layout: what the text stands on there is the template's to know.
  assert.equal(new Set(runs.map((run) => run.color)).size, 1);
  // Every run of the Hebrew paragraph is marked Hebrew, the strong one too: PowerPoint orders words by the language of their run.
  const hebrew = [...body.matchAll(/<a:p>.*?<\/a:p>/gs)].map((match) => match[0]).find((paragraph) => paragraph.includes('דקות')) as string;
  assert.equal([...hebrew.matchAll(/<a:rPr[^>]*lang="he-IL"/g)].length, [...hebrew.matchAll(/<a:r>/g)].length);
  assert.match(hebrew, /<a:pPr[^>]*rtl="1"/);
});

test('a writer is told of the one mark it may use, unless an app turns it off', () => {
  assert.ok(systemPrompt('en', 'brief').includes(STRONG_RULE));
  assert.ok(systemPrompt('he', 'deck', true, 'read').includes(STRONG_RULE));
  assert.ok(!systemPrompt('en', 'brief', false, 'present', false).includes('asterisk'));
  assert.equal(systemPrompt('en', 'brief', false, 'present', false), systemPrompt('en', 'brief').replace(`${STRONG_RULE}\n`, ''));
});

// ---- A process, a timeline, a comparison -----------------------------------------------------------

const process = freeSlide({
  title: 'How a station opens',
  blocks: [
    { phase: { heading: 'Survey', text: 'Count riders for two weeks.' } },
    { phase: { heading: 'Permit', text: 'The city signs off.' } },
    { phase: { heading: 'Build', text: 'Docks go in.' } },
    { phase: { heading: 'Launch' } },
  ],
});
const timeline = freeSlide({
  title: 'The road to 2027',
  blocks: [
    { milestone: { when: 'Q1 2026', heading: 'North Bridge opens', text: 'Forty docks.' } },
    { milestone: { when: 'Q3 2026', heading: 'E-bikes arrive', text: 'A fleet of 120.' } },
    { milestone: { when: 'Q1 2027', heading: 'One pass', text: 'Bikes and trams.' } },
  ],
});
const comparison = freeSlide({
  title: 'Three ways to grow',
  blocks: [
    { option: { heading: 'More docks', points: ['Cheapest', 'No new riders'] } },
    { option: { heading: 'E-bikes', badge: 'Recommended', points: ['Opens the hills', 'Needs charging'] } },
    { option: { heading: 'New district', points: ['20,000 homes', 'Slow permits'] } },
  ],
});

test('the new kinds are written and read back block by block, and named in a sentence', () => {
  assert.deepEqual(process.blocks, ['phase', 'phase', 'phase', 'phase']);
  assert.deepEqual(process.slots, { title: 'How a station opens', b1Heading: 'Survey', b1Text: 'Count riders for two weeks.', b2Heading: 'Permit', b2Text: 'The city signs off.', b3Heading: 'Build', b3Text: 'Docks go in.', b4Heading: 'Launch' });
  assert.deepEqual(timeline.slots.b2When, 'Q3 2026');
  assert.deepEqual(comparison.slots.b2Badge, 'Recommended');
  assert.deepEqual(comparison.slots.b1Items, ['Cheapest', 'No new riders']);
  for (const slide of [process, timeline, comparison]) assert.deepEqual(freeSlide(freeContent(slide)), slide);
  assert.equal(describeBlocks(['phase', 'phase', 'milestone', 'option', 'option', 'option']), '2 phases, a milestone and 3 options');
  // They are offered where cards are: in every deck that is not text only.
  assert.deepEqual(freeOffer(template, { figures: false, photos: false, graphics: false, icons: false })?.kinds.filter((kind) => ['phase', 'milestone', 'option'].includes(kind)), ['phase', 'milestone', 'option']);
  assert.equal(drawnOnOneSlide(['text', 'phase', 'option']), 'a process as a chain of arrows or a comparison of options');
  assert.equal(drawnOnOneSlide(['text', 'list']), '');
});

test('up to six phases or milestones stand on one line, and four options, where each still has its least width', () => {
  const frame: LayoutFrame = { x: 0.04, w: 0.92, top: 0.3, bottom: 0.9, denseTop: 0.25, denseBottom: 0.93, gap: 0.02, aspect: 16 / 9 };
  const linesOf = (kinds: FreeBlockKind[]) => new Set(arrangeBlocks(kinds, { frame, rtl: false }).blocks.map((block) => block.box.y)).size;
  for (const kind of ['phase', 'milestone'] as const) {
    for (let count = 2; count <= 6; count++) assert.equal(linesOf(Array<FreeBlockKind>(count).fill(kind)), 1, `${count} ${kind}s`);
    assert.ok(linesOf(Array<FreeBlockKind>(9).fill(kind)) > 1, `nine ${kind}s`);
  }
  for (let count = 2; count <= 4; count++) assert.equal(linesOf(Array<FreeBlockKind>(count).fill('option')), 1, `${count} options`);
  // In a room too narrow for them on one line they are set out as other blocks are.
  assert.ok(new Set(arrangeBlocks(Array<FreeBlockKind>(6).fill('phase'), { frame: { ...frame, w: 0.5 }, rtl: false }).blocks.map((block) => block.box.y)).size > 1);
});

test('a process is a chain of arrows that point where reading goes, each with its name on it', () => {
  assert.deepEqual(validateDeck(template, deckOf([process]), { layout: 'adaptive' }), []);
  const shapes = shapesOf(slideXml(renderDeck(template, deckOf([process]), { layout: 'adaptive' }).pptx, 1));
  const arrows = [1, 2, 3, 4].map((n) => shapes.get(`Block ${n} arrow`) as string);
  // The first starts flat; the others take the point of the arrow before them.
  assert.deepEqual(arrows.map((arrow) => /prst="(\w+)"/.exec(arrow)?.[1]), ['homePlate', 'chevron', 'chevron', 'chevron']);
  for (const arrow of arrows) assert.match(arrow, /<a:solidFill><a:srgbClr val="0E7C7B"\/><\/a:solidFill><a:ln><a:noFill\/><\/a:ln>/);
  // Each reaches into the gap towards the next, and does not lie on it; the last ends with its block.
  const places = arrows.map(offOf);
  for (let at = 0; at < 3; at++) {
    const next = places[at + 1] as ReturnType<typeof offOf>;
    const here = places[at] as ReturnType<typeof offOf>;
    assert.ok(here.x + here.cx > here.x + (places[3] as ReturnType<typeof offOf>).cx, 'longer than the last, which has no arrow to point into');
    assert.ok(here.x + here.cx <= next.x, 'and not over the next');
  }
  // The name stands on the arrow, in the middle, in a colour that reads on it.
  const name = shapes.get('Block 2 heading') as string;
  assert.match(name, /algn="ctr"/);
  assert.match(name, /anchor="ctr"/);
  assert.equal(runsOf(name)[0]?.color, 'FFFFFF');
  // A phase without a text has none drawn.
  assert.ok(shapes.has('Block 3 text'));
  assert.ok(!shapes.has('Block 4 text'));

  // Read from the right, the arrows are mirrored and the first stands at the right.
  const rtl = shapesOf(slideXml(renderDeck(template, deckOf([process], { direction: 'rtl' }), { layout: 'adaptive' }).pptx, 1));
  for (const n of [1, 2, 3, 4]) assert.match(rtl.get(`Block ${n} arrow`) as string, /<a:xfrm flipH="1">/);
  assert.ok(offOf(rtl.get('Block 1 arrow') as string).x > offOf(rtl.get('Block 4 arrow') as string).x);
});

test('a timeline is a line with a dot for each milestone, its date over the line and what happened under it', () => {
  assert.deepEqual(validateDeck(template, deckOf([timeline]), { layout: 'adaptive' }), []);
  const xml = slideXml(renderDeck(template, deckOf([timeline]), { layout: 'adaptive' }).pptx, 1);
  const shapes = shapesOf(xml);
  const line = offOf(shapes.get('Timeline 1') as string);
  const dots = [1, 2, 3].map((n) => offOf(shapes.get(`Block ${n} dot`) as string));
  // One line from the first milestone to the end of the last, behind the dots, through the middle of each.
  assert.ok(xml.indexOf('name="Timeline 1"') < xml.indexOf('name="Block 1 dot"'));
  for (const dot of dots) {
    assert.match(String(Math.abs(dot.y + dot.cy / 2 - (line.y + line.cy / 2)) < 12700), /true/, 'a dot sits on the line');
    assert.ok(Math.abs(dot.cx - dot.cy) < 1270, 'a dot is round, to a tenth of a point');
  }
  assert.ok(line.x <= (dots[0] as (typeof dots)[number]).x && line.x + line.cx > (dots[2] as (typeof dots)[number]).x + (dots[2] as (typeof dots)[number]).cx);
  // The date is over the line, bold, in the colour of the dots; the heading under it.
  const when = shapes.get('Block 1 when') as string;
  assert.ok(offOf(when).y < line.y && offOf(shapes.get('Block 1 heading') as string).y > line.y);
  assert.deepEqual(runsOf(when).map((run) => [run.text, run.bold, run.color]), [['Q1 2026', true, '0E7C7B']]);
});

test('a comparison picks out the option with a badge: its ground in the template\'s colour, its text and bullets in white', () => {
  assert.deepEqual(validateDeck(template, deckOf([comparison]), { layout: 'adaptive' }), []);
  const shapes = shapesOf(slideXml(renderDeck(template, deckOf([comparison]), { layout: 'adaptive' }).pptx, 1));
  const fillOf = (shape: string) => /<p:spPr>.*?<a:(?:prst|cust)Geom.*?<\/a:(?:prst|cust)Geom>(?:<a:solidFill><a:srgbClr val="([0-9A-F]{6})")?/s.exec(shape)?.[1];
  assert.equal(fillOf(shapes.get('Block 2 panel') as string), '0E7C7B');
  assert.notEqual(fillOf(shapes.get('Block 1 panel') as string), '0E7C7B');
  // The three stand side by side, equally high.
  const panels = [1, 2, 3].map((n) => offOf(shapes.get(`Block ${n} panel`) as string));
  assert.equal(new Set(panels.map((panel) => `${panel.y}/${panel.cy}`)).size, 1);
  // On the option that is picked out every text is white and its bullets follow the text; elsewhere they are the template's.
  for (const part of ['badge', 'heading', 'list']) assert.equal(runsOf(shapes.get(`Block 2 ${part}`) as string)[0]?.color, 'FFFFFF', part);
  assert.match(shapes.get('Block 2 list') as string, /<a:buClrTx\/>/);
  assert.doesNotMatch(shapes.get('Block 1 list') as string, /<a:buClrTx\/>/);
  assert.notEqual(runsOf(shapes.get('Block 1 heading') as string)[0]?.color, 'FFFFFF');
  // Only the option with a badge has one drawn; the others keep its line free, so that the names start on one line.
  assert.ok(!shapes.has('Block 1 badge') && !shapes.has('Block 3 badge'));
  assert.equal(offOf(shapes.get('Block 1 heading') as string).y, offOf(shapes.get('Block 2 heading') as string).y);

  // Where every option has a badge none is picked out, and without any badge no line is kept for one.
  const all = freeSlide({ blocks: [{ option: { heading: 'A', badge: 'Fast', points: ['One'] } }, { option: { heading: 'B', badge: 'Cheap', points: ['Two'] } }] });
  const even = shapesOf(slideXml(renderDeck(template, deckOf([all]), { layout: 'adaptive' }).pptx, 1));
  assert.notEqual(fillOf(even.get('Block 1 panel') as string), '0E7C7B');
  assert.notEqual(fillOf(even.get('Block 2 panel') as string), '0E7C7B');
});

test('a writer is told what each part of the new kinds is for, and how much its box holds', () => {
  const planned = freeFor(template, {}, { blocks: ['phase', 'phase', 'phase', 'milestone', 'milestone', 'option', 'option'], slots: {} }, { icons: false });
  assert.ok(typeof planned !== 'string');
  if (typeof planned === 'string') return;
  const slots = new Map(planned.plan.pattern.slots.map((slot) => [slot.name, slot]));
  // Every part is planned in, the optional ones too: room is kept for a text under a phase and for the badge of an option.
  for (const name of ['b1Heading', 'b1Text', 'b4When', 'b4Heading', 'b4Text', 'b6Badge', 'b6Heading', 'b6Items']) assert.ok(slots.has(name), name);
  for (const name of ['b1Text', 'b4Text', 'b6Badge']) assert.equal(slots.get(name)?.empty, 'remove', `${name} may be left out`);
  for (const name of ['b1Heading', 'b4When', 'b4Heading', 'b6Heading', 'b6Items']) assert.equal(slots.get(name)?.empty, undefined, `${name} is required`);
  for (const slot of slots.values()) if (slot.kind === 'text' && slot.name !== 'title') assert.ok((slot.maxChars ?? 0) >= 8, `${slot.name} holds ${slot.maxChars}`);
  assert.match(slots.get('b6Badge')?.description ?? '', /the one option that is recommended/);
  assert.match(slots.get('b4When')?.description ?? '', /a date, a quarter, a year/);
});

// ---- A band of colour -------------------------------------------------------------------------------

const figures = (tone?: 'accent' | 'tint') =>
  freeSlide({
    title: 'A year in three numbers',
    ...(tone ? { tone } : {}),
    blocks: [{ figure: { value: '34%', label: 'more passes' } }, { figure: { value: '11', label: 'new stations' } }, { figure: { value: '2 s', label: 'to unlock' } }],
  });

test('a slide may set its blocks on a band of the template\'s colour, their text in a colour that reads on it', () => {
  assert.equal(figures('accent').arrange?.tone, 'accent');
  assert.deepEqual(freeSlide(freeContent(figures('accent'))), figures('accent'));
  assert.deepEqual(validateDeck(template, deckOf([figures('accent')]), { layout: 'adaptive' }), []);
  const xml = slideXml(renderDeck(template, deckOf([figures('accent')]), { layout: 'adaptive' }).pptx, 1);
  const shapes = shapesOf(xml);
  const ground = shapes.get('Ground') as string;
  assert.match(ground, /<a:solidFill><a:srgbClr val="0E7C7B"\/><\/a:solidFill>/);
  // It lies behind every block, is as wide as the template's content and holds all that is drawn on it.
  const band = offOf(ground);
  const blocks = [...shapes].filter(([name]) => name.startsWith('Block '));
  assert.ok(blocks.length >= 6);
  for (const [name, shape] of blocks) {
    assert.ok(xml.indexOf('name="Ground"') < xml.indexOf(`name="${name}"`), `${name} is drawn over the band`);
    const at = offOf(shape);
    assert.ok(at.x >= band.x && at.x + at.cx <= band.x + band.cx && at.y >= band.y && at.y + at.cy <= band.y + band.cy, `${name} stands inside the band`);
    assert.equal(runsOf(shape)[0]?.color, 'FFFFFF', `${name} is white`);
  }
  // No panel of the template is drawn on it: a panel was made for the slide's own background.
  assert.ok(![...shapes.keys()].some((name) => / panel$/.test(name)));

  // A tint keeps the text as it is and sets a pale band behind it.
  const pale = shapesOf(slideXml(renderDeck(template, deckOf([figures('tint')]), { layout: 'adaptive' }).pptx, 1));
  const fill = /<a:solidFill><a:srgbClr val="([0-9A-F]{6})"/.exec(pale.get('Ground') as string)?.[1] as string;
  assert.notEqual(fill, '0E7C7B');
  assert.ok(parseInt(fill.slice(0, 2), 16) > 200, 'a light colour');
  assert.notEqual(runsOf(pale.get('Block 1 value') as string)[0]?.color, 'FFFFFF');
});

test('on a band in the main colour a card has no icon and no panel, and a strong figure no colour of its own', () => {
  const icon = new Uint8Array(readFileSync(join(root, 'examples', 'icons', 'bolt.png')));
  const cards = (tone?: 'accent' | 'tint') =>
    freeSlide({ title: 'Why', ...(tone ? { tone } : {}), blocks: [{ card: { heading: 'Fast', text: 'Unlocks in **2 seconds**.', icon: 'bolt' } }, { card: { heading: 'Near', text: 'A station close by.', icon: 'bolt' } }] });
  const shapesFor = (tone?: 'accent' | 'tint') => shapesOf(slideXml(renderDeck(template, deckOf([cards(tone)]), { layout: 'adaptive', assets: { bolt: icon } }).pptx, 1));
  // On the slide itself: an icon, a panel, and the strong figure in the template's accent.
  const plain = shapesFor();
  assert.ok(plain.has('Block 1 icon') && plain.has('Block 1 panel'));
  assert.equal(runsOf(plain.get('Block 1 text') as string).find((run) => run.bold)?.color, '0E7C7B');
  // On the band: neither, and the figure is bold in the one colour every text has there.
  const band = shapesFor('accent');
  assert.ok(!band.has('Block 1 icon') && !band.has('Block 1 panel') && band.has('Ground'));
  assert.deepEqual(runsOf(band.get('Block 1 text') as string).map((run) => [run.text, run.bold, run.color]), [['Unlocks in ', false, 'FFFFFF'], ['2 seconds', true, 'FFFFFF'], ['.', false, 'FFFFFF']]);
  // A pale band keeps the icon: it is a light ground.
  assert.ok(shapesFor('tint').has('Block 1 icon'));
  assert.deepEqual(validateDeck(template, deckOf([cards('accent')]), { layout: 'adaptive', assets: ['bolt'] }).filter((issue) => issue.severity === 'error'), []);
});

test('without a band a slide is what it was, and a slide a band cannot carry is set out as without one', () => {
  const plain = renderDeck(template, deckOf([figures()]), { layout: 'adaptive' });
  assert.ok(!slideXml(plain.pptx, 1).includes('name="Ground"'));
  assert.equal(freeTone({ blocks: ['figure', 'note'], arrange: { tone: 'accent' } }), 'accent');
  assert.equal(freeTone({ blocks: ['figure', 'note'], arrange: {} }), undefined);
  assert.equal(freeTone({ blocks: ['figure'], arrange: { tone: 'purple' as 'accent' } }), undefined, 'a band that is none');
  for (const kind of ['chart', 'table', 'photo', 'graphic', 'option'] as const) assert.equal(freeTone({ blocks: ['figure', kind], arrange: { tone: 'accent' } }), undefined, kind);
  const chart = { type: 'column' as const, categories: ['A', 'B'], series: [{ name: 'x', values: [1, 2] }] };
  const withChart = (tone?: 'accent') => freeSlide({ title: 'Sales', ...(tone ? { tone } : {}), blocks: [{ chart }, { note: 'Up.' }] });
  const drawn = (slide: SlideSpec) => [...shapesOf(slideXml(renderDeck(template, deckOf([slide]), { layout: 'adaptive' }).pptx, 1))].map(([name, xml]) => [name, offOf(xml)]);
  assert.deepEqual(drawn(withChart('accent')), drawn(withChart()));
});

test('a plan may set a slide on a band, and code keeps such slides few', async () => {
  // What an answer says is read where every part can stand on a band, and passed over elsewhere.
  const reading = { kinds: ['text', 'figure', 'quote', 'chart', 'note'] as FreeBlockKind[], photos: 6 };
  assert.deepEqual(readShows({ shows: [{ kind: 'figure', count: 3 }], shows_main: 'auto', side_by_side: 0, shows_on: 'accent' }, 1, 'Numbers', reading, false), { blocks: ['figure', 'figure', 'figure'], arrange: { tone: 'accent' } });
  assert.deepEqual(readShows({ shows: [{ kind: 'chart', count: 2 }], shows_main: 'auto', side_by_side: 0, shows_on: 'accent' }, 1, 'Sales', reading, false), { blocks: ['chart', 'chart'] });
  assert.deepEqual(readShows({ shows: [{ kind: 'quote', count: 1 }], shows_main: 'auto', side_by_side: 0, shows_on: 'plain' }, 1, 'Said', reading, false), { blocks: ['quote'] });
  // A change to a plan that says nothing of it keeps the band the slide had.
  const old = [{ pattern: 'free', purpose: 'Numbers', blocks: ['figure', 'figure'] as FreeBlockKind[], arrange: { tone: 'accent' as const } }];
  assert.deepEqual(readShows({ shows: [{ kind: 'figure', count: 3 }], shows_main: 'auto', side_by_side: 0 }, 1, 'Numbers', { ...reading, old }, false).arrange, { tone: 'accent' });

  // Never two in a row, and no more than one slide in four.
  const entries: Array<{ arrange?: { tone?: 'accent' | 'tint'; columns?: number } }> = [{ arrange: { tone: 'accent' } }, { arrange: { tone: 'accent', columns: 2 } }, {}, { arrange: { tone: 'tint' } }, {}, {}, { arrange: { tone: 'accent' } }, {}];
  calmTones(entries);
  assert.deepEqual(entries, [{ arrange: { tone: 'accent' } }, { arrange: { columns: 2 } }, {}, { arrange: { tone: 'tint' } }, {}, {}, {}, {}]);

  // In a plan: the model's answer gives three slides a band, and the plan keeps the first.
  const requests: JsonRequest[] = [];
  const free = (purpose: string, kind: string, count: number) => ({ pattern: 'free', purpose, shows: [{ kind, count }], shows_main: 'auto', side_by_side: 0, shows_on: 'accent' });
  const llm: LlmProvider = {
    async generateJson(request) {
      requests.push(request);
      return { language: 'en', title: 'Riverbend', slides: [free('Numbers', 'figure', 3), free('Said', 'quote', 1), free('Plan', 'phase', 3)] } as never;
    },
  };
  const plan = await planOutline(template, { prompt: 'Riverbend grew 34% with 11 new stations', slideCount: 3, fields: { organizationDepartment: 'Riverbend' } }, { llm });
  assert.deepEqual(plan.slides.map((slide) => slide.arrange?.tone), ['accent', undefined, undefined]);
  assert.match(JSON.stringify(requests[0]?.schema), /"shows_on"/);
});
