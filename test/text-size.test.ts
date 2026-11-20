// How large the running text of a free slide is set (0.6.7): a template that states how wide its
// text is set (`characterWidth`), the least size its running text takes (`leastTextSize`), the
// headings of a person's document counted as they are set, and a long text set dense that takes
// three columns where two do not hold it. No test calls a model.
import assert from 'node:assert/strict';
import { readFileSync } from 'node:fs';
import { dirname, join } from 'node:path';
import { test } from 'node:test';
import { fileURLToPath } from 'node:url';
import { strFromU8, unzipSync } from 'fflate';
import * as fontkit from 'fontkit';
import { changeManifest, checkManifest, editManifest, fontkitMetrics, loadTemplate, renderDeck, textMeasurer, validateDeck } from '../src/index.ts';
import type { JsonRequest, LlmProvider } from '../src/index.ts';
import type { DeckSpec, SlideSpec, TemplateManifest } from '../src/index.ts';
import { headingPlaces, headingsOver } from '../src/deck/headings.ts';
import { freeKit } from '../src/free/kit.ts';
import { freeTextRoom } from '../src/free/offer.ts';
import { freeFor } from '../src/free/resolve.ts';
import type { FreeResolved } from '../src/free/resolve.ts';
import { documentHeadingPlaces, fillParagraphs, readDocument, settleDocumentSlide, shapeOf } from '../src/plan/document.ts';
import { headingRoom, runsIn } from '../src/render/text.ts';
import { openFonts } from './fonts.ts';

const root = join(dirname(fileURLToPath(import.meta.url)), '..');
const demoManifest = JSON.parse(readFileSync(join(root, 'templates', 'demo.manifest.json'), 'utf8')) as TemplateManifest;
const demoPptx = new Uint8Array(readFileSync(join(root, 'templates', 'demo.pptx')));
const demo = loadTemplate(demoPptx, demoManifest);
/** The demo template with settings of its own. */
const withSettings = (settings: Partial<TemplateManifest>) => loadTemplate(demoPptx, { ...demoManifest, ...settings });
const deckOf = (slides: SlideSpec[]): DeckSpec => ({ version: 1, language: 'en', fields: { organizationDepartment: 'Retail' }, slides });

const SENTENCE = 'Riders rely on the service every day, and the city gains from it in ways that are easy to see. ';
/** A text of so many characters in so many paragraphs, as alike as they can be. */
const textOf = (chars: number, count: number): string[] =>
  Array.from({ length: count }, (_, at) => {
    const length = Math.floor(chars / count) + (at < chars % count ? 1 : 0);
    return SENTENCE.repeat(Math.ceil(length / SENTENCE.length) + 1).slice(0, length - 1) + '.';
  });
const slideOf = (text: string[], dense = false, more: Partial<SlideSpec> = {}): SlideSpec => ({ pattern: 'free', blocks: ['text'], ...(dense ? { arrange: { dense: true } } : {}), slots: { title: 'A title', b1Text: text }, ...more });
const resolved = (value: FreeResolved | string): FreeResolved => {
  assert.notEqual(typeof value, 'string', String(value));
  return value as FreeResolved;
};
/** The size, the columns and the budget of the text of a free slide of one text, and the size its slide is told it reads at. */
const textOf1 = (template: typeof demo, slide: Pick<SlideSpec, 'blocks' | 'arrange' | 'slots' | 'headings'>) => {
  const { plan } = resolved(freeFor(template, {}, slide));
  const slot = plan.pattern.slots.find((entry) => entry.name === 'b1Text');
  const shape = plan.shapes.find((entry) => entry.id === slot?.shapeId);
  return { size: shape?.size, columns: shape?.columns ?? 1, most: slot?.kind === 'text' ? slot.maxChars : undefined, comfortable: plan.comfortable };
};

test('a template that states how wide its text is set has its running text set larger; cards, and a template without it, as before', () => {
  const narrow = withSettings({ characterWidth: 0.44 });
  // Running text of six paragraphs, and a long text set dense.
  const text = slideOf(textOf(1400, 6));
  assert.deepEqual([textOf1(demo, text).size, textOf1(narrow, text).size], [14.5, 16]);
  const dense = slideOf(textOf(2000, 8), true);
  assert.deepEqual([textOf1(demo, dense).size, textOf1(narrow, dense).size], [13.5, 14.5]);
  // The text set larger still fits as the deck's check counts it.
  for (const slide of [text, dense]) assert.deepEqual(validateDeck(narrow, deckOf([slide]), { layout: 'adaptive' }), []);
  // Cards keep their own count.
  const cards: SlideSpec = { pattern: 'free', blocks: ['card', 'card', 'card'], slots: { title: 'A title', b1Heading: 'One', b1Text: SENTENCE.slice(0, 90), b2Heading: 'Two', b2Text: SENTENCE.slice(0, 90), b3Heading: 'Three', b3Text: SENTENCE.slice(0, 90) } };
  const sizes = (template: typeof demo) => resolved(freeFor(template, {}, cards)).plan.shapes.filter((shape) => shape.kind === 'text').map((shape) => shape.size);
  assert.deepEqual(sizes(narrow), sizes(demo));
  // The kit carries the setting; a template without it has none.
  assert.equal(freeKit(narrow)?.characterWidth, 0.44);
  assert.equal(freeKit(demo)?.characterWidth, undefined);
});

test('a characterWidth or leastTextSize that is wrong is reported and not used', () => {
  for (const [settings, said] of [
    [{ characterWidth: 0.2 }, "the template's characterWidth is 0.2; it is how wide a character of its text is set, as a share of its size, between 0.3 and 0.8, such as 0.43; it is not used"],
    [{ characterWidth: '0.43' }, `the template's characterWidth is "0.43"; it is how wide a character of its text is set, as a share of its size, between 0.3 and 0.8, such as 0.43; it is not used`],
    [{ leastTextSize: 40 }, "the template's leastTextSize is 40; it is a size in points between 5 and 24, such as 10; it is not used"],
  ] as const) {
    const template = withSettings(settings as Partial<TemplateManifest>);
    assert.deepEqual(
      checkManifest(template).filter((issue) => issue.message.includes('characterWidth') || issue.message.includes('leastTextSize')),
      [{ severity: 'warning', message: said }],
    );
    const slide = slideOf(textOf(3000, 12), true);
    assert.deepEqual(textOf1(template, slide), textOf1(demo, slide));
  }
  // null is a setting that is not given.
  assert.deepEqual(checkManifest(withSettings({ characterWidth: null, leastTextSize: null } as unknown as Partial<TemplateManifest>)), checkManifest(demo));
});

test('changeManifest sets and clears characterWidth and leastTextSize, and refuses one out of range', () => {
  const set = changeManifest(demo, [{ op: 'setManifest', characterWidth: 0.43, leastTextSize: 10 }]);
  assert.deepEqual(set.rejected, []);
  assert.deepEqual([set.manifest.characterWidth, set.manifest.leastTextSize], [0.43, 10]);
  assert.match(set.applied[0]?.message ?? '', /counts a character of its running text as 0.43 of its size wide/);
  const cleared = changeManifest(loadTemplate(demoPptx, set.manifest), [{ op: 'setManifest', characterWidth: null, leastTextSize: null }]);
  assert.deepEqual(['characterWidth' in cleared.manifest, 'leastTextSize' in cleared.manifest], [false, false]);
  assert.match(changeManifest(demo, [{ op: 'setManifest', characterWidth: 0.1 }]).rejected[0]?.reason ?? '', /a character width is a share of the size between 0\.3 and 0\.8/);
  assert.match(changeManifest(demo, [{ op: 'setManifest', leastTextSize: 2 }]).rejected[0]?.reason ?? '', /the least text size is a size in points between 5 and 24/);
});

test('leastTextSize: a long text set dense is set no smaller, also where it fits the first room, never over the template’s size, and its budget is what that size holds', () => {
  const twelve = withSettings({ leastTextSize: 12 });
  const long = slideOf(textOf(4000, 12), true);
  assert.deepEqual(textOf1(demo, long), { size: 10, columns: 3, most: 4555, comfortable: 9 });
  const set = textOf1(twelve, long);
  // Set at it; a slide is told its text is small under 9 pt as before, the size its captions and labels may take.
  assert.deepEqual([set.size, set.comfortable], [12, 9]);
  assert.ok((set.most ?? Infinity) < 4000, String(set.most));
  // The deck's check says the text is longer than its slide holds at that size.
  assert.equal(validateDeck(twelve, deckOf([long]), { layout: 'adaptive' }).some((issue) => issue.slot === 'b1Text'), true);
  // What a free slide of one text holds is counted at it.
  assert.equal(freeTextRoom(twelve, {})?.sizes.densest, 12);
  assert.equal(freeTextRoom(demo, {})?.sizes.densest, 9);
  // A text the first room holds at four fifths of the template's size is no exception: under 16 pt it is set at 16, and cut.
  const sixteen = withSettings({ leastTextSize: 16 });
  const middling = slideOf(textOf(1500, 6), true);
  assert.ok((textOf1(demo, middling).size ?? 0) < 16);
  assert.equal(textOf1(sixteen, middling).size, 16);
  // A value over the template's own size of running text (18 pt here) is taken as that size.
  const most = withSettings({ leastTextSize: 24 });
  assert.equal(textOf1(most, long).size, 18);
  assert.equal(freeTextRoom(most, {})?.sizes.densest, 18);
  // A lower value lowers the size under which a slide is told its text is small.
  assert.equal(textOf1(withSettings({ leastTextSize: 7 }), long).comfortable, 7);
});

/** A document's text: a heading over each of its parts. */
const sectioned = (sections: number, chars: number): { paragraphs: string[]; headings: string[] } => {
  const NAMES = ['Riders', 'Costs', 'Stations', 'Staff', 'Timetable', 'Safety', 'Tickets', 'Outlook', 'Budget', 'Parking'];
  const parts = textOf(chars, sections);
  const headings = parts.map((_, at) => NAMES[at % NAMES.length] as string);
  return { paragraphs: parts.flatMap((part, at) => [headings[at] as string, part]), headings };
};

test('the headings of a document are counted as they are set: run in, a text with them is set larger than one with each on a line', () => {
  for (const [sections, chars, plain, headed] of [
    [6, 2400, 12, 12.5],
    [12, 2400, 10.5, 11.5],
  ] as const) {
    const { paragraphs, headings } = sectioned(sections, chars);
    assert.equal(textOf1(demo, slideOf(paragraphs, true)).size, plain, `${sections} as plain paragraphs`);
    assert.equal(textOf1(demo, slideOf(paragraphs, true, { headings })).size, headed, `${sections} set as headings`);
    // At that size every heading is set run in, and the deck's check finds it fits.
    const slide = slideOf(paragraphs, true, { headings });
    assert.deepEqual(validateDeck(demo, deckOf([slide]), { layout: 'adaptive' }), []);
    const xml = strFromU8(unzipSync(renderDeck(demo, deckOf([slide]), { layout: 'adaptive' }).pptx)['ppt/slides/slide1.xml'] as Uint8Array);
    const text = /<p:sp>(?:(?!<\/p:sp>)[\s\S])*?Riders(?:(?!<\/p:sp>)[\s\S])*<\/p:sp>/.exec(xml)?.[0] ?? '';
    assert.equal((text.match(/<a:p>/g) ?? []).length, sections, `${sections}: one paragraph a part`);
  }
});

test('the plan asks the renderer’s own question of each heading, in a line a character shorter', () => {
  const { paragraphs, headings } = sectioned(4, 800);
  const heads = headingsOver(paragraphs, headingPlaces(paragraphs, headings));
  assert.deepEqual(heads.map((head) => [head.at, head.heading]), [[0, 'Riders'], [2, 'Costs'], [4, 'Stations'], [6, 'Staff']]);
  // A heading over another heading, or over the last paragraph's absence, is none of these.
  assert.deepEqual(headingsOver(['Part one', 'Riders', 'Text.'], [0, 1]).map((head) => head.at), [1]);
  // The room of a line: counted 0.55 of the size a character, or the template's own where it is wider, 85% full.
  assert.equal(headingRoom(200, 10), Math.floor(200 / 5.5) * 0.85);
  assert.equal(headingRoom(200, 10, 0.43), headingRoom(200, 10));
  assert.equal(headingRoom(200, 10, 0.6), Math.floor(200 / 6) * 0.85);
  assert.equal(runsIn('Riders', 'Students and commuters.', 'ltr', headingRoom(200, 10)), true);
  assert.equal(runsIn('Riders', '- Students', 'ltr', headingRoom(200, 10)), false);
  assert.equal(runsIn('A heading far too long for a line of this column at all', 'Text.', 'ltr', headingRoom(200, 10)), false);
});

test('a document planned to be set as it is carries its headings, as its slide settles them', () => {
  const doc = readDocument(['Background', 'Opening text of the report, a sentence long.', '', 'Riders', 'Students and commuters make up most members.', '', 'Tasks', '- Repair the docks', '- Hire two mechanics', '', 'Costs', 'A yearly pass costs sixty euros.'].join('\n'), 'en');
  const numbers = doc.passages.map((_, at) => at + 1);
  const paragraphs = fillParagraphs(doc.passages, numbers, { bullets: true });
  const shape = shapeOf(doc, numbers);
  const places = documentHeadingPlaces(paragraphs, doc.passages, numbers);
  assert.deepEqual(shape.asIs?.heads?.map((head) => head.heading), places.map((at) => paragraphs[at]));
  // The slide, once settled, marks the same headings.
  const settled = settleDocumentSlide({ pattern: 'free', blocks: ['text'], slots: { title: 'Background', b1Text: paragraphs } }, doc.passages, numbers);
  assert.deepEqual(settled.headings, places.map((at) => (paragraphs[at] as string).trim()));
  assert.deepEqual(settled.headings, ['Riders', 'Tasks', 'Costs']);
});

test('a long text set dense that two columns do not hold at a size takes three, where a line of three holds 45 characters', () => {
  const slide = slideOf(textOf(1600, 10), true);
  const set = textOf1(demo, slide);
  assert.deepEqual([set.size, set.columns], [14.5, 3]);
  // At that size its lines in two columns would be no longer than 95 characters as PowerPoint sets them: two is what the length of its lines asks.
  const kit = freeKit(demo);
  assert.ok(kit);
  const { plan } = resolved(freeFor(demo, {}, slide));
  const box = plan.shapes.find((shape) => shape.name.startsWith('Block 1'));
  const width = (box?.box.w ?? 0) * (kit.slideWidth / 12700);
  assert.ok(width / (95 * 0.42 * 14.5) <= 2, String(width));
  assert.ok((width - 2 * kit.gap * (kit.slideWidth / 12700)) / 3 / (0.42 * 14.5) >= 45, String(width));
  assert.deepEqual(validateDeck(demo, deckOf([slide]), { layout: 'adaptive' }), []);
});

test('an instruction to a studio can set and clear how wide the template sets its text and the least size of a long text set dense', async () => {
  const answers = [
    { note: '', changes: [{ op: 'setManifest', pattern: '', shape: '', slot: '', values: [{ property: 'characterWidth', value: '0.43' }, { property: 'leastTextSize', value: '10' }] }] },
    { note: '', changes: [{ op: 'setManifest', pattern: '', shape: '', slot: '', values: [{ property: 'characterWidth', value: '' }, { property: 'leastTextSize', value: '' }] }] },
  ];
  const prompts: string[] = [];
  const llm: LlmProvider = {
    generateJson: async <T>(request: JsonRequest) => {
      prompts.push(request.prompt);
      return answers[prompts.length - 1] as T;
    },
  };
  const set = await editManifest(demo, 'We measured our text at 0.43; never set running text under 10 pt.', { llm });
  assert.deepEqual(set.rejected, []);
  assert.deepEqual([set.manifest.characterWidth, set.manifest.leastTextSize], [0.43, 10]);
  assert.match(prompts[0] ?? '', /characterWidth \(how wide PowerPoint sets a character of the template's text/);
  const cleared = await editManifest(loadTemplate(demoPptx, set.manifest), 'Go back to the estimate and the usual least size.', { llm });
  assert.deepEqual(cleared.rejected, []);
  assert.deepEqual(['characterWidth' in cleared.manifest, 'leastTextSize' in cleared.manifest], [false, false]);
  // The model is told the settings the template has.
  assert.match(prompts[1] ?? '', /a character of its text is counted 0\.43 of its size wide; a long text a person asked to have whole, set dense on a slide the software lays out, is set no smaller than 10 pt/);
});

test('headings set run in, or a narrower width, never keep a long text out of the larger room where it is set larger', () => {
  // Nine short parts under headings: counted run in, they fit the first room at four fifths of the
  // template's size, and kept there they came out at 14.5 pt in one column; in the larger room,
  // as without headings, they are set at 22 in two.
  const { paragraphs, headings } = sectioned(9, 300);
  const plain = textOf1(demo, slideOf(paragraphs, true));
  const headed = textOf1(demo, slideOf(paragraphs, true, { headings }));
  assert.deepEqual([plain.size, plain.columns], [22, 2]);
  assert.ok((headed.size ?? 0) >= (plain.size ?? 0), JSON.stringify(headed));
  for (const [sections, chars] of [[8, 600], [10, 300], [11, 300], [4, 1200], [8, 800]] as const) {
    const text = sectioned(sections, chars);
    assert.ok((textOf1(demo, slideOf(text.paragraphs, true, { headings: text.headings })).size ?? 0) >= (textOf1(demo, slideOf(text.paragraphs, true)).size ?? 0), `${sections} x ${chars}`);
  }
  // A template that states a narrower width: a text is never set smaller than without it.
  const narrow = withSettings({ characterWidth: 0.44 });
  for (const [chars, count] of [[1050, 10], [1350, 8], [1000, 10], [2000, 8], [3000, 12]] as const) {
    const slide = slideOf(textOf(chars, count), true);
    assert.ok((textOf1(narrow, slide).size ?? 0) >= (textOf1(demo, slide).size ?? 0), `${chars} in ${count}`);
  }
});

test('a document’s text set as it is is planned with its headings as the slide sets it, also in the larger room', () => {
  for (const sections of [6, 12]) {
    const { paragraphs, headings } = sectioned(sections, 4000);
    const each = paragraphs.map((paragraph) => paragraph.length);
    const chars = each.reduce((sum, length) => sum + length, 0);
    const heads = headingsOver(paragraphs, headingPlaces(paragraphs, headings));
    const planning = { icons: false, lengths: { b1Text: chars }, paragraphs: { b1Text: each.length }, asIs: { b1Text: each }, heads: { b1Text: heads } };
    const planned = resolved(freeFor(demo, {}, { blocks: ['text'], arrange: { dense: true }, slots: {} }, planning)).plan;
    const set = resolved(freeFor(demo, {}, slideOf(paragraphs, true, { headings }))).plan;
    const of = (plan: typeof planned) => {
      const slot = plan.pattern.slots.find((entry) => entry.name === 'b1Text');
      const shape = plan.shapes.find((entry) => entry.id === slot?.shapeId);
      return { size: shape?.size, columns: shape?.columns ?? 1, paragraphs: slot?.kind === 'text' ? slot.maxParagraphs : undefined };
    };
    assert.deepEqual(of(planned), of(set), `${sections} sections`);
    // And the slide holds it: code sets it as it is.
    assert.deepEqual(validateDeck(demo, deckOf([slideOf(paragraphs, true, { headings })]), { layout: 'adaptive' }), []);
  }
});

test('a deck that asks for larger text gets what 0.6.6 gave it, where a text is now counted larger without the wish', () => {
  const part = (length: number) => SENTENCE.repeat(Math.ceil(length / SENTENCE.length) + 1).slice(0, length - 1).trim() + '.';
  const sizes = (template: typeof demo, slide: SlideSpec) => {
    const free = resolved(freeFor(template, { textScale: 1.15 }, slide));
    const of = (name: string) => {
      const slot = free.plan.pattern.slots.find((entry) => entry.name === name);
      return slot?.kind === 'text' ? slot.fontSize : undefined;
    };
    return { text: of('b1Text'), note: of('b2Text'), enlarged: free.enlarged ?? 'all' };
  };
  const deck = (template: typeof demo, slide: SlideSpec) => validateDeck(template, { ...deckOf([slide]), textScale: 1.15 }, { layout: 'adaptive' });
  // Headings run in: the text is 18.5 pt without the wish, and 0.6.6 gave the slide all of the wish, 18 and 19.5.
  const headed: SlideSpec = { pattern: 'free', blocks: ['text', 'note'], slots: { title: 'A title', b1Text: ['Riders', part(133), 'Costs and budget', part(133), 'Stations', part(133)], b2Text: part(60) }, headings: ['Riders', 'Costs and budget', 'Stations'] };
  assert.deepEqual(sizes(demo, headed), { text: 18, note: 19.5, enlarged: 'all' });
  assert.deepEqual(deck(demo, headed), []);
  // A template's width: the same.
  const plain: SlideSpec = { pattern: 'free', blocks: ['text', 'note'], slots: { title: 'A title', b1Text: ['Riders', part(275), 'Costs and budget', part(275)], b2Text: part(60) } };
  const narrow = withSettings({ characterWidth: 0.45 });
  assert.deepEqual(sizes(narrow, plain), { text: 18, note: 19.5, enlarged: 'all' });
  assert.deepEqual(deck(narrow, plain), []);
});

test('a writer inside what it is told is not sent back where the template states its width', async () => {
  const { buildFixturePptx, fixtureManifest } = await import('./fixture.ts');
  const SENT = 'הצוות עובד יחד כל שבוע כדי לשפר את השירות לכל מי שגר בעיר ובסביבתה. ';
  const part = (length: number) => SENT.repeat(Math.ceil(length / SENT.length) + 1).slice(0, length - 1).trim() + '.';
  // A slide 10 inches wide at four fifths of its size: the least size that reads is over four fifths of the template's.
  for (const characterWidth of [0.45, 0.47, 0.48, 0.5]) {
    const template = loadTemplate(buildFixturePptx(), { ...fixtureManifest, characterWidth });
    const deck: DeckSpec = { version: 1, language: 'he', direction: 'rtl', textScale: 0.8, fields: { closing: 'תודה' }, slides: [] };
    const planned = resolved(freeFor(template, deck, { blocks: ['text'], arrange: { dense: true }, slots: {} }, { icons: false, lengths: { b1Text: 1800 }, paragraphs: { b1Text: 6 } })).plan;
    const told = planned.pattern.slots.find((slot) => slot.name === 'b1Text');
    assert.ok(told?.kind === 'text' && told.maxChars !== undefined);
    const written = [0.28, 0.16, 0.32, 0.215, 0.025].map((share) => part(Math.floor(share * ((told.maxChars ?? 0) - 10))));
    const slide: SlideSpec = { pattern: 'free', blocks: ['text'], arrange: { dense: true }, slots: { title: 'כותרת', b1Text: written } };
    assert.deepEqual(validateDeck(template, { ...deck, slides: [slide] }, { layout: 'adaptive' }).filter((issue) => issue.slot === 'b1Text'), [], String(characterWidth));
  }
});

test('a template’s width does not set a short text of many paragraphs further down than its box holds', () => {
  const slide = slideOf(Array.from({ length: 10 }, () => SENTENCE.slice(0, 29).trim() + '.'));
  const box = (template: typeof demo) => {
    const slot = resolved(freeFor(template, {}, slide)).plan.pattern.slots.find((entry) => entry.name === 'b1Text');
    return slot?.kind === 'text' ? slot.box : undefined;
  };
  for (const characterWidth of [0.43, 0.44, 0.45]) {
    const narrow = withSettings({ characterWidth });
    assert.deepEqual(box(narrow), box(demo), String(characterWidth));
    assert.deepEqual(validateDeck(narrow, deckOf([slide]), { layout: 'adaptive' }), [], String(characterWidth));
  }
});

test('on a template wider than the estimate, a text with headings is set no smaller than without them', () => {
  const wide = withSettings({ characterWidth: 0.6 });
  const text = ['Riders', ...textOf(1300, 1)];
  assert.ok((textOf1(wide, slideOf(text, true, { headings: ['Riders'] })).size ?? 0) >= (textOf1(wide, slideOf(text, true)).size ?? 0));
});

test('leastTextSize: what a slide set dense is said to hold, it holds; a size between half points is taken up; where nothing fits, the columns are as before', () => {
  for (const least of [12, 15, 16, 24]) {
    const template = withSettings({ leastTextSize: least });
    const room = freeTextRoom(template, {});
    assert.ok(room);
    const planned = resolved(freeFor(template, {}, { blocks: ['text'], arrange: { dense: true }, slots: {} }, { icons: false, lengths: { b1Text: room.densest } })).plan;
    const slot = planned.pattern.slots.find((entry) => entry.name === 'b1Text');
    assert.ok(slot?.kind === 'text' && (slot.maxChars ?? 0) >= room.densest, `${least}: ${room.densest} said, ${slot?.kind === 'text' ? slot.maxChars : ''} held`);
  }
  // 12.2 is 12.5: text is set in half points, and 12 would be under it.
  const between = withSettings({ leastTextSize: 12.2 });
  assert.equal(freeKit(between)?.leastTextSize, 12.5);
  assert.equal(textOf1(between, slideOf(textOf(11000, 12), true)).size, 12.5);
  // A text that fits at no size takes the columns its lines take at the least size, as in 0.6.6: three at 9 pt on the demo, two at 12.
  assert.deepEqual([textOf1(demo, slideOf(textOf(11000, 12), true)).size, textOf1(demo, slideOf(textOf(11000, 12), true)).columns], [9, 3]);
  const twelve = withSettings({ leastTextSize: 12 });
  for (const [chars, count] of [[3000, 40], [5000, 6]] as const) assert.deepEqual([textOf1(twelve, slideOf(textOf(chars, count), true)).size, textOf1(twelve, slideOf(textOf(chars, count), true)).columns], [12, 2], `${chars} in ${count}`);
});

test('with the template’s fonts, a text with headings set run in is measured as it is set: the check and the renderer keep the size planned', () => {
  const fonts = fontkitMetrics(fontkit, openFonts());
  const template = loadTemplate(demoPptx, { ...demoManifest, shrinkText: { to: 0.6 } }, { fonts });
  const measure = textMeasurer(template, fonts);
  for (const [sections, chars] of [[12, 300], [12, 900]] as const) {
    const { paragraphs, headings } = sectioned(sections, chars);
    const slide = slideOf(paragraphs, true, { headings });
    const planned = textOf1(template, slide).size;
    assert.deepEqual(validateDeck(template, deckOf([slide]), { layout: 'adaptive', measure }), [], `${sections} x ${chars}`);
    const xml = strFromU8(unzipSync(renderDeck(template, deckOf([slide]), { layout: 'adaptive' }).pptx)['ppt/slides/slide1.xml'] as Uint8Array);
    const shape = xml.split('<p:sp>').find((part) => part.includes((paragraphs[1] ?? '').slice(0, 20))) ?? '';
    assert.deepEqual([...new Set([...shape.matchAll(/sz="(\d+)"/g)].map((match) => Number(match[1]) / 100))], [planned], `${sections} x ${chars}`);
  }
});

test('with a least size stated, a writer inside what a slide set dense tells it is not sent back, in any paragraphs', () => {
  for (const [settings, textScale, asked] of [
    [{ leastTextSize: 16 }, 1, 1200],
    [{ characterWidth: 0.43, leastTextSize: 16 }, 1, 1100],
    [{ characterWidth: 0.45, leastTextSize: 12.2 }, 0.8, 2000],
    [{ characterWidth: 0.43, leastTextSize: 16 }, 0.8, 1000],
    [{ characterWidth: 0.45, leastTextSize: 12 }, 0.8, 1900],
  ] as const) {
    const template = withSettings(settings);
    const planned = resolved(freeFor(template, { textScale }, { blocks: ['text'], arrange: { dense: true }, slots: {} }, { icons: false, lengths: { b1Text: asked } })).plan;
    const told = planned.pattern.slots.find((slot) => slot.name === 'b1Text');
    assert.ok(told?.kind === 'text' && told.maxChars !== undefined && told.maxParagraphs !== undefined);
    for (let count = 1; count <= (told.maxParagraphs ?? 1); count++) {
      const slide = slideOf(textOf((told.maxChars ?? 0) - 5, count), true);
      const issues = validateDeck(template, { ...deckOf([slide]), textScale }, { layout: 'adaptive' }).filter((issue) => issue.slot === 'b1Text');
      assert.deepEqual(issues, [], `${JSON.stringify(settings)} at ${textScale}: told ${told.maxChars} in ${told.maxParagraphs}, wrote ${count}`);
    }
  }
});

test('three texts side by side in a deck that asks for larger text keep what 0.6.6 held, on a template that states its width', () => {
  const narrow = withSettings({ characterWidth: 0.43 });
  const slide: SlideSpec = { pattern: 'free', blocks: ['text', 'text', 'text'], slots: { title: 'A title', b1Text: textOf(400, 4), b2Text: textOf(400, 4), b3Text: textOf(400, 4) } };
  for (const textScale of [1.15, 1.25, 1.5]) assert.deepEqual(validateDeck(narrow, { ...deckOf([slide]), textScale }, { layout: 'adaptive' }), [], String(textScale));
});

test('with a least size stated, a text with headings is set no smaller than without them, and what a slide set dense is said to hold it holds', () => {
  for (const [leastTextSize, textScale, sections, chars] of [[12, 0.8, 14, 900], [14, 0.8, 6, 450], [14, 0.8, 8, 450], [16, 1, 4, 600], [18, 1.15, 6, 600]] as const) {
    const least = withSettings({ leastTextSize });
    const text = sectioned(sections, chars);
    const size = (slide: SlideSpec) => {
      const { plan } = resolved(freeFor(least, { textScale }, slide));
      const slot = plan.pattern.slots.find((entry) => entry.name === 'b1Text');
      return slot?.kind === 'text' ? (slot.fontSize ?? 0) : 0;
    };
    assert.ok(size(slideOf(text.paragraphs, true, { headings: text.headings })) >= size(slideOf(text.paragraphs, true)), `${leastTextSize} pt at ${textScale}: ${sections} x ${chars}`);
  }
  for (const settings of [{ characterWidth: 0.45, leastTextSize: 12 }, { leastTextSize: 14 }]) {
    for (const textScale of [0.8, 1]) {
      const template = withSettings(settings);
      const room = freeTextRoom(template, { textScale });
      assert.ok(room);
      const planned = resolved(freeFor(template, { textScale }, { blocks: ['text'], arrange: { dense: true }, slots: {} }, { icons: false, lengths: { b1Text: room.densest } })).plan;
      const slot = planned.pattern.slots.find((entry) => entry.name === 'b1Text');
      assert.ok(slot?.kind === 'text' && (slot.maxChars ?? 0) >= room.densest, `${JSON.stringify(settings)} at ${textScale}: ${room.densest} said, ${slot?.kind === 'text' ? slot.maxChars : ''} held`);
    }
  }
});

test('a least size at the template’s own size on a small template still offers a slide set dense what it holds', async () => {
  const { buildFixturePptx, fixtureManifest } = await import('./fixture.ts');
  for (const leastTextSize of [12, 16]) {
    const template = loadTemplate(buildFixturePptx(), { ...fixtureManifest, leastTextSize });
    for (const textScale of [1.3, 1.5]) {
      const deck = { direction: 'rtl' as const, textScale };
      const room = freeTextRoom(template, deck);
      assert.ok(room && room.densest > 0, `${leastTextSize} at ${textScale}: ${JSON.stringify(room)}`);
      const planned = resolved(freeFor(template, deck, { blocks: ['text'], arrange: { dense: true }, slots: {} }, { icons: false, lengths: { b1Text: room.densest } })).plan;
      const slot = planned.pattern.slots.find((entry) => entry.name === 'b1Text');
      assert.ok(slot?.kind === 'text' && (slot.maxChars ?? 0) >= room.densest, `${leastTextSize} at ${textScale}: ${room.densest} said, ${slot?.kind === 'text' ? slot.maxChars : ''} held`);
    }
  }
});
