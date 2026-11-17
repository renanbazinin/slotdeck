// A long text a person asked to have whole (`FreeArrange.dense`): a free slide of one text that its
// slide does not hold at four fifths of the template's size is set smaller, down to the least size
// that reads, in two or three columns, and is planned for the most its writer may write. Without
// the flag every free slide is laid out as in 0.6.1. No test calls a model.
import assert from 'node:assert/strict';
import { readFileSync } from 'node:fs';
import { dirname, join } from 'node:path';
import { test } from 'node:test';
import { fileURLToPath } from 'node:url';
import { loadTemplate, validateDeck } from '../src/index.ts';
import type { DeckSpec, FreeArrange, SlideSpec, TemplateManifest } from '../src/index.ts';
import { freeTextRoom } from '../src/free/offer.ts';
import { readableSize } from '../src/free/plan.ts';
import type { FreeTextSlot } from '../src/free/plan.ts';
import { freeFor } from '../src/free/resolve.ts';
import type { FreeResolved } from '../src/free/resolve.ts';
import { buildFixturePptx, fixtureManifest } from './fixture.ts';

const root = join(dirname(fileURLToPath(import.meta.url)), '..');
const demoManifest = JSON.parse(readFileSync(join(root, 'templates', 'demo.manifest.json'), 'utf8')) as TemplateManifest;
const demo = loadTemplate(new Uint8Array(readFileSync(join(root, 'templates', 'demo.pptx'))), demoManifest);
// A right-to-left template on a slide 10 inches wide, whose frame leaves a text a narrow room.
const rtl = loadTemplate(buildFixturePptx(), fixtureManifest);
const decks: Record<'demo' | 'rtl', Omit<DeckSpec, 'slides'>> = {
  demo: { version: 1, language: 'en', fields: { organizationDepartment: 'Retail' } },
  rtl: { version: 1, language: 'he', direction: 'rtl', fields: { closing: 'Thank you' } },
};
const SENTENCES = { en: 'Riders rely on the service every day, and the city gains from it in ways that are easy to see. ', he: 'הצוות עובד יחד כל שבוע כדי לשפר את השירות לכל מי שגר בעיר ובסביבתה. ' };

/** A text of so many characters in so many paragraphs, as alike as they can be: paragraphs of about 300 when no number is given. */
const textOf = (chars: number, sentence: string, count = Math.max(1, Math.round(chars / 300))): string[] =>
  Array.from({ length: count }, (_, at) => {
    const length = Math.floor(chars / count) + (at < chars % count ? 1 : 0);
    return sentence.repeat(Math.ceil(length / sentence.length) + 1).slice(0, length - 1) + '.';
  });
const resolved = (value: FreeResolved | string): FreeResolved => {
  assert.notEqual(typeof value, 'string', String(value));
  return value as FreeResolved;
};
/** The plan of a free slide of one text before there is content, set out for so many characters. */
const planned = (template = demo, chars: number, arrange?: FreeArrange) => resolved(freeFor(template, {}, { blocks: ['text'], slots: {}, ...(arrange ? { arrange } : {}) }, { icons: false, lengths: { b1Text: chars } })).plan;
/** A free slide of one text with its content. */
const slideOf = (text: string[], arrange?: FreeArrange): SlideSpec => ({ pattern: 'free', blocks: ['text'], ...(arrange ? { arrange } : {}), slots: { title: 'A title', b1Text: text } });
const textSlot = (plan: FreeResolved['plan']): FreeTextSlot => plan.pattern.slots.find((slot) => slot.name === 'b1Text') as FreeTextSlot;
const issuesOf = (template: typeof demo, name: 'demo' | 'rtl', slide: SlideSpec) => validateDeck(template, { ...decks[name], slides: [slide] }, { layout: 'adaptive' });
const tens = (chars: number) => Math.floor(chars / 10) * 10;

test('without dense every free slide is laid out as before', () => {
  const before = JSON.parse(readFileSync(join(root, 'test', 'fixtures', 'free-0.6.1.json'), 'utf8')) as Record<string, unknown>;
  let compared = 0;
  for (const [name, template, sentence] of [['demo', demo, SENTENCES.en], ['rtl', rtl, SENTENCES.he]] as const) {
    for (let chars = 300; chars <= 3000; chars += 300) {
      for (const arrange of [undefined, {}]) {
        const plan = (slide: Pick<SlideSpec, 'blocks' | 'arrange' | 'slots'>, planning?: Parameters<typeof freeFor>[3]) => JSON.parse(JSON.stringify(resolved(freeFor(template, {}, slide, planning)).plan)) as unknown;
        const extra = arrange ? { arrange } : {};
        assert.deepEqual(plan({ blocks: ['text'], slots: {}, ...extra }, { icons: false, lengths: { b1Text: chars } }), before[`${name}/planning/${chars}`], `${name} planning ${chars}`);
        assert.deepEqual(plan({ blocks: ['text'], slots: { title: 'A title', b1Text: textOf(chars, sentence) }, ...extra }), before[`${name}/content/${chars}`], `${name} content ${chars}`);
        compared += 2;
      }
    }
  }
  assert.equal(compared, 80);
});

test('a dense text that fits at four fifths is set as without the flag', () => {
  assert.deepEqual(planned(demo, 1000, { dense: true }), planned(demo, 1000));
  const text = textOf(1000, SENTENCES.en);
  assert.deepEqual(resolved(freeFor(demo, {}, slideOf(text, { dense: true }))).plan, resolved(freeFor(demo, {}, slideOf(text))).plan);
});

test('a long text set dense is set smaller in the larger room, in two or three columns', () => {
  const content = (template: typeof demo, chars: number, count: number, sentence: string) => resolved(freeFor(template, {}, slideOf(textOf(chars, sentence, count), { dense: true })));
  const three = content(demo, 3000, 7, SENTENCES.en);
  const slot = textSlot(three.plan);
  assert.ok((slot.fontSize ?? 0) >= 11 && (slot.fontSize ?? 0) <= 12.5, String(slot.fontSize));
  assert.ok([2, 3].includes(slot.drawn.columns ?? 1), String(slot.drawn.columns));
  assert.equal(three.plan.layout.dense, true);
  // Without the flag the same text is set at four fifths of the template's size, and runs over.
  const plain = resolved(freeFor(demo, {}, slideOf(textOf(3000, SENTENCES.en, 7))));
  assert.ok((textSlot(plain.plan).fontSize ?? 0) > (slot.fontSize ?? 0));
  assert.ok((textSlot(plain.plan).maxChars ?? 0) < 3000);

  const five = textSlot(content(demo, 4900, 9, SENTENCES.en).plan);
  assert.equal(five.fontSize, 9);
  assert.equal(five.drawn.columns, 3);
  assert.deepEqual(issuesOf(demo, 'demo', slideOf(textOf(4900, SENTENCES.en, 9), { dense: true })), []);

  // Under the least size that reads it is not set: the text is over its budget.
  const seven = textSlot(content(demo, 7000, 23, SENTENCES.en).plan);
  assert.equal(seven.fontSize, 9);
  assert.ok((seven.maxChars ?? 0) < 7000);

  const hebrew = textSlot(content(rtl, 3250, 7, SENTENCES.he).plan);
  assert.ok((hebrew.fontSize ?? 0) >= readableSize(9144000) && (hebrew.fontSize ?? 0) <= 8, String(hebrew.fontSize));
});

test('readableSize', () => {
  assert.equal(readableSize(12192000), 9);
  assert.equal(readableSize(9144000), 7);
  assert.equal(readableSize(3000000), 5);
});

test('readableSize and freeTextRoom stay inside the package', async () => {
  const api: Record<string, unknown> = await import('../src/index.ts');
  for (const name of ['readableSize', 'freeTextRoom']) assert.equal(name in api, false, name);
});

test('a dense slide validates and stays in its room', () => {
  const slide = slideOf(textOf(3000, SENTENCES.en, 7), { dense: true });
  assert.deepEqual(issuesOf(demo, 'demo', slide), []);
  const { plan } = resolved(freeFor(demo, {}, slide));
  const room = plan.layout.blocks[0]?.box;
  assert.ok(room);
  const texts = plan.shapes.filter((shape) => shape.kind === 'text');
  assert.ok(texts.length > 0);
  for (const shape of texts) assert.ok(shape.box.y + shape.box.h <= room.y + room.h + 1e-4, shape.name);
});

test('a dense text is planned for the most the writer is told', () => {
  const counts = Array.from({ length: 12 }, (_, at) => at + 1);
  /** Whether a text of so many characters, in so many equal paragraphs, is flagged or set smaller than `size`. */
  const over = (template: typeof demo, name: 'demo' | 'rtl', total: number, sentence: string, count: number, size: number) => {
    const slide = slideOf(textOf(total, sentence, count), { dense: true });
    return issuesOf(template, name, slide).length > 0 || (textSlot(resolved(freeFor(template, {}, slide)).plan).fontSize ?? 0) < size;
  };
  // 800 English words are 4,800 characters.
  for (const [name, template, chars, sentence] of [['rtl', rtl, 1500, SENTENCES.he], ['rtl', rtl, 3250, SENTENCES.he], ['demo', demo, 3000, SENTENCES.en], ['demo', demo, 4800, SENTENCES.en], ['demo', demo, 4900, SENTENCES.en]] as const) {
    const plan = planned(template, chars, { dense: true });
    const slot = textSlot(plan);
    const budget = slot.maxChars ?? 0;
    // It says at most 12 paragraphs: round 2 said 16 for 4,900 characters, one for each 300.
    assert.equal(slot.maxParagraphs, 12, `${name} ${chars}`);
    // Written at the top of the range its writer is told, or at its budget, in any number of equal
    // paragraphs up to 12, the text is set as large as it fits and validates. Each paragraph ends
    // on a line of its own, so 7, 10 or 11 can take more lines than 12: a budget of what 12 hold
    // sent a writer inside it back to shorten its text.
    for (const total of [Math.min(tens(1.15 * chars), budget), budget]) {
      for (const count of counts) {
        const slide = slideOf(textOf(total, sentence, count), { dense: true });
        assert.deepEqual(issuesOf(template, name, slide), [], `${name} ${chars}: ${total} in ${count} paragraphs`);
        // Its budget is at most what they hold at the size it was planned at.
        const written = textSlot(resolved(freeFor(template, {}, slide)).plan);
        assert.ok((written.fontSize ?? 0) >= (slot.fontSize ?? 0), `${name} ${chars}: ${total} in ${count} written at ${written.fontSize}, planned at ${slot.fontSize}`);
      }
    }
    // Ten more characters do not fit at the size it was planned at, in some number of paragraphs.
    assert.ok(counts.some((count) => over(template, name, budget + 10, sentence, count, slot.fontSize ?? 0)), `${name} ${chars}: ${budget + 10}`);
  }
  // 500 English words (3,000 characters): the slot holds at least 3,450, more than the 3,440 its writer may write.
  assert.ok((textSlot(planned(demo, 3000, { dense: true })).maxChars ?? 0) >= 3450);
  // 800 English words are held.
  assert.ok((textSlot(planned(demo, 4800, { dense: true })).maxChars ?? 0) >= 4800);
  // A text over what the slide holds at the least size that reads is planned at that size, and
  // its budget is what the slide holds there in any number of paragraphs up to 12.
  const floor = readableSize(9144000);
  const most = textSlot(planned(rtl, 3600, { dense: true }));
  assert.equal(most.fontSize, floor);
  assert.ok((most.maxChars ?? 0) < 3600);
  for (const count of counts) {
    const slide = slideOf(textOf(most.maxChars ?? 0, SENTENCES.he, count), { dense: true });
    assert.deepEqual(issuesOf(rtl, 'rtl', slide), [], `${most.maxChars} in ${count} paragraphs`);
    assert.equal(textSlot(resolved(freeFor(rtl, {}, slide)).plan).fontSize, floor);
  }
});

test('a planner that gives the number of paragraphs gets a budget for that many', () => {
  const twelve = textSlot(resolved(freeFor(demo, {}, { blocks: ['text'], slots: {}, arrange: { dense: true } }, { icons: false, lengths: { b1Text: 3000 }, paragraphs: { b1Text: 12 } })).plan);
  const three = textSlot(resolved(freeFor(demo, {}, { blocks: ['text'], slots: {}, arrange: { dense: true } }, { icons: false, lengths: { b1Text: 3000 }, paragraphs: { b1Text: 3 } })).plan);
  // Set out for the text as it will be (3,000 characters), not for the top of a range.
  assert.ok((three.fontSize ?? 0) >= (twelve.fontSize ?? 0));
  assert.ok((three.maxChars ?? 0) >= 3000 && (twelve.maxChars ?? 0) >= 3000);
  assert.deepEqual(issuesOf(demo, 'demo', slideOf(textOf(twelve.maxChars ?? 0, SENTENCES.en, 12), { dense: true })), []);
  // Without the flag too, the paragraphs are counted one by one: more of them hold less.
  const plain = (count: number) => textSlot(resolved(freeFor(demo, {}, { blocks: ['text'], slots: {} }, { icons: false, lengths: { b1Text: 3000 }, paragraphs: { b1Text: count } })).plan);
  assert.ok((plain(12).maxChars ?? 0) < (plain(3).maxChars ?? 0));
  assert.equal(plain(12).maxParagraphs, 12);
});

test('freeTextRoom finds what a free slide of one text holds, set as usual and set dense', () => {
  const room = freeTextRoom(demo, {});
  assert.ok(room);
  assert.ok(room.comfortable >= 900 && room.comfortable <= 1000, String(room.comfortable));
  assert.equal(room.sizes.comfortable, 18);
  assert.ok(room.dense >= 1300 && room.dense <= 1400, String(room.dense));
  assert.equal(room.sizes.dense, 14.5);
  const twelve = freeTextRoom(demo, {}, 12);
  assert.ok(twelve && twelve.densest >= 4800 && twelve.densest <= 5400, String(twelve?.densest));
  assert.equal(twelve?.sizes.densest, 9);
  // Without a number of paragraphs, a text set dense is written as a writer told a length writes
  // it, in any number up to 12, and holds less than in 12: `densest` is the most a slide can be
  // asked for and hold, in steps of 50 characters.
  assert.ok(twelve && room.densest < twelve.densest, `${room.densest} of ${twelve?.densest}`);
  assert.ok((textSlot(planned(demo, room.densest, { dense: true })).maxChars ?? 0) >= room.densest);
  assert.ok((textSlot(planned(demo, room.densest + 50, { dense: true })).maxChars ?? 0) < room.densest + 50);
  assert.equal(room.sizes.densest, 9);
  // Remembered: a second call gives the same object.
  assert.equal(freeTextRoom(demo, {}), room);
  // A frame that leaves a text a narrow room holds less, and is set dense down to the least size that reads.
  const narrow = freeTextRoom(rtl, {}, 12);
  assert.equal(narrow?.sizes.densest, readableSize(9144000));
  assert.ok(narrow && narrow.densest > narrow.dense && narrow.densest < 3000, String(narrow?.densest));
  // A template whose manifest turns free slides off has none.
  assert.equal(freeTextRoom(loadTemplate(buildFixturePptx(), { ...fixtureManifest, free: false }), {}), undefined);
});
