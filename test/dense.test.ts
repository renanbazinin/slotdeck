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
import { freeKit } from '../src/free/kit.ts';
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

test('a dense text that fits at four fifths keeps its two columns, fitted column by column in a box as high as they need', () => {
  // Its columns hold its lines as PowerPoint fills them, each paragraph ending on a line of its own
  // with the space after it, and each column keeps a line free. It stays in its room, at the size
  // of the text without the flag or up to a point under it, and its slot holds the text.
  const compare = (dense: FreeResolved['plan'], plain: FreeResolved['plan'], chars: number) => {
    assert.equal(textSlot(dense).drawn.columns, 2);
    assert.equal(textSlot(plain).drawn.columns, 2);
    const room = dense.layout.blocks[0]?.box as { y: number; h: number };
    const box = textSlot(dense).box as { y: number; h: number };
    assert.ok(box.y + box.h <= room.y + room.h + 1e-4, JSON.stringify([box, room]));
    assert.ok((textSlot(plain).fontSize ?? 0) - (textSlot(dense).fontSize ?? 0) <= 1, JSON.stringify([textSlot(plain).fontSize, textSlot(dense).fontSize]));
    assert.ok((textSlot(dense).maxChars ?? 0) >= chars, String(textSlot(dense).maxChars));
  };
  for (const chars of [1000, 1300]) {
    compare(planned(demo, chars, { dense: true }), planned(demo, chars), chars);
    const text = textOf(chars, SENTENCES.en);
    compare(resolved(freeFor(demo, {}, slideOf(text, { dense: true }))).plan, resolved(freeFor(demo, {}, slideOf(text))).plan, chars);
  }
  // Before there is content its box counts its lines by length, paragraph by paragraph: higher
  // than the box of a text without the flag, which a first column ran under.
  const height = (plan: FreeResolved['plan']): number => textSlot(plan).box?.h ?? 0;
  assert.ok(height(planned(demo, 1000, { dense: true })) > height(planned(demo, 1000)));
  // Once written, its box is as high as its lines counted by length need, and a line more: the face
  // of a template is not known, and a box made for the words of the demo's narrow face was a line or
  // more too low in another.
  const demoKit = freeKit(demo);
  assert.ok(demoKit);
  const { slideWidth, slideHeight: high, gap } = demoKit;
  for (const chars of [1000, 1300, 3000]) {
    const written: string[] = textOf(chars, SENTENCES.en, Math.max(2, Math.round(chars / 250)));
    const drawn: FreeTextSlot = textSlot(resolved(freeFor(demo, {}, slideOf(written, { dense: true }))).plan);
    const pt: number = drawn.fontSize ?? 0;
    const count: number = drawn.drawn.columns ?? 1;
    const perColumn: number = ((drawn.box?.w ?? 0) * slideWidth) / 12700 / count - ((count - 1) * gap * slideWidth) / 12700 / count;
    const across: number = Math.floor(perColumn / (0.48 * pt)) * 0.85;
    const counted: number = written.reduce((sum, paragraph) => sum + Math.ceil(paragraph.length / across), 0);
    assert.ok(((drawn.box?.h ?? 0) * high) / 12700 >= Math.ceil(counted / count) * 1.2 * pt, `${chars}: ${JSON.stringify([drawn.box, pt, count, counted])}`);
  }
  // Twenty-four paragraphs of a line each: each column holds twelve lines and the space after
  // each of them, and the box has a line more.
  const lines = Array.from({ length: 24 }, (_, at) => `Point ${at + 1} of the plan, in a short line.`);
  const set = resolved(freeFor(demo, {}, slideOf(lines, { dense: true }))).plan;
  const slot = textSlot(set);
  const size = slot.fontSize ?? 0;
  assert.equal(slot.drawn.columns, 2);
  const slideHeight = freeKit(demo)?.slideHeight ?? 0;
  assert.ok(((slot.box?.h ?? 0) * slideHeight) / 12700 >= 13 * 1.2 * size + 12 * Math.round(0.5 * size) - 0.05, JSON.stringify([slot.box, size]));
});

test('a short dense text of many one-line paragraphs is not set down the slide, and is planned in the box it is drawn in', () => {
  // A lead line and eleven short items. Counted by its length such a text came to a few lines, and
  // was set so far down that its last items ran under its box: in PowerPoint, a list of 446
  // characters in twelve paragraphs ran two lines past its box. The deck's check said its slot
  // held fewer paragraphs.
  const lines = ['Notes', ...Array.from({ length: 11 }, (_, at) => `Point ${at + 1}.`)];
  const chars = lines.reduce((sum, line) => sum + line.length, 0);
  for (const [name, template] of [
    ['demo', demo],
    ['rtl', rtl],
  ] as const) {
    const set = resolved(freeFor(template, {}, slideOf(lines, { dense: true }))).plan;
    const slot = textSlot(set);
    const room = set.layout.blocks[0]?.box as { y: number; h: number };
    // It starts at the top of its room and holds all its paragraphs: the deck's check says nothing.
    assert.ok(Math.abs((slot.box?.y ?? 0) - room.y) < 0.02, `${name}: ${JSON.stringify([slot.box, room])}`);
    assert.deepEqual(issuesOf(template, name, slideOf(lines, { dense: true })), [], name);
    // Planned as code sets it (a document's text), it is planned in the same box, and holds the text.
    const asIs = { icons: false, lengths: { b1Text: chars }, paragraphs: { b1Text: lines.length }, asIs: { b1Text: lines.map((line) => line.length) } };
    const plan = resolved(freeFor(template, {}, { blocks: ['text'], slots: {}, arrange: { dense: true } }, asIs)).plan;
    assert.deepEqual(textSlot(plan).box, slot.box, name);
    assert.ok((textSlot(plan).maxChars ?? 0) >= chars, `${name}: ${textSlot(plan).maxChars}`);
    assert.equal(textSlot(plan).maxParagraphs, lines.length, name);
    // Planned for that many paragraphs of any length, as a writer may write them, it still holds the text.
    const told = resolved(freeFor(template, {}, { blocks: ['text'], slots: {}, arrange: { dense: true } }, { icons: false, lengths: { b1Text: chars }, paragraphs: { b1Text: lines.length } })).plan;
    assert.ok((textSlot(told).maxChars ?? 0) >= chars, `${name}: ${textSlot(told).maxChars}`);
    assert.equal(textSlot(told).maxParagraphs, lines.length, name);
  }
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
    // Twice the budget does not fit, in any number of paragraphs.
    assert.ok(counts.some((count) => over(template, name, 2 * budget, sentence, count, slot.fontSize ?? 0)), `${name} ${chars}: ${2 * budget}`);
  }
  // 500 English words (3,000 characters): the slot holds at least 3,450, more than the 3,440 its writer may write.
  assert.ok((textSlot(planned(demo, 3000, { dense: true })).maxChars ?? 0) >= 3450);
  // 725 English words (4,350 characters) are held, at 9 pt; 790 are not. Each column keeps a line
  // free and holds its lines with the space after each paragraph, and each of the twelve paragraphs
  // a writer may write may end a character into a line.
  const held = textSlot(planned(demo, 4350, { dense: true }));
  assert.deepEqual([held.fontSize, (held.maxChars ?? 0) >= 4350], [9, true]);
  assert.ok((textSlot(planned(demo, 4740, { dense: true })).maxChars ?? 0) < 4740);
  // A text over what the slide holds at the least size that reads is planned at that size, and
  // its budget is what the slide holds there in any number of paragraphs up to 12.
  const floor = readableSize(9144000);
  const most = textSlot(planned(rtl, 3600, { dense: true }));
  assert.equal(most.fontSize, floor);
  assert.ok((most.maxChars ?? 0) < 3600);
  for (const count of counts) {
    const slide = slideOf(textOf(most.maxChars ?? 0, SENTENCES.he, count), { dense: true });
    assert.deepEqual(issuesOf(rtl, 'rtl', slide), [], `${most.maxChars} in ${count} paragraphs`);
    assert.ok((textSlot(resolved(freeFor(rtl, {}, slide)).plan).fontSize ?? 0) >= floor);
  }
});

test('a writer inside its budget and paragraphs is not sent back, however its paragraphs fall', () => {
  // Planned for the 32 paragraphs of a document that its slide does not hold as it is, a slide is
  // told a budget and a number of paragraphs. Where the slide holds the text's characters only in
  // twelve or fewer, the text is cut and told twelve; where it holds them in more, the most it
  // holds them in. A text inside both passes the deck's check however its paragraphs fall: alike,
  // all short but one long one, or all just over a line but one. Each may lose up to a line where
  // it ends, and its budget leaves room for that.
  for (const [name, template, sentence, held] of [
    ['rtl', rtl, SENTENCES.he, 1200],
    ['demo', demo, SENTENCES.en, 3000],
  ] as const) {
    for (const length of [8000, held]) {
      const plan = resolved(freeFor(template, {}, { blocks: ['text'], slots: {}, arrange: { dense: true } }, { icons: false, lengths: { b1Text: length }, paragraphs: { b1Text: 32 } })).plan;
      const slot = textSlot(plan);
      const budget = slot.maxChars ?? 0;
      const told = slot.maxParagraphs ?? 0;
      if (length === held) assert.ok(told > 12 && told < 32 && budget >= held, `${name}: ${budget} in ${told}`);
      else assert.ok(told === 12 && budget < length, `${name}: ${budget} in ${told}`);
      for (const short of [20, 40, 50, 60, 70, 90, 120]) {
        const shorts: string[] = Array.from({ length: told - 1 }, (_, at) => textOf(short, sentence.slice(at % 9) + sentence, 1)[0] as string);
        const used: number = shorts.reduce((sum, paragraph) => sum + paragraph.length, 0);
        if (used >= budget) continue;
        const shaped = [...shorts, ...textOf(budget - used, sentence, 1)];
        assert.deepEqual(issuesOf(template, name, slideOf(shaped, { dense: true })), [], `${name} ${length}: ${told - 1} of ${short} and one of ${budget - used}`);
      }
      assert.deepEqual(issuesOf(template, name, slideOf(textOf(budget, sentence, told), { dense: true })), [], `${name} ${length}: ${budget} in ${told} alike`);
      // Over it, the deck says what the slide holds in as many paragraphs as the text has: never
      // that it takes fewer paragraphs than its writer was told.
      const over = issuesOf(template, name, slideOf(textOf(2 * budget, sentence, told), { dense: true }));
      assert.ok(over.some((issue) => /characters; the slot fits about/.test(issue.message)), JSON.stringify(over));
      assert.ok(!over.some((issue) => /paragraphs; the slot fits/.test(issue.message)), JSON.stringify(over));
    }
  }
});

test('a dense text is accepted only where 0.6.2 accepted it', () => {
  // The longest text a dense slide took without a length warning in 0.6.2, in 3 to 64 paragraphs
  // alike: one judge counts by length as 0.6.2 did, keeps a line free in each column and the space
  // after each paragraph, and accepts no longer text. At 80 and 100 paragraphs 0.6.2 took none,
  // and neither does it.
  const limits = {
    demo: { 3: 5500, 8: 5220, 16: 5090, 24: 3910, 32: 3480, 40: 4320, 48: 2620, 53: 2886, 64: 3458, 80: 0, 100: 0 },
    rtl: { 3: 2000, 8: 1920, 16: 1690, 24: 1310, 32: 890, 40: 1100, 48: 1300, 53: 1286, 64: 0, 80: 0, 100: 0 },
  } as const;
  for (const [name, template, sentence] of [
    ['demo', demo, SENTENCES.en],
    ['rtl', rtl, SENTENCES.he],
  ] as const) {
    for (const [count, most] of Object.entries(limits[name])) {
      // Where 0.6.2 took none, from ten characters a paragraph.
      const from = most || 10 * Number(count);
      for (const more of [50, 500, 1500]) {
        const issues = issuesOf(template, name, slideOf(textOf(from + more, sentence, Number(count)), { dense: true }));
        assert.ok(issues.some((issue) => /characters; the slot fits about|paragraphs; the slot fits/.test(issue.message)), `${name} ${count}: ${from + more}`);
      }
    }
  }
});

test('a dense text of more paragraphs than its box holds is always told so', () => {
  // Once written, a text that does not fit is told so: by its characters, in as many paragraphs as
  // it has, or by its paragraphs where it has more than leave half its room for its lines. Told the
  // characters of fewer paragraphs, 120 paragraphs of ten characters passed with nothing said and
  // ran far under their box. At 80 paragraphs and more, as in 0.6.2, no such text is accepted.
  for (const [name, template, sentence] of [
    ['demo', demo, SENTENCES.en],
    ['rtl', rtl, SENTENCES.he],
  ] as const) {
    for (const count of [80, 100, 120]) {
      for (const chars of [10, 20, 27, 30, 45]) {
        const issues = issuesOf(template, name, slideOf(Array.from({ length: count }, (_, at) => textOf(chars, sentence.slice(at % 9) + sentence, 1)[0] as string), { dense: true }));
        assert.ok(
          issues.some((issue) => /characters; the slot fits about|paragraphs; the slot fits/.test(issue.message)),
          `${name}: ${count} of ${chars}`,
        );
      }
    }
  }
});

test('a slide that opens on a title line is told about as much as one without it', () => {
  // Counted as paragraphs all as long as the next, 32 paragraphs of just over a line took two lines
  // each, where 30 fitted: the budget fell by almost half. Counted also as of any length, it does not.
  const budget = (template: typeof demo, count: number) =>
    textSlot(resolved(freeFor(template, {}, { blocks: ['text'], slots: {}, arrange: { dense: true } }, { icons: false, lengths: { b1Text: 8000 }, paragraphs: { b1Text: count } })).plan).maxChars ?? 0;
  for (const template of [rtl, demo]) {
    for (const count of [24, 28, 30, 32, 34, 36]) assert.ok(budget(template, count) >= 0.75 * budget(template, count - 2), `${count}: ${budget(template, count)} after ${budget(template, count - 2)}`);
  }
});

test('a short list set dense that one column does not hold is set in two, whole', () => {
  // A lead line and fifteen points of 30 characters: one column holds them at four fifths of the
  // template's size only cut, and the slide was a third empty. In two columns the slide holds them.
  const lines = ['Notes', ...Array.from({ length: 15 }, (_, at) => `Point ${at + 1}: the bikes are near the dock`)];
  const chars = lines.reduce((sum, line) => sum + line.length, 0);
  assert.ok(chars <= 600);
  const plan = resolved(freeFor(demo, {}, { blocks: ['text'], slots: {}, arrange: { dense: true } }, { icons: false, lengths: { b1Text: chars }, paragraphs: { b1Text: lines.length } })).plan;
  assert.equal(textSlot(plan).drawn.columns, 2);
  assert.ok((textSlot(plan).maxChars ?? 0) >= chars, String(textSlot(plan).maxChars));
  assert.deepEqual(issuesOf(demo, 'demo', slideOf(lines, { dense: true })), []);
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
  assert.ok(twelve && twelve.densest >= 4200 && twelve.densest <= 4500, String(twelve?.densest));
  assert.equal(twelve?.sizes.densest, 9);
  // Without a number of paragraphs, a text set dense is written as a writer told a length writes
  // it, in any number up to 12 of any length, and holds what twelve hold: `densest` is the most a
  // slide can be asked for and hold, in steps of 50 characters.
  assert.ok(twelve && room.densest === twelve.densest, `${room.densest} of ${twelve?.densest}`);
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
