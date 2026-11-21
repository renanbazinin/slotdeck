// What a writer is told a layout's slot of paragraphs holds counts the space its template keeps
// around paragraphs (src/template/told.ts, 0.6.8): the lines its box holds at that number of
// paragraphs, each paragraph counted in whole lines. The deck's check stays at the stored budget.
import assert from 'node:assert/strict';
import { readFileSync } from 'node:fs';
import { dirname, join } from 'node:path';
import { test } from 'node:test';
import { fileURLToPath } from 'node:url';
import { strFromU8, strToU8, unzipSync, zipSync } from 'fflate';
import { TOOL_NAMES, loadTemplate, memoryStore, planEdit, planOutline, planTextFit, presentationTools, renderDeck, validateDeck, writeDeck } from '../src/index.ts';
import type { Brief, DeckSpec, JsonRequest, LlmProvider, TemplateManifest } from '../src/index.ts';
import { FREE_PATTERN } from '../src/deck/spec.ts';
import { buildCatalog, describeCatalog, slotNote } from '../src/plan/catalog.ts';
import type { CatalogContext, UsablePattern } from '../src/plan/catalog.ts';
import { documentRoom, holdsAsIs } from '../src/plan/length.ts';
import { effectiveSlots, textBudget } from '../src/template/manifest.ts';
import type { PatternDef, TextSlotDef } from '../src/template/manifest.ts';
import { mostHeld, toldCounter, toldFor, toldPair } from '../src/template/told.ts';
import type { SlotCount } from '../src/template/told.ts';

const root = join(dirname(fileURLToPath(import.meta.url)), '..');
const demoManifest = JSON.parse(readFileSync(join(root, 'templates', 'demo.manifest.json'), 'utf8')) as TemplateManifest;
const demoPptx = new Uint8Array(readFileSync(join(root, 'templates', 'demo.pptx')));
const demo = loadTemplate(demoPptx, demoManifest);
const fields = { organizationDepartment: 'Retail' };

const patternOf = (manifest: TemplateManifest, id: string): PatternDef => manifest.patterns.find((pattern) => pattern.id === id) as PatternDef;
const slotOf = (pattern: PatternDef, name: string, items?: number): TextSlotDef => effectiveSlots(pattern, items).find((slot) => slot.name === name) as TextSlotDef;
/** The demo with a change to the shape of the body of its text-only slide. */
const bodyChanged = (change: (shape: string) => string): Uint8Array => {
  const files = unzipSync(new Uint8Array(demoPptx));
  const name = 'ppt/slides/slide2.xml';
  const xml = strFromU8(files[name] as Uint8Array);
  const start = xml.indexOf('<p:cNvPr id="7" name="Body"/>');
  const end = xml.indexOf('</p:sp>', start);
  assert.ok(start > 0 && end > start);
  files[name] = strToU8(xml.slice(0, start) + change(xml.slice(start, end)) + xml.slice(end));
  return zipSync(files);
};
/** A count made up for the pair's rule: a stored budget and what any text of 1, 2, ... paragraphs holds. */
const made = (maxChars: number, curve: number[]): SlotCount => {
  const count: SlotCount = { stored: { maxChars, maxParagraphs: curve.length }, curve, maxChars, maxParagraphs: curve.length, holds: () => false };
  return { ...count, ...toldPair(count) };
};

test("the count of a slot of paragraphs reads the space its template keeps around them, and speaks only where there is some", () => {
  const counter = toldCounter(demo);
  const textOnly = patternOf(demoManifest, 'text-only');
  const body = slotOf(textOnly, 'body');
  // The body keeps 12 points after each paragraph: three paragraphs of any lengths hold 263 of its
  // 331 characters, one or two hold all of them. The writer is told 263 in three.
  const count = counter(textOnly, body, 1) as SlotCount;
  assert.deepEqual([count.curve, count.stored, count.maxChars, count.maxParagraphs], [[331, 331, 263], { maxChars: 331, maxParagraphs: 3 }, 263, 3]);
  const pair = (id: string, name: string) => {
    const found = counter(patternOf(demoManifest, id), slotOf(patternOf(demoManifest, id), name), 1);
    return found && [found.maxChars, found.maxParagraphs];
  };
  // A list of four that would lose half its characters in four is told three; columns of three, two.
  assert.deepEqual(pair('list-with-photo', 'bulletList'), [244, 3]);
  assert.deepEqual(pair('two-columns', 'col1Text'), [231, 2]);

  // Where the shape says so (`spcFirstLastPara`), PowerPoint draws the first paragraph's space
  // before and the last one's after too: two paragraphs hold less.
  const firstLast = loadTemplate(bodyChanged((shape) => shape.replace('<a:bodyPr', '<a:bodyPr spcFirstLastPara="1"')), demoManifest);
  assert.deepEqual(toldCounter(firstLast)(textOnly, body, 1)?.curve, [331, 316, 263]);
  // So is the space before each paragraph, drawn over the first only where the shape says so.
  const before = (shape: string) => shape.replaceAll('<a:pPr algn="l">', '<a:pPr algn="l"><a:spcBef><a:spcPts val="1200"/></a:spcBef>');
  const curveOf = (file: Uint8Array) => toldCounter(loadTemplate(file, demoManifest))(textOnly, body, 1)?.curve;
  assert.deepEqual([curveOf(bodyChanged(before)), curveOf(bodyChanged((shape) => before(shape).replace('<a:bodyPr', '<a:bodyPr spcFirstLastPara="1"')))], [
    [331, 316, 210],
    [331, 263, 158],
  ]);

  // No space between paragraphs and lines 1.2 times their size apart: the stored budget stands.
  assert.equal(counter(textOnly, { ...body, paragraphGap: 0 }, 1), undefined);
  // Nor does it speak for a free slide's slots, a slot styled after another, a deck's own field, or a slot of one paragraph.
  assert.equal(counter({ ...textOnly, id: FREE_PATTERN }, body, 1), undefined);
  assert.equal(counter(textOnly, { ...body, styleOf: { pattern: 'text-only', slot: 'body' } } as unknown as TextSlotDef, 1), undefined);
  assert.equal(counter(textOnly, { ...body, deckField: 'organizationDepartment' }, 1), undefined);
  assert.equal(counter(textOnly, slotOf(textOnly, 'title'), 1), undefined);
  // Kept per template and manifest.
  assert.equal(toldCounter(demo), counter);
});

test('the count follows the box: its size at the text size, its line spacing, its columns, and a line never wider than the stated width', () => {
  const textOnly = patternOf(demoManifest, 'text-only');
  /** The count of the text-only body of a template file, with the body slot changed. */
  const countOf = (file: Uint8Array, manifest: TemplateManifest, scale: number, change: (slot: TextSlotDef) => void = () => undefined) => {
    const changed = structuredClone(manifest);
    const pattern = patternOf(changed, 'text-only');
    change(slotOf(pattern, 'body'));
    return toldCounter(loadTemplate(file, changed))(pattern, slotOf(pattern, 'body'), scale);
  };
  const pairOf = (count: SlotCount | undefined) => count && [count.maxChars, count.maxParagraphs, ...count.curve];
  // A slot that scales with the deck's text size is counted at that size, as its stored budget is.
  const scaled = [0.8, 1, 1.25, 1.5].map((scale) => pairOf(countOf(demoPptx, demoManifest, scale, (slot) => (slot.scales = true))));
  assert.deepEqual(scaled, [
    [464, 3, 517, 517, 464],
    [263, 3, 331, 331, 263],
    [211, 2, 211, 211],
    [139, 2, 147, 139],
  ]);
  // Lines set 1.5 times apart speak for the count with no space between paragraphs.
  const spaced = bodyChanged((shape) => shape.replaceAll('<a:pPr algn="l">', '<a:pPr algn="l"><a:lnSpc><a:spcPct val="150000"/></a:lnSpc>'));
  assert.deepEqual(pairOf(countOf(spaced, demoManifest, 1, (slot) => (slot.paragraphGap = 0))), [210, 2, 263, 210, 158]);
  // In two columns, each narrower, the second losing a line and a paragraph's spacing where the text breaks into it.
  const columns = bodyChanged((shape) => shape.replace('<a:bodyPr', '<a:bodyPr numCol="2" spcCol="457200"'));
  assert.deepEqual(pairOf(countOf(columns, demoManifest, 1)), [295, 3, 331, 331, 295]);
  // A template wider than the estimate is counted at its own width on a short line too: the count never rises as the box narrows.
  for (const characterWidth of [0.5, 0.7, 0.72]) {
    let last = Number.POSITIVE_INFINITY;
    for (let w = 0.4; w >= 0.18; w -= 0.01) {
      const held = countOf(demoPptx, { ...demoManifest, characterWidth }, 1, (slot) => (slot.box = { ...(slot.box as NonNullable<TextSlotDef['box']>), w }))?.curve[0] ?? 0;
      assert.ok(held <= last, `${characterWidth} at ${w.toFixed(2)}: ${held} after ${last}`);
      last = held;
    }
  }
  // What code sets as it is: each paragraph in its own whole lines, never more than the stored budget or its paragraphs.
  const count = toldCounter(demo)(textOnly, slotOf(textOnly, 'body'), 1) as SlotCount;
  assert.deepEqual(
    [count.holds([331]), count.holds([332]), count.holds([]), count.holds([10, 10, 10, 10]), count.holds([96, 93, 97]), count.holds([200, 65, 66])],
    [true, false, false, false, true, false],
  );
});

test('the pair keeps the paragraphs where that costs few characters, else gives fewer and longer, and never tells more than the stored budget', () => {
  // A small loss: three paragraphs, 280 characters.
  assert.deepEqual(toldPair(made(300, [300, 290, 280])), { maxChars: 280, maxParagraphs: 3 });
  // Half the characters lost in three: two paragraphs, all 300.
  assert.deepEqual(toldPair(made(300, [300, 300, 150])), { maxChars: 300, maxParagraphs: 2 });
  // Where the count holds the stored budget, the stored budget.
  assert.deepEqual(toldPair(made(300, [300, 300, 300])), { maxChars: 300, maxParagraphs: 3 });
  // A text of known paragraphs: what that many hold, and more paragraphs where they hold as much.
  const flat = made(300, [300, 280, 280]);
  assert.deepEqual(toldPair(flat, 1), { maxChars: 300, maxParagraphs: 1 });
  assert.deepEqual(toldPair(flat, 2), { maxChars: 280, maxParagraphs: 3 });

  // Of equal losses, the one of more characters: one paragraph of 300 before two of 150.
  assert.deepEqual(toldPair(made(300, [300, 150])), { maxChars: 300, maxParagraphs: 1 });
  // A text of one paragraph where the count holds the stored budget in any: all the paragraphs it may take.
  assert.deepEqual(toldPair(made(300, [300, 300, 300]), 1), { maxChars: 300, maxParagraphs: 3 });

  // An asked length, whose writer chooses its paragraphs: the pair where it holds the length, else
  // the most paragraphs that hold it, down to half the pair's, else the pair. What it holds of a
  // length is the most of those.
  const small = made(300, [300, 290, 280]);
  assert.deepEqual(toldFor(small, 250), { maxChars: 280, maxParagraphs: 3 });
  assert.deepEqual(toldFor(small, 290), { maxChars: 290, maxParagraphs: 2 });
  assert.deepEqual(toldFor(small, 295), { maxChars: 280, maxParagraphs: 3 });
  assert.deepEqual(toldFor(small, 310), { maxChars: 280, maxParagraphs: 3 });
  assert.equal(mostHeld(small), 290);
  // The demo's list beside a photo, told 244 in 3: a length of 300 in two items, and none of 330 in one item: a list stays a list.
  const listed = patternOf(demoManifest, 'list-with-photo');
  const list = toldCounter(demo)(listed, slotOf(listed, 'bulletList'), 1) as SlotCount;
  assert.deepEqual([toldFor(list, 300), toldFor(list, 330), mostHeld(list)], [{ maxChars: 306, maxParagraphs: 2 }, { maxChars: 244, maxParagraphs: 3 }, 306]);

  // On the demo, at every text size and number of paragraphs: never more than the stored budget,
  // and what a text holds never rises with its paragraphs.
  const counter = toldCounter(demo);
  for (const scale of [0.8, 1, 1.25, 1.5]) {
    for (const pattern of demoManifest.patterns) {
      for (const slot of effectiveSlots(pattern)) {
        if (slot.kind !== 'text') continue;
        const count = counter(pattern, slot, scale);
        if (!count) continue;
        const stored = textBudget(slot, scale);
        assert.deepEqual(count.stored, { maxChars: stored.maxChars, maxParagraphs: stored.maxParagraphs }, `${pattern.id}.${slot.name}`);
        for (let at = 0; at < count.curve.length; at++) {
          assert.ok((count.curve[at] as number) <= count.stored.maxChars, `${pattern.id}.${slot.name} ${at + 1}`);
          if (at > 0) assert.ok((count.curve[at] as number) <= (count.curve[at - 1] as number), `${pattern.id}.${slot.name} ${at + 1}`);
          const told = toldPair(count, at + 1);
          assert.ok(told.maxChars <= count.stored.maxChars && told.maxParagraphs <= count.stored.maxParagraphs);
        }
      }
    }
  }
});

interface Measured {
  pattern: string;
  slot: string;
  items?: number;
  textScale: number;
  each?: number[];
  chars?: number;
  paragraphs?: number;
  /** Points PowerPoint set the text past its box; negative where there was room left. */
  over: number;
}
const measured = (JSON.parse(readFileSync(join(root, 'test', 'fixtures', 'powerpoint-spacing.json'), 'utf8')) as { texts: Measured[] }).texts;

test('PowerPoint sets inside its box every measured text that the count says a slot holds', () => {
  for (const characterWidth of [undefined, 0.43]) {
    const template = loadTemplate(demoPptx, characterWidth === undefined ? demoManifest : { ...demoManifest, characterWidth });
    const counter = toldCounter(template);
    const seen = { pair: 0, curve: 0, asIs: 0, storedOver: 0 };
    for (const text of measured) {
      const pattern = patternOf(template.manifest, text.pattern);
      const slot = slotOf(pattern, text.slot, text.items);
      const scale = text.textScale * (template.manifest.textScale ?? 1);
      const chars = text.each ? text.each.reduce((sum, one) => sum + one, 0) : (text.chars as number);
      const paragraphs = text.each ? text.each.length : (text.paragraphs as number);
      const over = text.over > 0.5;
      const stored = textBudget(slot, scale);
      if (chars <= (stored.maxChars as number) && paragraphs <= (stored.maxParagraphs ?? 1) && over) seen.storedOver++;
      const count = counter(pattern, slot, scale, text.items);
      assert.ok(count, `${text.pattern}.${text.slot}`);
      const label = `${text.pattern}.${text.slot} at ${text.textScale}, ${characterWidth ?? 'the estimate'}: ${text.each ? `[${text.each.join(', ')}]` : `${chars} in ${paragraphs}`} ran ${text.over} pt over`;
      // The pair the writer is told, what any text of its own paragraphs holds, and its own lengths counted line by line.
      if (chars <= count.maxChars && paragraphs <= count.maxParagraphs) {
        seen.pair++;
        assert.ok(!over, `told ${count.maxChars}/${count.maxParagraphs}: ${label}`);
      }
      if (paragraphs <= count.stored.maxParagraphs && chars <= (count.curve[paragraphs - 1] as number)) {
        seen.curve++;
        assert.ok(!over, `held in ${paragraphs}: ${label}`);
      }
      if (text.each && count.holds(text.each)) {
        seen.asIs++;
        assert.ok(!over, `held as it is: ${label}`);
      }
    }
    // The fixture is not empty of what the count is for: texts inside the stored budget that ran over.
    assert.ok(seen.storedOver >= 20 && seen.pair >= 40 && seen.curve >= 150 && seen.asIs >= 200, JSON.stringify(seen));
  }
});

const context = (told: boolean): CatalogContext => ({ layout: 'adaptive', fields, icons: [], assets: [], imageDensity: 'none', figures: false, textScale: 1, ...(told ? { told: toldCounter(demo) } : {}) });

test("the writer is told the pair, without a length its description states where the pair is less, and the deck is checked against the stored budget", () => {
  const counted = describeCatalog(buildCatalog(demoManifest, context(true)));
  const stored = describeCatalog(buildCatalog(demoManifest, context(false)));
  assert.match(counted, /body \(at most 263 characters in total, at most 3 items\)/);
  assert.match(counted, /colNText \(at most 231 characters in total, at most 2 items\)/);
  assert.match(stored, /body \(at most 331 characters in total, at most 3 items\)/);
  assert.match(stored, /colNText \(at most 320 characters in total, at most 3 items\)/);

  // A description that says a length of its own says it of the stored budget: where the pair is less, it is left out.
  const described = structuredClone(demoManifest);
  const body = slotOf(patternOf(described, 'text-only'), 'body');
  body.description = 'About 300 characters in all. Two or three short paragraphs';
  const planned = (told: boolean) => (buildCatalog(described, context(told)).patterns.find((pattern) => pattern.def.id === 'text-only') as UsablePattern).slots.find((one) => one.slot.name === 'body');
  const lowered = planned(true);
  assert.equal(slotNote(lowered?.slot as TextSlotDef, context(true), lowered), 'Two or three short paragraphs. at most 263 characters in total, at most 3 items');
  assert.equal(slotNote(body, context(false), planned(false)), 'About 300 characters in all. Two or three short paragraphs. at most 331 characters in total, at most 3 items');

  // A text of two columns is written, never set as it is: each column is told the pair, whatever its lengths.
  const columns = buildCatalog(demoManifest, context(true)).patterns.find((pattern) => pattern.def.id === 'two-columns') as UsablePattern;
  const asIs = { chars: 150, each: [50, 50, 50] };
  const column = columns.slots.find((planned) => planned.slot.name === 'col1Text') as { told?: SlotCount };
  assert.equal(column.told?.holds(asIs.each), true, 'one column holds the text as it is');
  assert.equal(holdsAsIs(columns, asIs), false);
  assert.equal(documentRoom(columns, context(true), { paragraphs: 3 }, undefined, asIs), documentRoom(columns, context(true), { paragraphs: 3 }));

  // The check is the stored budget: 320 characters in three paragraphs are more than the pair and pass.
  const deck: DeckSpec = { version: 1, title: 'Bikes', fields, slides: [{ pattern: 'text-only', slots: { title: 'Keep it', body: ['A'.repeat(106), 'B'.repeat(107), 'C'.repeat(107)] } }] };
  // Within the stored budget; its words of a hundred letters are told as words no line holds (0.7.0), and that is all.
  assert.deepEqual(validateDeck(demo, deck, { layout: 'adaptive' }).filter((issue) => issue.slot === 'body').map((issue) => issue.code), ['word-wider-than-line', 'word-wider-than-line', 'word-wider-than-line']);
});

/** A model that gives the same outline to every call. */
const outlineModel = (answer: unknown): LlmProvider => ({ generateJson: async () => answer });
const brief: Brief = { prompt: 'Why the city should keep its bike share', language: 'en', fields, imageDensity: 'none', slideCount: 1 };

test('countSpacing: false tells every slot its stored budget, as 0.6.7 did: in planning, what a plan holds, and the chat tools', async () => {
  const outline = { title: 'Bikes', slides: [{ pattern: 'text-only', purpose: 'Keep the bike share', shows: [], shows_main: 'auto', side_by_side: 0, words: 200 }] };
  const plan = await planOutline(demo, brief, { llm: outlineModel(outline), layout: 'adaptive', free: false });
  // Counted, two paragraphs hold all 331 characters (55 words); not counted, nine tenths of them (49).
  assert.deepEqual(planTextFit(demo, plan, { layout: 'adaptive', free: false }), [{ slide: 1, words: 200, holds: 55 }]);
  assert.deepEqual(planTextFit(demo, plan, { layout: 'adaptive', free: false, countSpacing: false }), [{ slide: 1, words: 200, holds: 49 }]);
  // The writer of two columns is told two paragraphs a column, as many as hold most; without the count, the stored three.
  for (const [countSpacing, most, chars] of [[undefined, 2, 231], [false, 3, 320]] as const) {
    const requests: JsonRequest[] = [];
    const llm: LlmProvider = {
      generateJson: async (request) => {
        requests.push(request);
        if (requests.length === 1) return { title: 'Bikes', slides: [{ pattern: 'two-columns', purpose: 'Before and after', shows: [], shows_main: 'auto', side_by_side: 0 }] };
        const slide = { title: 'Before and after', col1Heading: 'Before', col2Heading: 'After', col1Text: ['Cars.'], col2Text: ['Bikes.'] };
        return { s1: slide };
      },
    };
    const columns = await planOutline(demo, brief, { llm, layout: 'adaptive', ...(countSpacing === undefined ? {} : { countSpacing }) });
    await writeDeck(demo, columns, { llm, layout: 'adaptive', ...(countSpacing === undefined ? {} : { countSpacing }) });
    const text = ((requests[1]?.schema as unknown as { properties: Record<string, { properties: Record<string, { maxItems?: number; description?: string }> }> }).properties.s1?.properties.col1Text);
    assert.equal(text?.maxItems, most, String(countSpacing));
    assert.match(String(text?.description), new RegExp(`at most ${chars} characters in total, at most ${most} items$`));
  }

  // A change to a deck that asks for an amount of text: the room said and the writer's limits.
  const deck: DeckSpec = { version: 1, language: 'en', fields, title: 'Bikes', slides: [{ pattern: 'text-only', slots: { title: 'Keep it', body: ['Riders rely on it.'] } }] };
  const change = { summary: 'Longer.', language: 'en', title: 'Bikes', slides: [{ from: 1, pattern: 'text-only', change: 'Make the text 200 words', images: false, words: 200, parts: [], shows_main: 'auto', side_by_side: 0 }], text_size: 'normal', fields };
  for (const [countSpacing, holds, items] of [[undefined, 55, 2], [false, 49, 3]] as const) {
    const requests: JsonRequest[] = [];
    const answers: unknown[] = [change, { s1: { title: 'Keep it', body: ['Riders rely on it.'] } }];
    const llm: LlmProvider = { generateJson: async (request) => (requests.push(request), answers.shift()) };
    const edited = await planEdit(demo, deck, 'Make slide 1 200 words', { llm, layout: 'adaptive', free: false, ...(countSpacing === undefined ? {} : { countSpacing }) });
    assert.ok(edited.warnings.some((issue) => issue.message.startsWith(`200 words were asked for this slide, and it holds about ${holds}:`)), JSON.stringify(edited.warnings));
    const body = (requests[1]?.schema as unknown as { properties: Record<string, { properties: Record<string, { maxItems?: number }> }> }).properties.s1?.properties.body;
    assert.equal(body?.maxItems, items, String(countSpacing));
  }

  for (const [countSpacing, holds, body] of [[undefined, 55, 263], [false, 49, 331]] as const) {
    let next = 0;
    const prompts: string[] = [];
    const llm: LlmProvider = { generateJson: async (request) => (prompts.push(request.prompt), outline) };
    const tools = presentationTools({ template: demo, llm, store: memoryStore(), fields, layout: 'adaptive', free: false, newId: (kind) => `${kind}_${++next}`, ...(countSpacing === undefined ? {} : { countSpacing }) });
    const planned = await tools.call(TOOL_NAMES.plan, { topic: 'Why the city should keep its bike share', language: 'en', slide_count: 1 });
    // The planner's list of layouts, and what the plan is said to hold.
    assert.match(prompts[0] ?? '', new RegExp(`body \\(at most ${body} characters in total, at most 3 items\\)`), String(countSpacing));
    assert.match(String(planned.output.note), new RegExp(`slide 1 holds about ${holds} of its 200 words`), String(countSpacing));
  }
});

test('with fonts, a text in columns is broken at its column less the space between columns, and the space at its ends is counted where the shape says so', async () => {
  const fontkit = await import('fontkit');
  const { fontkitMetrics, textMeasurer } = await import('../src/index.ts');
  const { openFonts } = await import('./fonts.ts');
  const fonts = fontkitMetrics(fontkit, openFonts());
  const textOnly = patternOf(demoManifest, 'text-only');
  const body = slotOf(textOnly, 'body');
  const fitIn = (pptx: Uint8Array, text: string[]) => textMeasurer(loadTemplate(pptx, demoManifest, { fonts }), fonts).fit(textOnly, body, text);
  const three = ['Riders rely on the service every day.', 'The city gains from it.', 'Both are easy to see.'];
  // Two columns half an inch apart: each line is a quarter of an inch narrower than half the box (0.6.10's count left the gap in).
  const plain = fitIn(bodyChanged((shape) => shape.replace('<a:bodyPr', '<a:bodyPr numCol="2"')), three);
  const apart = fitIn(bodyChanged((shape) => shape.replace('<a:bodyPr', '<a:bodyPr numCol="2" spcCol="457200"')), three);
  assert.ok(plain && apart);
  assert.ok(Math.abs((plain.lineRoom[0] ?? 0) - (apart.lineRoom[0] ?? 0) - 18) < 0.01, `${plain.lineRoom[0]} and ${apart.lineRoom[0]}`);
  // The body keeps 12 points after each paragraph: with spcFirstLastPara the last one's is drawn too.
  const ends = fitIn(bodyChanged((shape) => shape.replace('<a:bodyPr', '<a:bodyPr spcFirstLastPara="1"')), three);
  const usual = fitIn(demoPptx, three);
  assert.ok(ends && usual);
  assert.equal(Math.round((ends.height - usual.height) * 100) / 100, 12);
});

test('a free slide drawn from a shape that draws the space at the ends of its text does not draw it: its plan does not count it', () => {
  const pptx = bodyChanged((shape) => shape.replace('<a:bodyPr', '<a:bodyPr spcFirstLastPara="1"'));
  assert.ok(strFromU8(unzipSync(pptx)['ppt/slides/slide2.xml'] as Uint8Array).includes('spcFirstLastPara="1"'));
  const template = loadTemplate(pptx, { ...demoManifest, free: { styles: { body: { slide: 2, shape: '7' } } } });
  const deck: DeckSpec = { version: 1, fields, slides: [{ pattern: 'free', blocks: ['text'], slots: { title: 'A title', b1Text: ['One paragraph.', 'And another one.'] } }] };
  const slide = strFromU8(unzipSync(renderDeck(template, deck, { layout: 'adaptive' }).pptx)['ppt/slides/slide1.xml'] as Uint8Array);
  assert.ok(slide.includes('One paragraph.') && !slide.includes('spcFirstLastPara'));
});
