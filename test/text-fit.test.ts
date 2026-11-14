// Measuring text with font metrics: the line breaker, the measurer on the demo template, and
// everything that uses it (validation, the renderer's shrinking, the planner's repair round, the
// studio's checks). Most tests use a made-up font in which every character is half an em wide,
// so the expected numbers can be worked out by hand; the last ones use real fonts and compare
// with line counts recorded from PowerPoint.
import assert from 'node:assert/strict';
import { readFileSync } from 'node:fs';
import { dirname, join } from 'node:path';
import { test } from 'node:test';
import { fileURLToPath } from 'node:url';
import { strFromU8, strToU8, unzipSync, zipSync } from 'fflate';
import * as fontkit from 'fontkit';
import {
  checkManifest,
  fontkitMetrics,
  loadTemplate,
  measuredBudgets,
  planDeck,
  renderDeck,
  sampleDeck,
  templateFonts,
  textMeasurer,
  validateDeck,
} from '../src/index.ts';
import type { DeckSpec, FontFace, FontMetrics, JsonRequest, TemplateManifest } from '../src/index.ts';
import { bodyTextScale, effectiveSlots, slotSources } from '../src/template/manifest.ts';
import { textWidth } from '../src/text/fonts.ts';
import { breakLines } from '../src/text/layout.ts';
import { officeFonts, openFonts } from './fonts.ts';

const root = join(dirname(fileURLToPath(import.meta.url)), '..');
const pptx = readFileSync(join(root, 'templates', 'demo.pptx'));
const manifest = JSON.parse(readFileSync(join(root, 'templates', 'demo.manifest.json'), 'utf8')) as TemplateManifest;

/** Every character half an em wide, in every typeface: n characters at s points are n * s / 2 points wide. */
const fixedWidth: FontMetrics = { face: () => ({ has: () => true, advances: (text) => [...text].map(() => 0.5) }) };
/** The width of a text in ems, as a face gives it glyph by glyph. */
const ems = (face: FontFace | undefined, text: string, kerning: boolean): number => (face?.advances(text, kerning) ?? []).reduce((sum, advance) => sum + advance, 0);
const template = loadTemplate(pptx, manifest, { fonts: fixedWidth });
const measure = textMeasurer(template);
// The same template with real fonts: open ones that have the metrics of the template's own.
const realTemplate = loadTemplate(pptx, manifest, { fonts: fontkitMetrics(fontkit, openFonts()) });
/** What the made-up font makes of the body of "text-only": the lines its box holds, the characters on a line, its size. */
const body = {
  fits: measure.fit('text-only', 'body', 'a')?.capacity ?? 0,
  perLine: measure.fit('text-only', 'body', 'a'.repeat(400))?.paragraphs[0]?.[0]?.length ?? 0,
  size: measure.fit('text-only', 'body', 'a')?.size ?? 0,
};

/** Words of four letters: with the space after it, each takes five characters. */
const words = (count: number) => Array.from({ length: count }, () => 'word').join(' ');
const slideXml = (bytes: Uint8Array, number = 1) => strFromU8(unzipSync(bytes)[`ppt/slides/slide${number}.xml`] as Uint8Array);

// --- the line breaker -------------------------------------------------------------------------

const chars = (text: string) => [...text].length;

test('words go on a line while they fit, and the space at the end of a line does not count', () => {
  // Ten characters fit: "aaaa bbbb" is nine, and its trailing space hangs outside.
  assert.deepEqual(breakLines('aaaa bbbb cccc dddd', 10, 10, chars), ['aaaa bbbb ', 'cccc dddd']);
  assert.deepEqual(breakLines('aaaa bbbbb cccc', 10, 10, chars), ['aaaa bbbbb ', 'cccc']);
  assert.deepEqual(breakLines('one', 10, 10, chars), ['one']);
  // The first line may be narrower or wider than the rest: an indent, or a hanging one.
  assert.deepEqual(breakLines('aaaa bbbb cccc dddd', 5, 10, chars), ['aaaa ', 'bbbb cccc ', 'dddd']);
});

test('a line breaks after a hyphen or a dash, and not at a slash, a full stop or a Hebrew maqaf', () => {
  assert.deepEqual(breakLines('aaaaaa-bbbbbb-cccccc', 14, 14, chars), ['aaaaaa-bbbbbb-', 'cccccc']);
  assert.deepEqual(breakLines('aaaa–bbbb—cccc', 10, 10, chars), ['aaaa–bbbb—', 'cccc']);
  // No break is allowed here, so the word is cut where the line ends, as PowerPoint cuts it.
  assert.deepEqual(breakLines('aaaaaa/bbbbbb/cccccc', 14, 14, chars), ['aaaaaa/bbbbbb/', 'cccccc']);
  assert.deepEqual(breakLines('aaaaaa/bbbbbb/cccccc', 10, 10, chars), ['aaaaaa/bbb', 'bbb/cccccc']);
  assert.deepEqual(breakLines('אאאא־בבבב', 6, 6, chars), ['אאאא־ב', 'בבב']);
});

test('a word longer than a line is cut by characters, and Chinese or Japanese text breaks anywhere', () => {
  assert.deepEqual(breakLines('ab abcdefghijklmnopqrstuvwxyz', 10, 10, chars), ['ab ', 'abcdefghij', 'klmnopqrst', 'uvwxyz']);
  assert.deepEqual(breakLines('日本語のテキスト', 3, 3, chars), ['日本語', 'のテキ', 'スト']);
  // Closing punctuation never starts a line: it stays with the character before it.
  assert.deepEqual(breakLines('日本語。日本', 3, 3, chars), ['日本', '語。日', '本']);
});

// --- the measurer -----------------------------------------------------------------------------

test('a text is measured in lines at the slot\'s size, against the room its box gives', () => {
  // The body of "text-and-photo" is 446.4 x 208.8 pt with no inset, set at 18 pt: 9 pt a character
  // in the made-up font, so 49 characters a line (ten words), and lines 21.6 pt apart (1.2 times the size).
  const fit = measure.fit('text-and-photo', 'body', words(30));
  assert.ok(fit);
  assert.equal(fit.size, 18);
  assert.deepEqual(fit.paragraphs.map((lines) => lines.length), [3]);
  assert.equal(fit.paragraphs[0]?.[0], `${words(10)} `);
  assert.equal(fit.height, 64.8);
  assert.equal(fit.room, 208.8);
  assert.deepEqual([fit.fits, fit.overflow, fit.capacity], [true, 0, 9]);

  // Ten lines are one more than the box holds.
  const long = measure.fit('text-and-photo', 'body', words(98));
  assert.deepEqual([long?.lines, long?.fits, long?.overflow, long?.capacity], [10, false, 1, 9]);
  // The hint for whoever shortens the text: what fits is about nine tenths of it.
  assert.ok((long?.fitsChars ?? 0) > 380 && (long?.fitsChars ?? 0) < 450, String(long?.fitsChars));

  // Paragraphs are kept apart by the template's own 12 pt, which takes room that lines would.
  const two = measure.fit('text-and-photo', 'body', [words(30), words(30)]);
  assert.deepEqual([two?.lines, two?.height], [6, 141.6]); // six lines of 21.6 pt and the gap
});

test('larger body text takes more lines, and a title is not body text', () => {
  const normal = measure.fit('text-and-photo', 'body', words(60));
  const large = measure.fit('text-and-photo', 'body', words(60), { textScale: 1.5 });
  assert.deepEqual([normal?.size, large?.size], [18, 18]); // 18 pt is above body text, so the scale leaves it
  const card = measure.fit('icon-cards', 'card1Text', words(20));
  const cardLarge = measure.fit('icon-cards', 'card1Text', words(20), { textScale: 1.5 });
  assert.equal(card?.size, 15);
  assert.equal(cardLarge?.size, 15);
  // The organization line is 12 pt body text, but a deck field: it keeps its size as well.
  assert.equal(measure.fit('text-and-photo', 'organizationDepartment', 'Name', { textScale: 1.5 })?.size, 12);
});

test('text in a rounded box keeps clear of the corners', () => {
  // The takeaway is a rounded rectangle 347.76 pt wide with 18 pt insets; its corners take 3.54 pt more a side.
  // At 17 pt a character is 8.5 pt: 35 fit in 304.7 pt, where the plain box would hold 36.
  const fit = measure.fit('chart-with-takeaway', 'takeawayText', 'a'.repeat(36));
  assert.deepEqual(fit?.paragraphs, [['a'.repeat(35), 'a']]);
});

test('a slot of a copied card is measured like the card it copies', () => {
  const cards = manifest.patterns.find((pattern) => pattern.id === 'icon-cards');
  assert.ok(cards?.repeat);
  const most = cards.repeat.max ?? cards.repeat.shapes.length;
  const last = effectiveSlots(cards, most).find((slot) => slot.name === `card${most}Text`);
  assert.ok(last && last.kind === 'text');
  assert.deepEqual(measure.fit(cards, last, words(20))?.paragraphs, measure.fit('icon-cards', 'card1Text', words(20))?.paragraphs);
  assert.deepEqual(measure.fit('icon-cards', `card${most}Text`, words(20))?.lines, measure.fit('icon-cards', 'card1Text', words(20))?.lines);
});

test('a slot whose typeface was not supplied is not measured, and the template says which fonts it needs', () => {
  assert.deepEqual(
    templateFonts(template).map((font) => `${font.typeface}${font.bold ? ' bold' : ''}${font.italic ? ' italic' : ''}`),
    ['Calibri', 'Calibri italic', 'Calibri bold', 'Cambria italic', 'Cambria bold'],
  );
  const onlyCalibri: FontMetrics = { face: (typeface) => (typeface === 'Calibri' ? fixedWidth.face(typeface, { bold: false, italic: false }) : undefined) };
  const partial = textMeasurer(loadTemplate(pptx, manifest), onlyCalibri);
  assert.ok(partial.fit('text-and-photo', 'body', 'Some text'));
  assert.equal(partial.fit('text-and-photo', 'title', 'A title'), undefined); // set in Cambria
  assert.deepEqual(partial.missingFonts().map((font) => font.typeface), ['Cambria', 'Cambria']);
  // Hebrew in the title is set in the complex-script typeface, which is Calibri: that can be measured.
  assert.ok(partial.fit('text-and-photo', 'title', 'כותרת'));
  assert.throws(() => textMeasurer(loadTemplate(pptx, manifest)), /needs font metrics/);
});

// --- validation -------------------------------------------------------------------------------

const oneSlide = (slots: DeckSpec['slides'][number]['slots'], extra: Partial<DeckSpec> = {}): DeckSpec => ({
  version: 1,
  fields: { organizationDepartment: 'Name' },
  slides: [{ pattern: 'text-only', slots: { title: 'A title', subtitle: 'A subtitle', ...slots } }],
  ...extra,
});

test('with fonts, validation reports overflow in lines and leaves the character budget aside', () => {
  const slot = manifest.patterns.find((pattern) => pattern.id === 'text-only')?.slots.find((entry) => entry.name === 'body');
  assert.ok(slot?.kind === 'text' && slot.maxChars);
  const { fits, perLine, size } = body;
  assert.ok(fits > 3 && perLine > 50 && size >= 18);

  // Longer than the budget, yet it fits: no warning with fonts, the old warning without.
  const roomy = oneSlide({ body: words(Math.ceil(slot.maxChars / 5) + 4) });
  assert.deepEqual(validateDeck(manifest, roomy, { measure }), []);
  assert.match(validateDeck(manifest, roomy)[0]?.message ?? '', /characters; the slot fits about/);

  // Two lines more than the box holds.
  const long = oneSlide({ body: 'a'.repeat(perLine * (fits + 2)) });
  const [issue] = validateDeck(manifest, long, { measure });
  assert.deepEqual([issue?.severity, issue?.slide, issue?.slot], ['warning', 1, 'body']);
  assert.match(issue?.message ?? '', new RegExp(`^text takes ${fits + 2} lines at ${size} pt; the box holds ${fits} \\(2 lines too many\\)\\. Text of about \\d+ characters fits\\.$`));

  // The renderer validates the same way for a template loaded with fonts.
  assert.deepEqual(renderDeck(template, long, { layout: 'adaptive' }).warnings.map((warning) => warning.message), [issue?.message]);
  assert.match(renderDeck(loadTemplate(pptx, manifest), long, { layout: 'adaptive' }).warnings[0]?.message ?? '', /characters; the slot fits about/);
});

// --- shrinking --------------------------------------------------------------------------------

test('text that does not fit is set smaller when the manifest allows it, and no further than it allows', () => {
  const shrinking = loadTemplate(pptx, { ...manifest, shrinkText: { to: 0.8 } }, { fonts: fixedWidth });
  const { fits, perLine, size } = body;
  const text = 'a'.repeat(perLine * (fits + 1));
  // Sizes are whole and half points, so the smallest one allowed is the first at or above four fifths.
  const smallest = Math.ceil(size * 0.8 * 2) / 2;

  // One line too many fits a little smaller; the size goes down by half points.
  const rendered = renderDeck(shrinking, oneSlide({ body: text }), { layout: 'adaptive' });
  assert.deepEqual(rendered.warnings, []);
  assert.equal(rendered.shrunk.length, 1);
  const [shrunk] = rendered.shrunk;
  assert.deepEqual([shrunk?.slide, shrunk?.slot], [1, 'body']);
  assert.ok(shrunk && shrunk.size < size && shrunk.size >= smallest && Number.isInteger(shrunk.size * 2), String(shrunk?.size));
  assert.equal(shrunk.scale, Math.round((shrunk.size / size) * 1000) / 1000);
  const xml = slideXml(rendered.pptx);
  const run = xml.slice(xml.lastIndexOf('<a:r>', xml.indexOf(`<a:t>${text}</a:t>`)));
  assert.equal(/sz="(\d+)"/.exec(run)?.[1], String((shrunk?.size ?? 0) * 100));
  // Text that fits is left at its size, and nothing is shrunk for a template without the setting.
  assert.match(slideXml(rendered.pptx), /sz="3400"[^>]*>(?:<[^>]+>)*<a:t>A title</);
  assert.deepEqual(renderDeck(template, oneSlide({ body: text }), { layout: 'adaptive' }).shrunk, []);

  // Far too long: set at the smallest size allowed, and still reported.
  const far = renderDeck(shrinking, oneSlide({ body: 'a'.repeat(perLine * (fits + 9)) }), { layout: 'adaptive' });
  assert.equal(far.shrunk[0]?.size, smallest);
  assert.match(far.warnings[0]?.message ?? '', new RegExp(`^text takes \\d+ lines even at ${smallest} pt; the box holds \\d+`));

  // A floor in points wins over the factor.
  const floored = loadTemplate(pptx, { ...manifest, shrinkText: { to: 0.5, minSize: size - 1 } }, { fonts: fixedWidth });
  assert.equal(renderDeck(floored, oneSlide({ body: 'a'.repeat(perLine * (fits + 9)) }), { layout: 'adaptive' }).shrunk[0]?.size, size - 1);
});

// --- the planner ------------------------------------------------------------------------------

test('the planner sends a draft back by the lines it runs over, with the length that fits', async () => {
  const { fits, perLine, size } = body;
  const requests: JsonRequest[] = [];
  const answers: unknown[] = [
    { title: 'Deck', slides: [{ pattern: 'text-only', purpose: 'One point' }] },
    { s1: { title: 'A title', subtitle: 'A subtitle', body: ['a'.repeat(perLine * (fits + 3))] } },
    { s1: { title: 'A title', subtitle: 'A subtitle', body: [words(20)] } },
  ];
  const llm = {
    async generateJson(request: JsonRequest) {
      requests.push(request);
      return answers.shift();
    },
  };
  const { deck, warnings } = await planDeck(template, { prompt: 'One point', slideCount: 1, fields: { organizationDepartment: 'Name' } }, { llm, patterns: ['text-only'] });
  assert.equal(requests.length, 3);
  assert.match(requests[2]?.prompt ?? '', new RegExp(`s1\\.body: text takes ${fits + 3} lines at ${size} pt; the box holds ${fits} \\(3 lines too many\\)\\. Text of about \\d+ characters fits\\.`));
  assert.deepEqual(warnings, []);
  assert.deepEqual(deck.slides[0]?.slots.body, [words(20)]);
});

// --- the studio -------------------------------------------------------------------------------

test('checkManifest tries each budget out with the fonts and reports one that is too generous', () => {
  // The demo manifest's budgets were estimated without fonts and are on the safe side: all of them fit.
  assert.deepEqual(checkManifest(realTemplate), []);
  const generous: TemplateManifest = {
    ...manifest,
    patterns: manifest.patterns.map((pattern) =>
      pattern.id !== 'text-only' ? pattern : { ...pattern, slots: pattern.slots.map((slot) => (slot.name === 'body' && slot.kind === 'text' ? { ...slot, maxChars: 3000 } : slot)) },
    ),
  };
  const issues = checkManifest({ ...realTemplate, manifest: generous });
  assert.equal(issues.length, 1);
  assert.deepEqual([issues[0]?.severity, issues[0]?.pattern, issues[0]?.slot], ['warning', 'text-only', 'body']);
  assert.match(issues[0]?.message ?? '', /^has a budget of 3000 characters, but sample text of that length takes \d+ lines where \d+ fit; about \d+ characters fit$/);
  // Without fonts the same manifest passes: nothing measures it.
  assert.deepEqual(checkManifest(loadTemplate(pptx, generous)), []);

  // A typeface that was not supplied is named once, for the whole manifest.
  const missing = checkManifest(loadTemplate(pptx, manifest, { fonts: fontkitMetrics(fontkit, { Calibri: openFonts().Calibri as Uint8Array[] }) }));
  assert.deepEqual(missing.map((issue) => issue.message), ['no font was supplied for Cambria italic, Cambria bold; text set in them is checked by its character budget']);
});

test('measured budgets say what each box holds, and a sample deck can be filled to that', () => {
  const budgets = measuredBudgets(realTemplate);
  const measured = budgets.find((budget) => budget.pattern === 'text-only' && budget.slot === 'body');
  assert.ok(measured && measured.maxChars !== undefined);
  assert.ok(measured.fits > measured.maxChars, 'the estimated budget is on the safe side of what fits');
  assert.equal(measured.suggested, Math.floor(measured.fits * 0.9));
  assert.equal(realTemplate.fonts && textMeasurer(realTemplate).fit('text-only', 'body', 'x'.repeat(10))?.fits, true);
  // Every text slot a deck fills is measured, and what is said to fit does fit.
  const textSlots = manifest.patterns.flatMap((pattern) => pattern.slots.filter((slot) => slot.kind === 'text' && slot.empty !== 'drop' && slot.empty !== 'keep'));
  assert.equal(budgets.length, textSlots.length);
  // Text of another language or width gives another number.
  const wide = measuredBudgets(realTemplate, { sample: 'averyverylongwordthatwrapsbadly and more' });
  assert.notEqual(wide.find((budget) => budget.pattern === 'text-only' && budget.slot === 'body')?.fits, measured.fits);

  const byBudget = sampleDeck(realTemplate);
  const byBox = sampleDeck(realTemplate, { fit: 'box', fill: 1 });
  const length = (deck: DeckSpec) => JSON.stringify(deck.slides.map((slide) => slide.slots)).length;
  assert.ok(length(byBox.deck) > length(byBudget.deck) * 1.2, 'the boxes hold more than the budgets allow');
  assert.deepEqual(renderDeck(realTemplate, byBox.deck, { assets: byBox.assets, layout: 'adaptive' }).warnings, []);
  // Without fonts, "box" has nothing to measure with and fills to the budgets.
  assert.deepEqual(sampleDeck(loadTemplate(pptx, manifest), { fit: 'box' }).deck, byBudget.deck);
});

// --- real fonts -------------------------------------------------------------------------------

test('fontkit reads font files: weights, kerning, Hebrew, and a stand-in for a missing weight or glyph', () => {
  const files = openFonts();
  const metrics = fontkitMetrics(fontkit, files);
  const regular = metrics.face('Calibri', { bold: false, italic: false });
  const bold = metrics.face('calibri', { bold: true, italic: false }); // names are matched whatever their case
  assert.ok(regular && bold && !regular.substitute && !bold.substitute);
  assert.ok(ems(bold, 'Hamburg', false) > ems(regular, 'Hamburg', false));
  assert.equal(regular.advances('Hamburg', false).length, 7); // one advance for each glyph
  // "AV" is a kerned pair: closer together with kerning than without.
  assert.ok(ems(regular, 'AVAVAV', true) < ems(regular, 'AVAVAV', false));
  assert.equal(metrics.face('Georgia', { bold: false, italic: false }), undefined);

  // Heebo was given in two weights: an italic is measured with the upright face, and says so.
  const heebo = metrics.face('Heebo', { bold: false, italic: false });
  assert.ok(heebo && heebo.has(0x5d0) && !regular.has(0x5d0));
  const hebrew = ems(heebo, 'שלום עולם', true);
  assert.ok(hebrew > 3 && hebrew < 6, String(hebrew));
  assert.equal(metrics.face('Heebo', { bold: false, italic: true })?.substitute, true);
  // A character the typeface lacks is looked for in the other families.
  assert.ok(metrics.fallback?.(0x5d0, { bold: false, italic: false })?.has(0x5d0));

  // One regular file for a family: its bold is the same face drawn a little wider.
  const single = fontkitMetrics(fontkit, { Body: (files.Heebo as Uint8Array[])[0] as Uint8Array });
  const fake = single.face('Body', { bold: true, italic: false });
  assert.equal(fake?.substitute, true);
  assert.ok(ems(fake, 'abc', false) > ems(single.face('Body', { bold: false, italic: false }), 'abc', false));
});

test('Hebrew in a typeface without Hebrew letters is measured in one that has them, and marked approximate', () => {
  // Carlito stands in for Calibri and has no Hebrew; Heebo does.
  const real = textMeasurer(loadTemplate(pptx, manifest), fontkitMetrics(fontkit, openFonts()));
  const latin = real.fit('text-and-photo', 'body', 'Plain text in the Latin alphabet.');
  const hebrew = real.fit('text-and-photo', 'body', 'טקסט פשוט בעברית, עם 12 מילים.');
  assert.deepEqual([latin?.approximate, latin?.lines], [false, 1]);
  assert.deepEqual([hebrew?.approximate, hebrew?.lines], [true, 1]);
});

interface Recorded {
  cases: Array<{ name: string; small: boolean; deck: DeckSpec; shapes: Record<string, Record<string, { lines: number[][]; height: number }>> }>;
}
const recorded = JSON.parse(readFileSync(join(root, 'test', 'fixtures', 'powerpoint-lines.json'), 'utf8')) as Recorded;

/** The demo template with every stated font size halved: 18 pt body text becomes 9 pt in the same boxes. */
function halved(): Uint8Array {
  const files = unzipSync(pptx);
  for (const [name, data] of Object.entries(files)) {
    if (!name.startsWith('ppt/slides/slide')) continue;
    files[name] = strToU8(strFromU8(data).replace(/ sz="([0-9]+)"/g, (_, size: string) => ` sz="${Math.round(Number(size) / 100) * 50}"`));
  }
  return zipSync(files);
}
const smallPptx = halved();

/**
 * Measures every text slot of a recorded deck and compares with what PowerPoint did: how many
 * slots have its number of lines in every paragraph, and how many break every line where it did.
 */
function compare(name: string, fonts: FontMetrics): { slots: number; sameCount: number; sameBreaks: number; heightsOff: number } {
  const entry = recorded.cases.find((item) => item.name === name);
  assert.ok(entry, `no recorded case "${name}"`);
  const { deck } = entry;
  const measurer = textMeasurer(loadTemplate(entry.small ? smallPptx : pptx, manifest), fonts);
  const mirrored = (deck.direction ?? manifest.direction) !== manifest.direction;
  const result = { slots: 0, sameCount: 0, sameBreaks: 0, heightsOff: 0 };
  deck.slides.forEach((slide, index) => {
    const pattern = manifest.patterns.find((item) => item.id === slide.pattern);
    assert.ok(pattern);
    const sources = slotSources(pattern, mirrored, slide.items);
    for (const slot of effectiveSlots(pattern, slide.items)) {
      const truth = entry.shapes[String(index + 1)]?.[slot.shapeId];
      if (slot.kind !== 'text' || !truth) continue;
      const value = slide.slots[sources.get(slot.name) ?? slot.name] ?? (slot.deckField ? deck.fields?.[slot.deckField] : undefined);
      const fit = measurer.fit(pattern, slot, value as string | string[], { textScale: bodyTextScale(manifest, deck.textScale) });
      assert.ok(fit, `${name} slide ${index + 1} ${slot.name} was not measured`);
      result.slots++;
      const counts = JSON.stringify(fit.paragraphs.map((lines) => lines.length)) === JSON.stringify(truth.lines.map((lines) => lines.length));
      const breaks = JSON.stringify(fit.paragraphs.map((lines) => lines.map((line) => line.trimEnd().length))) === JSON.stringify(truth.lines);
      if (counts) result.sameCount++;
      if (breaks) result.sameBreaks++;
      else if (process.env.SLOTDECK_VERBOSE) console.log(`${name} slide ${index + 1} ${slot.name}: ${JSON.stringify(fit.paragraphs)} against line lengths ${JSON.stringify(truth.lines)}`);
      if (counts && Math.abs(fit.height - truth.height) > 0.6) result.heightsOff++;
    }
  });
  return result;
}

test("with the template's own fonts, English and Hebrew text break exactly where PowerPoint breaks them", { skip: officeFonts() ? false : 'Calibri and Cambria are not installed here' }, () => {
  // The recorded decks fill every text slot of the demo template with running text, from half its
  // budget to well over it: at the designed sizes, with larger text, and with every size halved
  // (body text of 6 to 10 pt, where PowerPoint's rounding of glyph widths decides many breaks).
  const fonts = fontkitMetrics(fontkit, officeFonts() ?? {});
  assert.deepEqual(recorded.cases.map((item) => item.name), ['english', 'english-larger-text', 'english-small', 'hebrew', 'hebrew-larger-text', 'hebrew-small']);
  for (const { name } of recorded.cases) {
    const { slots, sameCount, sameBreaks, heightsOff } = compare(name, fonts);
    assert.equal(slots, 70, name);
    assert.equal(sameCount, slots, `${name}: ${sameCount} of ${slots} slots have PowerPoint's line counts`);
    assert.equal(sameBreaks, slots, `${name}: ${sameBreaks} of ${slots} slots break every line where PowerPoint does`);
    assert.equal(heightsOff, 0, name);
  }
});

test('with open fonts of the same metrics, English text takes the number of lines PowerPoint gives it', () => {
  // Carlito and Caladea were drawn to the widths of Calibri and Cambria. They are not the same
  // fonts, so a word now and then sits on the other side of a break; the line counts hold.
  const fonts = fontkitMetrics(fontkit, openFonts());
  for (const name of ['english', 'english-larger-text', 'english-small']) {
    const { slots, sameCount, sameBreaks, heightsOff } = compare(name, fonts);
    assert.equal(slots, 70, name);
    assert.equal(sameCount, slots, `${name}: ${sameCount} of ${slots} slots have PowerPoint's line counts`);
    assert.ok(sameBreaks >= slots - 8, `${name}: ${sameBreaks} of ${slots} slots break every line where PowerPoint does`);
    assert.equal(heightsOff, 0, name);
  }
});

test('a placeholder that takes its size, typeface, place and bullets from the master is measured as PowerPoint sets it', () => {
  // Slide 2 of the demo template, with its title and body turned into placeholders that state
  // nothing themselves: 44 pt and 32 pt from the master's text styles, the theme's typeface, the
  // master's boxes, bullets with a hanging indent, and a second-level paragraph at 28 pt.
  const files = unzipSync(pptx);
  const swap = (xml: string, id: number, shape: string): string => {
    const at = xml.indexOf(`<p:cNvPr id="${id}"`);
    return xml.slice(0, xml.lastIndexOf('<p:sp>', at)) + shape + xml.slice(xml.indexOf('</p:sp>', at) + 7);
  };
  const run = (text: string) => `<a:r><a:rPr lang="en-US"/><a:t>${text}</a:t></a:r>`;
  const placeholder = (id: number, ph: string, paragraphs: string) =>
    `<p:sp><p:nvSpPr><p:cNvPr id="${id}" name="P${id}"/><p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr><p:nvPr>${ph}</p:nvPr></p:nvSpPr><p:spPr/><p:txBody><a:bodyPr/><a:lstStyle/>${paragraphs}</p:txBody></p:sp>`;
  const phrases = ['Body text', 'that takes', 'everything from', 'the layout and', 'the master:', 'its size,', 'its bullets', 'and its margins'];
  const para = (count: number) => Array.from({ length: count }, (_, index) => phrases[index % 8]).join(' ');
  const title = 'A long title that inherits its size and its place from the master and wraps on to more lines than one';
  const body = [para(14), para(9), para(12)];

  let slide = strFromU8(files['ppt/slides/slide2.xml'] as Uint8Array);
  slide = swap(slide, 5, placeholder(5, '<p:ph type="title"/>', `<a:p>${run(title)}</a:p>`));
  slide = swap(slide, 7, placeholder(7, '<p:ph idx="1"/>', `<a:p>${run(body[0] as string)}</a:p><a:p>${run(body[1] as string)}</a:p><a:p><a:pPr lvl="1"/>${run(body[2] as string)}</a:p>`));
  files['ppt/slides/slide2.xml'] = strToU8(slide);
  // The slide now uses the "Title and Content" layout, whose placeholders point on to the master's.
  files['ppt/slides/_rels/slide2.xml.rels'] = strToU8(strFromU8(files['ppt/slides/_rels/slide2.xml.rels'] as Uint8Array).replace('slideLayout7.xml', 'slideLayout2.xml'));

  const probe: TemplateManifest = {
    version: 1,
    name: 'probe',
    direction: 'ltr',
    patterns: [{ id: 'p', sourceSlide: 2, description: 'Inherited styles', slots: [{ name: 'title', kind: 'text', shapeId: '5' }, { name: 'body', kind: 'text', shapeId: '7', maxParagraphs: 3 }] }],
  };
  const inherited = loadTemplate(zipSync(files), probe, { fonts: fontkitMetrics(fontkit, openFonts()) });
  assert.deepEqual(templateFonts(inherited), [{ typeface: 'Calibri', bold: false, italic: false }]);
  const measurer = textMeasurer(inherited);
  // The figures on the right are what PowerPoint reported for this file.
  const heading = measurer.fit('p', 'title', title);
  assert.deepEqual([heading?.sizes, heading?.paragraphs.map((lines) => lines.length), heading?.height], [[44], [3], 158.4]);
  const text = measurer.fit('p', 'body', body);
  assert.deepEqual([text?.sizes, text?.paragraphs.map((lines) => lines.length), text?.height], [[32, 32, 28], [4, 3, 4], 417.6]);
});

// --- PowerPoint's own rules, read off its layout ---------------------------------------------------

test('every glyph is rounded to an eighth of a point, as PowerPoint lays text out', () => {
  // A glyph 0.5254 em wide is 6.305 pt at 12 pt, which the grid makes 6.25: forty of them are 250 pt, not 252.2.
  const face: FontFace = { has: () => true, advances: (text) => [...text].map(() => 0.5254) };
  assert.equal(textWidth(face, 'n'.repeat(40), 12, false), 250);
  // At 13 pt the same glyph rounds up, to 6.875 pt: the text is wider than its metrics say.
  assert.equal(textWidth(face, 'n'.repeat(40), 13, false), 275);
  assert.ok(275 > 0.5254 * 13 * 40);
  // In a slot it decides a break: 49 such glyphs at 9 pt are 232.75 pt on the grid and 231.7 pt off it.
  assert.equal(textWidth(face, 'n'.repeat(49), 9, false), 232.75);
});

test('the space after a line needs room when it sits inside text that runs against the paragraph', () => {
  // An English paragraph: the space between two Hebrew words runs with them, right to left, and
  // stays inside the line. A space next to an English word runs with the paragraph and hangs.
  const hebrew = (text: string) => /[\u05d0-\u05ea]/.test(text);
  const against = (line: string, next: string) => hebrew(line.slice(-1)) && hebrew(next);
  const a = '\u05d0\u05d1\u05d2';
  const b = '\u05d3\u05d4\u05d5';
  // "abc", a Hebrew word and the space between them are seven characters. With a second Hebrew word
  // to follow, the space after the first needs an eighth place.
  assert.deepEqual(breakLines(`abc ${a} ${b}`, 7, 7, chars), [`abc ${a} `, b]);
  assert.deepEqual(breakLines(`abc ${a} ${b}`, 7, 7, chars, against), ['abc ', `${a} ${b}`]);
  assert.deepEqual(breakLines(`abc ${a} ${b}`, 8, 8, chars, against), [`abc ${a} `, b]);
  // With an English word to follow, the space hangs as usual.
  assert.deepEqual(breakLines(`abc ${a} def`, 7, 7, chars, against), [`abc ${a} `, 'def']);
  // The last word of a paragraph has no space after it, so nothing more is needed.
  assert.deepEqual(breakLines(`abc ${a}`, 7, 7, chars, against), [`abc ${a}`]);
});

test('Hebrew letters are always kerned with each other; Latin text only where the run turns kerning on', () => {
  // A font in which kerning takes an eighth of an em off every glyph, to see where it is applied.
  const kerned: FontMetrics = { face: () => ({ has: () => true, advances: (text, kerning) => [...text].map(() => (kerning ? 0.375 : 0.5)) }) };
  const width = (measurer: ReturnType<typeof textMeasurer>, slot: string, text: string) => measurer.fit('text-and-photo', slot, text)?.widths[0]?.[0];
  const normal = textMeasurer(loadTemplate(pptx, manifest), kerned);
  const small = textMeasurer(loadTemplate(smallPptx, manifest), kerned);

  // The body is 18 pt, and the template kerns from 12 pt up: Latin text is kerned, with its punctuation.
  assert.equal(width(normal, 'body', 'abc, def'), 8 * 0.375 * 18);
  // At 9 pt it is not.
  assert.equal(width(small, 'body', 'abc, def'), 8 * 0.5 * 9);
  // Hebrew letters are kerned at any size, and never with the comma or the space beside them.
  const hebrew = 'אבג, דהו';
  assert.equal(width(small, 'body', hebrew), 6 * 0.375 * 9 + 2 * 0.5 * 9);
  assert.equal(width(normal, 'body', hebrew), 6 * 0.375 * 18 + 2 * 0.5 * 18);
  // A Latin word in a Hebrew paragraph is not kerned, whatever its size.
  assert.equal(width(normal, 'body', 'אבג abc'), 3 * 0.375 * 18 + 4 * 0.5 * 18);
  // A Hebrew word in an English paragraph is.
  assert.equal(width(normal, 'body', 'abc אבג'), 4 * 0.375 * 18 + 3 * 0.375 * 18);
});
