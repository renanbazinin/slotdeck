// On a free slide, the spacing of running text and lists gives way before its size (0.6.9): a text
// that fits at the template's own spacing, and that closer spacing sets larger, is set larger, with
// the space between its paragraphs closed up to no less than 0.7 of it, and lines set wider than
// single by up to a tenth, never closer than single. What is cut, reported or set whole is decided
// at the template's own spacing, as 0.6.8 decided it.
import assert from 'node:assert/strict';
import { readFileSync } from 'node:fs';
import { dirname, join } from 'node:path';
import { test } from 'node:test';
import { fileURLToPath } from 'node:url';
import { strFromU8, strToU8, unzipSync, zipSync } from 'fflate';
import * as fontkit from 'fontkit';
import { fontkitMetrics, loadTemplate, renderDeck, textMeasurer, validateDeck } from '../src/index.ts';
import type { DeckSpec, SlideSpec, TemplateManifest } from '../src/index.ts';
import { freeKit } from '../src/free/kit.ts';
import { freeFor } from '../src/free/resolve.ts';
import type { FreeTextSlot } from '../src/free/plan.ts';
import type { FreeResolved } from '../src/free/resolve.ts';
import { openFonts } from './fonts.ts';

const root = join(dirname(fileURLToPath(import.meta.url)), '..');
const demoManifest = JSON.parse(readFileSync(join(root, 'templates', 'demo.manifest.json'), 'utf8')) as TemplateManifest;
const demoPptx = new Uint8Array(readFileSync(join(root, 'templates', 'demo.pptx')));
const demo = loadTemplate(demoPptx, demoManifest);
/** The demo with its running text's paragraphs given these properties first (its body style is the text-only slide's body). */
function withBody(insert: string): Uint8Array {
  const files = unzipSync(new Uint8Array(demoPptx));
  const name = 'ppt/slides/slide2.xml';
  const xml = strFromU8(files[name] as Uint8Array);
  const start = xml.indexOf('<p:cNvPr id="7" name="Body"/>');
  const end = xml.indexOf('</p:sp>', start);
  assert.ok(start > 0 && end > start);
  files[name] = strToU8(xml.slice(0, start) + xml.slice(start, end).replaceAll('<a:pPr algn="l">', `<a:pPr algn="l">${insert}`) + xml.slice(end));
  return zipSync(files);
}
/** The demo with its running text's lines set at 112%, as a right-to-left template sets them. */
const spacedPptx = withBody('<a:lnSpc><a:spcPct val="112000"/></a:lnSpc>');
const spaced = loadTemplate(spacedPptx, demoManifest);

const SENTENCE = 'Riders rely on the service every day, and the city gains from it in ways that are easy to see. ';
const textOf = (chars: number, count: number): string[] =>
  Array.from({ length: count }, (_, at) => {
    const length = Math.floor(chars / count) + (at < chars % count ? 1 : 0);
    return SENTENCE.repeat(Math.ceil(length / SENTENCE.length) + 1).slice(0, length - 1).trim() + '.';
  });
const slideOf = (text: string[], dense = false): SlideSpec => ({ pattern: 'free', blocks: ['text'], ...(dense ? { arrange: { dense: true } } : {}), slots: { title: 'A title', b1Text: text } });
const deckOf = (slides: SlideSpec[]): DeckSpec => ({ version: 1, language: 'en', fields: { organizationDepartment: 'Retail' }, slides });
const resolved = (value: FreeResolved | string): FreeResolved => {
  assert.notEqual(typeof value, 'string', String(value));
  return value as FreeResolved;
};
/** The text of a slide as planned: its size, the space between its paragraphs, and the line spacing its copy is drawn with where that is not its source's. */
const textOf1 = (template: typeof demo, slide: SlideSpec) => {
  const { plan } = resolved(freeFor(template, {}, slide));
  const slot = plan.pattern.slots.find((entry) => entry.name === 'b1Text');
  const shape = plan.shapes.find((entry) => entry.id === slot?.shapeId);
  assert.ok(slot?.kind === 'text' && shape?.kind === 'text');
  return { size: slot.fontSize ?? 0, gap: slot.paragraphGap ?? 0, lines: shape.lineSpacing, columns: (slot as FreeTextSlot).drawn.columns ?? 1 };
};
/** The paragraph properties of the slide's running text as rendered. */
const drawnSpacing = (template: typeof demo, slide: SlideSpec): string => {
  const xml = strFromU8(unzipSync(renderDeck(template, deckOf([slide]), { layout: 'adaptive' }).pptx)['ppt/slides/slide1.xml'] as Uint8Array);
  return /<p:sp>(?:(?!<\/p:sp>)[\s\S])*?Riders rely(?:(?!<\/p:sp>)[\s\S])*<\/p:sp>/.exec(xml)?.[0] ?? '';
};
const measuring = loadTemplate(spacedPptx, demoManifest, { fonts: fontkitMetrics(fontkit, openFonts()) });

test('a text a little too long keeps its size: the space between its paragraphs, and lines set wider than single, give way first', () => {
  // Six paragraphs of 600 characters: 0.6.8 set them at 16.5 pt, with the usual space of 8 between
  // them and lines 112% apart. Closed up to 7 and 101%, they fit at 19 pt.
  const slide = slideOf(textOf(600, 6));
  const set = textOf1(spaced, slide);
  assert.deepEqual(set, { size: 19, gap: 7, lines: { percent: 1.01 }, columns: 1 });
  // The space between paragraphs is no less than 0.7 of half the size, and lines no closer than single.
  assert.ok(set.gap >= Math.ceil(0.7 * Math.round(0.5 * set.size)));
  // Drawn as counted: whole points after each paragraph, and whole percents first among a paragraph's properties.
  const drawn = drawnSpacing(spaced, slide);
  assert.equal((drawn.match(/<a:pPr[^>]*><a:lnSpc><a:spcPct val="101000"\/><\/a:lnSpc>/g) ?? []).length, 6, drawn.slice(0, 600));
  assert.equal((drawn.match(/<a:spcAft><a:spcPts val="700"\/><\/a:spcAft>/g) ?? []).length, 6);
  // The deck's check finds it fits, by its length and measured with the template's fonts at the spacing drawn.
  assert.deepEqual(validateDeck(spaced, deckOf([slide]), { layout: 'adaptive' }), []);
  assert.deepEqual(validateDeck(measuring, deckOf([slide]), { layout: 'adaptive' }), []);
});

test('lines set at single give nothing: on such a template only the space between paragraphs gives way, and no line spacing is written', () => {
  // 0.6.8 set them at 18.5 pt with a space of 9: closed up to 7, at 19.5.
  const slide = slideOf(textOf(600, 6));
  assert.deepEqual(textOf1(demo, slide), { size: 19.5, gap: 7, lines: undefined, columns: 1 });
  assert.ok(!/<a:lnSpc>/.test(drawnSpacing(demo, slide)));
  // Dense text in two columns: 0.6.8 set it at 15.5 pt with a space of 8; closed up to 6, at 16.
  const dense = slideOf(textOf(1000, 6), true);
  assert.deepEqual(textOf1(demo, dense), { size: 16, gap: 6, lines: undefined, columns: 2 });
  assert.deepEqual(validateDeck(demo, deckOf([dense]), { layout: 'adaptive' }), []);
});

test('a text that does not fit at the template’s own spacing is cut and reported as before: its spacing does not give way', () => {
  for (const template of [demo, spaced]) {
    // Far too long to set whole at the least size: set at the template's spacing, and reported.
    const slide = slideOf(textOf(9000, 12), true);
    const set = textOf1(template, slide);
    assert.equal(set.gap, Math.round(0.5 * set.size));
    assert.equal(set.lines, undefined);
    assert.ok(validateDeck(template, deckOf([slide]), { layout: 'adaptive' }).some((issue) => /characters; the slot fits about/.test(issue.message)));
  }
});

test('a text of many short paragraphs is not set down the slide past its box', () => {
  // Twelve one-line paragraphs, 337 characters. Counted by their length they come to a few lines, and
  // 0.6.8 set them so far down the slide that the last two ran 38.5 pt under the box in PowerPoint
  // (the deck's check said so). Now no text is set further down than its paragraphs leave free.
  const lines = ['Riders wait less at stops.', 'The night line runs later.', 'Two depots share a workshop.', 'Ticket sales rose by a fifth.', 'Trains are full at eight.', 'New drivers train with mentors.', 'Bike racks fill up by nine.', 'The board meets in March.', 'Four stations still need lifts.', 'Student fares stay the same.', 'Repairs close the bridge line.', 'The city pays for new shelters.'];
  for (const template of [demo, spaced]) {
    const slide = slideOf(lines);
    const { plan } = resolved(freeFor(template, {}, slide));
    const slot = plan.pattern.slots.find((entry) => entry.name === 'b1Text') as FreeTextSlot;
    const size = slot.fontSize ?? 0;
    const drawn = plan.shapes.find((entry) => entry.id === slot.shapeId && entry.kind === 'text')?.lineSpacing;
    const line = drawn && 'points' in drawn ? drawn.points : 1.2 * size * (drawn?.percent ?? (template === spaced ? 1.12 : 1));
    const height = ((slot.box?.h ?? 0) * (freeKit(template)?.slideHeight ?? 0)) / 12700;
    assert.ok(12 * line + 11 * (slot.paragraphGap ?? 0) <= height + Math.min(3, 0.25 * size) + 0.1, JSON.stringify([size, slot.paragraphGap, height]));
    assert.deepEqual(validateDeck(template, deckOf([slide]), { layout: 'adaptive' }), []);
  }
});

test('a text set down its room past its box in 0.6.8 is set higher, as far as its paragraphs need, its spacing its own', () => {
  // Eight short paragraphs at 22 pt: 0.6.8 set them so far down that the check reported them. They
  // are set higher, at the same size and spacing, and fit.
  const slide = slideOf(Array.from({ length: 8 }, (_, at) => `Point ${at + 1}.`.padEnd(12, 'x')));
  assert.deepEqual(textOf1(demo, slide), { size: 22, gap: 11, lines: undefined, columns: 1 });
  assert.deepEqual(validateDeck(demo, deckOf([slide]), { layout: 'adaptive' }), []);
});

test('like texts share one spacing, closed up together', () => {
  const slide: SlideSpec = { pattern: 'free', blocks: ['text', 'text'], slots: { title: 'A title', b1Text: textOf(600, 4), b2Text: textOf(600, 4) } };
  const { plan } = resolved(freeFor(demo, {}, slide));
  const texts = plan.pattern.slots.filter((slot) => slot.name === 'b1Text' || slot.name === 'b2Text').map((slot) => [slot.kind === 'text' ? slot.fontSize : 0, slot.kind === 'text' ? slot.paragraphGap : 0]);
  assert.equal(texts.length, 2);
  assert.deepEqual(texts[0], texts[1]);
  const [size, gap] = texts[0] as [number, number];
  assert.ok(gap < Math.round(0.5 * size), JSON.stringify(texts));
});

test('a long text whose lines close up may take fewer columns, never more', () => {
  // Six paragraphs of 1,600 characters set dense on a template whose lines are 112%: 0.6.8 set them
  // at 14.5 pt in three columns. Closed up, two hold them at 15.
  assert.deepEqual(textOf1(spaced, slideOf(textOf(1600, 6), true)), { size: 15, gap: 6, lines: { percent: 1.01 }, columns: 2 });
});

test('a deck that asks for larger text takes as much of it as 0.6.8 gave, and closer spacing only sets a text larger', () => {
  const sizes = (scale: number, count: number, chars: number) => {
    const slide: SlideSpec = { pattern: 'free', blocks: ['text', 'note'], slots: { title: 'A title', b1Text: textOf(chars, count), b2Text: SENTENCE.slice(0, 80) } };
    const free = resolved(freeFor(demo, { textScale: scale }, slide));
    const of = (name: string) => {
      const slot = free.plan.pattern.slots.find((entry) => entry.name === name);
      return slot?.kind === 'text' ? [slot.fontSize, slot.paragraphGap] : undefined;
    };
    return { text: of('b1Text'), note: of('b2Text')?.[0], enlarged: free.enlarged ?? 'all' };
  };
  // 0.6.8: all of the wish, the text at 18.5 pt with a space of 9. Its spacing closed up, 20 pt.
  assert.deepEqual(sizes(1.15, 8, 150), { text: [20, 7], note: 19.5, enlarged: 'all' });
  // Where 0.6.8 gave half of it, or none, so does this; no text is smaller than it set it.
  assert.deepEqual(sizes(1.25, 6, 250), { text: [22, 11], note: 19, enlarged: 'half' });
  assert.deepEqual(sizes(1.5, 6, 250), { text: [22, 11], note: 17, enlarged: 'none' });
});

test('a long text that fits only with its spacing closed up is reported as 0.6.8 reported it, and drawn at its own spacing', async () => {
  const { buildFixturePptx, fixtureManifest } = await import('./fixture.ts');
  const rtl = loadTemplate(buildFixturePptx(), fixtureManifest);
  const HE = 'הרוכבים סומכים על השירות בכל יום, והעיר מרוויחה ממנו בדרכים שקל לראות אותן היטב. ';
  const text = Array.from({ length: 32 }, (_, at) => {
    const length = Math.floor(940 / 32) + (at < 940 % 32 ? 1 : 0);
    return HE.repeat(3).slice(at % 9, at % 9 + length - 1).trim() + '.';
  });
  const slide = slideOf(text, true);
  const set = textOf1(rtl, slide);
  assert.equal(set.gap, Math.round(0.5 * set.size), JSON.stringify(set));
  assert.ok(validateDeck(rtl, { ...deckOf([slide]), language: 'he', direction: 'rtl', fields: { closing: 'Thanks' } }, { layout: 'adaptive' }).some((issue) => /characters; the slot fits about|paragraphs; the slot fits/.test(issue.message)));
});

test('the gate, the loosest spacing, no more columns, lines in points: each as counted', () => {
  const headed = (text: string[], dense = false): SlideSpec => ({ ...slideOf(text, dense), slots: { title: 'A title', b1Heading: 'A heading here', b1Text: text } });
  // Over its budget at the template's spacing, as in 0.6.8: reported, and drawn at that spacing,
  // though closed up it would fit at 15.5 pt.
  const over = slideOf(textOf(1120, 10));
  assert.deepEqual(textOf1(spaced, over), { size: 14.5, gap: 7, lines: undefined, columns: 2 });
  assert.equal(validateDeck(spaced, deckOf([over]), { layout: 'adaptive' }).length, 1);
  // The loosest spacing that fits: 0.6.8 set it at 15 pt with a space of 8; at 15.5 a space of 7 fits, so not 6.
  assert.deepEqual(textOf1(demo, slideOf(textOf(940, 10))), { size: 15.5, gap: 7, lines: undefined, columns: 2 });
  // Closer spacing does not move a text into more columns: three would hold it at 12.5 pt with a space of 5.
  assert.deepEqual(textOf1(demo, headed(textOf(2220, 12), true)), { size: 12, gap: 6, lines: undefined, columns: 2 });
  // A template that states its line spacing in points at about single has no line to give, and is counted no higher for it: 0.6.8 set this at 16.5 pt.
  const points = loadTemplate(withBody('<a:lnSpc><a:spcPts val="2600"/></a:lnSpc>'), demoManifest);
  assert.deepEqual(textOf1(points, slideOf(textOf(340, 10), true)), { size: 22, gap: 8, lines: undefined, columns: 1 });
  // A short text 0.6.8 set down past its box is set only as much higher as its last line needs: at 0.4105 of the slide, where 0.6.8 set it at 0.4308.
  const short = headed(textOf(200, 8));
  const { plan } = resolved(freeFor(spaced, {}, short));
  assert.equal(Number(((plan.pattern.slots.find((slot) => slot.name === 'b1Text') as FreeTextSlot).box?.y ?? 0).toFixed(4)), 0.4105);
  assert.deepEqual(validateDeck(spaced, deckOf([short]), { layout: 'adaptive' }), []);
});

test('measured with the template’s fonts, a text whose lines closed up is measured at the spacing drawn', () => {
  const slide = slideOf(textOf(600, 6));
  const { plan } = resolved(freeFor(measuring, {}, slide));
  const slot = plan.pattern.slots.find((entry) => entry.name === 'b1Text') as FreeTextSlot;
  const fit = textMeasurer(measuring).fit(plan.pattern, slot, slide.slots.b1Text as string[]);
  assert.ok(fit?.fits);
  // At 101%, not the template's 112%.
  assert.ok((fit?.height ?? Infinity) <= (fit?.lines ?? 0) * 1.2 * 19 * 1.01 + 5 * 7 + 0.5, JSON.stringify([fit?.lines, fit?.height]));
});

test('a slide held at the sizes it has without the wish keeps them where spacing gives way', () => {
  // A list of three points beside a text, in a deck that asks for text 1.15 times larger: 0.6.8
  // held the text at 22.5 pt, its size without the wish. Planned again with spacing that gives
  // way, it is no smaller than that.
  const slide: SlideSpec = { pattern: 'free', blocks: ['list', 'text'], slots: { title: 'A title', b1Items: ['Riders rely on the service every day, a.', 'iders rely on the service every day, an.', 'ders rely on the service every day, and.'], b2Text: [SENTENCE.repeat(4).slice(0, 299).trim() + '.'] } };
  const { plan } = resolved(freeFor(demo, { textScale: 1.15 }, slide));
  const text = plan.pattern.slots.find((entry) => entry.name === 'b2Text');
  assert.ok(text?.kind === 'text' && (text.fontSize ?? 0) >= 22.5, String(text?.kind === 'text' && text.fontSize));
});
