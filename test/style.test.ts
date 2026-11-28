// A manifest made from a pptx alone (`styleManifest`, a prototype): enough of one for the template
// to carry a cover, a closing slide and the slides the software lays out itself.
import assert from 'node:assert/strict';
import { readFileSync } from 'node:fs';
import { dirname, join } from 'node:path';
import { test } from 'node:test';
import { fileURLToPath } from 'node:url';
import { strFromU8, strToU8, unzipSync, zipSync } from 'fflate';
import { pageColor } from '../src/color/recolor.ts';
import { freeKitOrReason } from '../src/free/kit.ts';
import { PptxPackage } from '../src/pptx/package.ts';
import { parseXml, rootElement } from '../src/pptx/xml.ts';
import { checkManifest, freeSlide, loadTemplate, renderDeck, validateDeck } from '../src/index.ts';
import type { DeckSpec } from '../src/index.ts';
import { styleManifest } from '../src/template/style.ts';

const root = join(dirname(fileURLToPath(import.meta.url)), '..');
const demoPptx = new Uint8Array(readFileSync(join(root, 'templates', 'demo.pptx')));

test('a manifest is made from the pptx alone, the same every time, and passes the check', () => {
  const style = styleManifest(demoPptx, { name: 'demo' });
  assert.deepEqual(styleManifest(demoPptx, { name: 'demo' }), style, 'the same file gives the same manifest');
  const { manifest } = style;
  assert.equal(manifest.name, 'demo');
  assert.equal(manifest.direction, 'ltr');
  assert.equal(manifest.language, 'en');
  // Every slide of the demo is read, and its first and last are the cover and the closing slide.
  assert.equal(manifest.patterns.length, 13);
  assert.equal(style.cover, 'cover');
  assert.equal(style.closing, 'closing');
  assert.equal(manifest.patterns[0]?.id, 'cover');
  assert.equal(manifest.patterns.at(-1)?.id, 'closing');
  for (const id of ['cover', 'closing']) assert.ok(manifest.patterns.find((pattern) => pattern.id === id)?.slots.some((slot) => slot.name === 'title'), `${id} has a title`);
  // The line of the organisation that every slide of the demo shows is the deck's own text: a field.
  assert.deepEqual(Object.values(style.fields), ['Organization name | Department']);
  assert.ok(manifest.patterns.some((pattern) => pattern.slots.some((slot) => slot.kind === 'text' && slot.deckField === 'field1')));
  // No slot is named twice on a slide, and none points at a shape the slide does not have.
  assert.deepEqual(checkManifest(loadTemplate(demoPptx, manifest)).filter((issue) => issue.severity === 'error'), []);
  for (const pattern of manifest.patterns) assert.equal(new Set(pattern.slots.map((slot) => slot.name)).size, pattern.slots.length, pattern.id);
});

test('the template\'s style is found without anyone naming a slot: its frame, its text styles, its colours', () => {
  const style = styleManifest(demoPptx);
  const template = loadTemplate(demoPptx, style.manifest);
  const kit = freeKitOrReason(template);
  assert.ok(typeof kit !== 'string', String(kit));
  if (typeof kit === 'string') return;
  // The frame is a slide whose title sits where most titles sit, with room under it and nothing drawn there.
  assert.equal(kit.base, (style.manifest.free as { base?: string }).base);
  assert.ok(kit.band.bottom - kit.band.top >= 0.45);
  assert.match(style.notes.join('\n'), /free slides stand on the frame of slide \d+: its title sits where the titles of \d+ slides sit, and nothing is drawn under it/);
  // The styles are the demo's own: running text at 18 pt, a bold heading over it, a big number in the main colour.
  assert.equal(kit.styles.body.size, 18);
  assert.ok(kit.styles.heading.bold && kit.styles.heading.size >= kit.styles.body.size && kit.styles.heading.set === undefined, 'a heading of the template, not one made up');
  assert.ok(kit.styles.figure.size >= 36 && kit.styles.figure.set === undefined, 'a big number of the template');
  assert.equal(kit.styles.figure.color, '0E7C7B');
  assert.equal(kit.palette.accent, '0E7C7B');
  assert.ok(kit.panel, 'the box its cards stand on');
});

test('a deck of the cover, slides the software lays out and the closing slide renders on it', () => {
  const style = styleManifest(demoPptx);
  const template = loadTemplate(demoPptx, style.manifest);
  const deck: DeckSpec = {
    version: 1,
    slides: [
      { pattern: 'cover', slots: { title: 'Riverbend 2027', subtitle: 'The plan for the city council' } },
      freeSlide({ title: 'A year in three numbers', tone: 'accent', blocks: [{ figure: { value: '34%', label: 'more passes' } }, { figure: { value: '11', label: 'new stations' } }] }),
      freeSlide({ title: 'How a station opens', blocks: [{ phase: { heading: 'Survey', text: 'Two weeks.' } }, { phase: { heading: 'Build', text: 'Three days.' } }, { phase: { heading: 'Launch' } }] }),
      { pattern: 'closing', slots: { title: 'Thank you' } },
    ],
  };
  assert.deepEqual(validateDeck(template, deck, { layout: 'adaptive' }).filter((issue) => issue.severity === 'error'), []);
  const files = unzipSync(renderDeck(template, deck, { layout: 'adaptive' }).pptx);
  const slide = (number: number) => strFromU8(files[`ppt/slides/slide${number}.xml`] as Uint8Array);
  assert.match(slide(1), /Riverbend 2027/);
  assert.match(slide(2), /name="Ground"/);
  assert.match(slide(3), /prst="homePlate"/);
  assert.match(slide(4), /Thank you/);
  // The deck gave the frame's own line no text, so no slide shows the template's sample of it.
  for (const number of [1, 2, 3, 4]) assert.doesNotMatch(slide(number), /Organization name/);
});

// A page of the template's colour in one corner and white over the rest: the look of many brand templates.
const page = (stops: Array<[number, string]>): string =>
  `<p:bg><p:bgPr><a:gradFill><a:gsLst>${stops.map(([pos, hex]) => `<a:gs pos="${pos}"><a:srgbClr val="${hex}"/></a:gs>`).join('')}</a:gsLst><a:path path="circle"><a:fillToRect l="100000" t="100000"/></a:path></a:gradFill><a:effectLst/></p:bgPr></p:bg>`;
const GLOW = page([[0, '7CF29C'], [30000, 'FFFFFF'], [100000, 'FFFFFF']]);
const BLEND = page([[0, '0E7C7B'], [100000, 'FFFFFF']]);
const withPage = (part: string, xml: string): Uint8Array => {
  const zip = unzipSync(demoPptx);
  zip[part] = strToU8(strFromU8(zip[part] as Uint8Array).replace('<p:spTree>', `${xml}<p:spTree>`));
  return zipSync(zip);
};
const pageOf = (pptx: Uint8Array, slide: number) => {
  const pkg = PptxPackage.load(pptx);
  const part = `ppt/slides/slide${slide}.xml`;
  return pageColor(pkg, part, rootElement(parseXml(pkg.readText(part))));
};

test('a page that is one colour but for a glow in a corner is read by that colour', () => {
  // Green at the corner, white from three tenths of the way on: white stands for 85% of the page.
  const glow = pageOf(withPage('ppt/slides/slide7.xml', GLOW), 7);
  assert.equal(glow.color, 'FFFFFF');
  assert.ok(Math.abs((glow.glow ?? 0) - 0.85) < 1e-9, String(glow.glow));
  // A blend of two colours is no page of one colour: it is read by its first, as it always was.
  assert.deepEqual(pageOf(withPage('ppt/slides/slide7.xml', BLEND), 7), { color: '0E7C7B' });
  // A page of one colour has no glow, and neither has a slide that states no page of its own.
  assert.deepEqual(pageOf(withPage('ppt/slides/slide7.xml', '<p:bg><p:bgPr><a:solidFill><a:srgbClr val="0E7C7B"/></a:solidFill><a:effectLst/></p:bgPr></p:bg>'), 7), { color: '0E7C7B' });
  assert.equal(pageOf(demoPptx, 7).glow, undefined);
  // The page a slide takes from its layout is read the same way.
  const layout = /slideLayouts\/(slideLayout\d+\.xml)/.exec(strFromU8(unzipSync(demoPptx)['ppt/slides/_rels/slide7.xml.rels'] as Uint8Array))?.[1];
  const inherited = pageOf(withPage(`ppt/slideLayouts/${layout}`, GLOW), 7);
  assert.equal(inherited.color, 'FFFFFF');
  assert.ok(Math.abs((inherited.glow ?? 0) - 0.85) < 1e-9);
});

test('free slides on a page with a glow choose their text for the page, and such a frame is taken before a plain one', () => {
  const plain = styleManifest(demoPptx);
  assert.equal((plain.manifest.free as { base?: string }).base, 'slide4');
  // Slide 7 is a frame as clear as slide 4: with the template's glow on its page, it is the one taken.
  const pptx = withPage('ppt/slides/slide7.xml', GLOW);
  const style = styleManifest(pptx);
  assert.equal((style.manifest.free as { base?: string }).base, 'slide7');
  assert.match(style.notes.join('\n'), /free slides stand on the frame of slide 7: .*its page is one colour over 85% of it, with the template's glow in the rest/);
  // A page that is a blend is not such a page: the first clear frame is taken, as before.
  assert.equal((styleManifest(withPage('ppt/slides/slide7.xml', BLEND)).manifest.free as { base?: string }).base, 'slide4');
  // The text of a free slide is chosen for the white the content stands on: the same styles as on the plain slide.
  const template = loadTemplate(pptx, style.manifest);
  const [kit, same] = [freeKitOrReason(template), freeKitOrReason(loadTemplate(demoPptx, { ...plain.manifest, free: { base: 'slide7' } }))];
  assert.ok(typeof kit !== 'string' && typeof same !== 'string', `${String(kit)} ${String(same)}`);
  if (typeof kit === 'string' || typeof same === 'string') return;
  assert.equal(kit.background, 'FFFFFF');
  assert.deepEqual(kit.styles, same.styles);
  assert.equal(kit.palette.accent, same.palette.accent);
  // And the slide that is made keeps the glow: it is the frame's own page.
  const deck: DeckSpec = { version: 1, slides: [freeSlide({ title: 'The road to 2027', blocks: [{ milestone: { when: 'Q1 2026', heading: 'North Bridge' } }, { milestone: { when: 'Q3 2026', heading: 'E-bikes arrive' } }] })] };
  assert.deepEqual(validateDeck(template, deck, { layout: 'adaptive' }).filter((issue) => issue.severity === 'error'), []);
  const made = strFromU8(unzipSync(renderDeck(template, deck, { layout: 'adaptive' }).pptx)['ppt/slides/slide1.xml'] as Uint8Array);
  assert.match(made, /<p:bg><p:bgPr><a:gradFill>/);
  assert.match(made, /7CF29C/);
});
